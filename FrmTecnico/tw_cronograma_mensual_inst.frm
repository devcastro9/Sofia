VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_cronograma_mensual_inst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Instalaciones - Cronograma por Grupo Piloto"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14400
   Icon            =   "tw_cronograma_mensual_inst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraInsumos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una Opción para Actualizar INSUMOS ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2400
      Left            =   5400
      TabIndex        =   159
      Top             =   4080
      Visible         =   0   'False
      Width           =   7980
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Para programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC) los insumos 3 y 4."
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
         Left            =   240
         TabIndex        =   164
         Top             =   1080
         Width           =   7335
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Para programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV) los insumos 3 y 4."
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
         Left            =   240
         TabIndex        =   163
         Top             =   600
         Width           =   7575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   7860
         TabIndex        =   160
         Top             =   1680
         Width           =   7860
         Begin VB.PictureBox BtnCancelar8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4320
            Picture         =   "tw_cronograma_mensual_inst.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   162
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGrabar8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1920
            Picture         =   "tw_cronograma_mensual_inst.frx":12EE
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   161
            Top             =   0
            Width           =   1280
         End
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000040&
      Height          =   3375
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   13260
      Begin VB.PictureBox FraGrabarCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   12960
         TabIndex        =   153
         Top             =   2640
         Visible         =   0   'False
         Width           =   12960
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6195
            Picture         =   "tw_cronograma_mensual_inst.frx":1ADC
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   155
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4800
            Picture         =   "tw_cronograma_mensual_inst.frx":23C8
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   154
            Top             =   0
            Width           =   1300
         End
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
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
            Left            =   14175
            TabIndex        =   156
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "fmes_nro_horarios_hab"
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
         Height          =   290
         Left            =   11440
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   51
         Top             =   1320
         Width           =   1410
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8990
         TabIndex        =   50
         Top             =   2095
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8990
         TabIndex        =   49
         Top             =   1570
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8290
         TabIndex        =   48
         Top             =   1040
         Width           =   255
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "observaciones"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "tw_cronograma_mensual_inst.frx":2B9E
         Top             =   2640
         Visible         =   0   'False
         Width           =   10320
      End
      Begin VB.TextBox txt_codigo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ges_gestion"
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
         Height          =   300
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   520
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7200
         TabIndex        =   12
         Top             =   1575
         Width           =   270
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "fmes_nro_hrs_habiles"
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
         Height          =   290
         Left            =   11445
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   555
         Width           =   1410
      End
      Begin VB.TextBox Txt_estado 
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
         Height          =   315
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2085
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7220
         TabIndex        =   5
         Top             =   1035
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2BA0
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   1020
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2BB9
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2BD2
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2BEB
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1020
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2C04
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   26
         Top             =   2085
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_cronograma_mensual_inst.frx":2C1D
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   27
         Top             =   2085
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "0"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "fmes_fecha_registro"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   9840
         TabIndex        =   64
         Top             =   2085
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   44600
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label lbl_texto3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1680
         TabIndex        =   113
         Top             =   525
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horarios Hábiles X Mes"
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
         Height          =   195
         Index           =   8
         Left            =   10875
         TabIndex        =   52
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_nro_dias_habiles"
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
         Left            =   9120
         TabIndex        =   47
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Hábiles X Mes"
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
         Height          =   195
         Index           =   4
         Left            =   8925
         TabIndex        =   46
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2650
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Zona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   2095
         Width           =   1605
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora"
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
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1570
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas Hábiles X Mes"
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
         Height          =   195
         Index           =   7
         Left            =   11025
         TabIndex        =   30
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Height          =   195
         Index           =   6
         Left            =   12240
         TabIndex        =   29
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Elaboracion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9840
         TabIndex        =   25
         Top             =   1845
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_nro_dias"
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
         Left            =   6975
         TabIndex        =   24
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dias X Mes"
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
         Height          =   195
         Index           =   3
         Left            =   6915
         TabIndex        =   23
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Crono."
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
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   22
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion"
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
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   280
         Width           =   660
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo Piloto"
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
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1035
         Width           =   1065
      End
      Begin VB.Label lbl_texto1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "fmes_correl"
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
         Height          =   300
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_plan"
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
         Left            =   4800
         TabIndex        =   16
         Top             =   525
         Width           =   1335
      End
   End
   Begin VB.Frame FraDet7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija las Fechas para: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2160
      Left            =   6360
      TabIndex        =   128
      Top             =   6720
      Visible         =   0   'False
      Width           =   5580
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   70
         ScaleHeight     =   660
         ScaleWidth      =   5445
         TabIndex        =   138
         Top             =   1440
         Width           =   5450
         Begin VB.PictureBox BtnCancelar7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "tw_cronograma_mensual_inst.frx":2C36
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   140
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGrabar7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1200
            Picture         =   "tw_cronograma_mensual_inst.frx":3522
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   139
            Top             =   0
            Width           =   1280
         End
      End
      Begin MSDataListLib.DataCombo DTPfecha4 
         Bindings        =   "tw_cronograma_mensual_inst.frx":3D10
         Height          =   315
         Left            =   600
         TabIndex        =   129
         Top             =   720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DTPfecha5 
         Bindings        =   "tw_cronograma_mensual_inst.frx":3D2A
         Height          =   315
         Left            =   3480
         TabIndex        =   130
         Top             =   720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha inicial (desde...)"
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
         Height          =   195
         Left            =   360
         TabIndex        =   132
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final (hasta...)"
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
         Height          =   195
         Left            =   3360
         TabIndex        =   131
         Top             =   480
         Width           =   1830
      End
   End
   Begin VB.Frame FraDet6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una Opción para Enviar Registros..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2400
      Left            =   6240
      TabIndex        =   123
      Top             =   2760
      Visible         =   0   'False
      Width           =   6900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   6780
         TabIndex        =   141
         Top             =   1680
         Width           =   6780
         Begin VB.PictureBox BtnGrabar6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1920
            Picture         =   "tw_cronograma_mensual_inst.frx":3D44
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   152
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "tw_cronograma_mensual_inst.frx":4532
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   142
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3. Envía SOLO los equipos PENDIENTES del Origen al Destino"
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
         Left            =   480
         TabIndex        =   126
         Top             =   1200
         Width           =   6135
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Envía TODO, a todos los días calendario, incluyendo días NO laborales"
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
         Left            =   480
         TabIndex        =   125
         Top             =   840
         Width           =   6375
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Envía TODO, solo a los Horarios Laborales definidos en el Destino"
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
         Left            =   480
         TabIndex        =   124
         Top             =   480
         Width           =   6015
      End
   End
   Begin VB.Frame FraDet5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija los parámetros para retornar al Crono. Origen..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2760
      Left            =   6360
      TabIndex        =   116
      Top             =   3120
      Visible         =   0   'False
      Width           =   6300
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   6180
         TabIndex        =   143
         Top             =   2040
         Width           =   6180
         Begin VB.PictureBox BtnGrabar5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "tw_cronograma_mensual_inst.frx":4E1E
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   151
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3000
            Picture         =   "tw_cronograma_mensual_inst.frx":560C
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   144
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "tw_cronograma_mensual_inst.frx":5EF8
         Height          =   315
         Left            =   240
         TabIndex        =   117
         Top             =   675
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_descripcion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DTPfecha2 
         Bindings        =   "tw_cronograma_mensual_inst.frx":5F11
         Height          =   315
         Left            =   840
         TabIndex        =   118
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DTPfecha3 
         Bindings        =   "tw_cronograma_mensual_inst.frx":5F2B
         Height          =   315
         Left            =   3720
         TabIndex        =   122
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final (hasta...)"
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
         Height          =   195
         Left            =   3600
         TabIndex        =   121
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha inicial (desde...)"
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
         Height          =   195
         Left            =   600
         TabIndex        =   120
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio..."
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
         Height          =   195
         Left            =   240
         TabIndex        =   119
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija para cambiar el Número de horas de Servicio ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1680
      Left            =   6360
      TabIndex        =   59
      Top             =   7440
      Visible         =   0   'False
      Width           =   4860
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H80000015&
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   2760
         Picture         =   "tw_cronograma_mensual_inst.frx":5F45
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancela sin Guardar"
         Top             =   840
         Width           =   1125
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H80000015&
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   960
         Picture         =   "tw_cronograma_mensual_inst.frx":614F
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Grabar los Datos"
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox txtnrohrs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   61
         Text            =   "tw_cronograma_mensual_inst.frx":6359
         Top             =   360
         Width           =   645
      End
      Begin VB.ComboBox cmd_campo2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_cronograma_mensual_inst.frx":635B
         Left            =   3960
         List            =   "tw_cronograma_mensual_inst.frx":636B
         TabIndex        =   60
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbl_orden 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.de Horas actual"
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
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   375
         Width           =   1725
      End
      Begin VB.Label lbl_orden_camb 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar a -->"
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
         Height          =   195
         Left            =   2760
         TabIndex        =   62
         Top             =   375
         Width           =   1140
      End
   End
   Begin VB.Frame FraDet4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MODIFICA DATOS:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5160
      Left            =   9960
      TabIndex        =   79
      Top             =   3240
      Visible         =   0   'False
      Width           =   9180
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   9060
         TabIndex        =   145
         Top             =   4440
         Width           =   9060
         Begin VB.PictureBox BtnCancelar4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4440
            Picture         =   "tw_cronograma_mensual_inst.frx":637B
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   147
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGraba4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2880
            Picture         =   "tw_cronograma_mensual_inst.frx":6C67
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   146
            Top             =   0
            Width           =   1280
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "_________ Descripción del Insumo _________________________ Código Insumo ___ Cant.X.Mes __"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2895
         Left            =   120
         TabIndex        =   86
         Top             =   1560
         Width           =   8895
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   96
            Top             =   295
            Width           =   255
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   94
            Top             =   775
            Width           =   255
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6960
            TabIndex        =   93
            Top             =   1255
            Width           =   255
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   6960
            TabIndex        =   92
            Top             =   1735
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   6960
            TabIndex        =   91
            Top             =   2215
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "tw_cronograma_mensual_inst.frx":7455
            DataField       =   "bien_codigo1"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   98
            Top             =   285
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6A 
            Bindings        =   "tw_cronograma_mensual_inst.frx":746E
            DataField       =   "bien_codigo2"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   99
            Top             =   765
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6B 
            Bindings        =   "tw_cronograma_mensual_inst.frx":7487
            DataField       =   "bien_codigo3"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   100
            Top             =   1245
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6C 
            Bindings        =   "tw_cronograma_mensual_inst.frx":74A0
            DataField       =   "bien_codigo4"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   101
            Top             =   1725
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6D 
            Bindings        =   "tw_cronograma_mensual_inst.frx":74B9
            DataField       =   "bien_codigo5"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   102
            Top             =   2205
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox Txt_cant1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cantidad1"
            DataSource      =   "Ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7560
            TabIndex        =   95
            Text            =   "0"
            Top             =   280
            Width           =   975
         End
         Begin VB.TextBox Txt_cant2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cantidad2"
            DataSource      =   "Ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7560
            TabIndex        =   90
            Text            =   "0"
            Top             =   760
            Width           =   975
         End
         Begin VB.TextBox Txt_cant3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cantidad3"
            DataSource      =   "Ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7560
            TabIndex        =   89
            Text            =   "0"
            Top             =   1240
            Width           =   975
         End
         Begin VB.TextBox Txt_cant4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cantidad4"
            DataSource      =   "Ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7560
            TabIndex        =   88
            Text            =   "0"
            Top             =   1720
            Width           =   975
         End
         Begin VB.TextBox Txt_cant5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cantidad5"
            DataSource      =   "Ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7560
            TabIndex        =   87
            Text            =   "0"
            Top             =   2200
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "tw_cronograma_mensual_inst.frx":74D2
            DataField       =   "bien_codigo1"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   1080
            TabIndex        =   97
            Top             =   285
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6A 
            Bindings        =   "tw_cronograma_mensual_inst.frx":74EB
            DataField       =   "bien_codigo2"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   1080
            TabIndex        =   103
            Top             =   765
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6B 
            Bindings        =   "tw_cronograma_mensual_inst.frx":7504
            DataField       =   "bien_codigo3"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   1080
            TabIndex        =   104
            Top             =   1245
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6C 
            Bindings        =   "tw_cronograma_mensual_inst.frx":751D
            DataField       =   "bien_codigo4"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   1080
            TabIndex        =   105
            Top             =   1725
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6D 
            Bindings        =   "tw_cronograma_mensual_inst.frx":7536
            DataField       =   "bien_codigo5"
            DataSource      =   "Ado_detalle2"
            Height          =   315
            Left            =   1080
            TabIndex        =   106
            Top             =   2205
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_insumo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insumo 5"
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
            Height          =   195
            Left            =   240
            TabIndex        =   111
            Top             =   2220
            Width           =   780
         End
         Begin VB.Label lbl_insumo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insumo 2"
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
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   780
            Width           =   780
         End
         Begin VB.Label lbl_insumo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insumo 4"
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
            Height          =   195
            Left            =   240
            TabIndex        =   109
            Top             =   1740
            Width           =   780
         End
         Begin VB.Label lbl_insumo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insumo 1"
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
            Height          =   195
            Left            =   240
            TabIndex        =   108
            Top             =   300
            Width           =   780
         End
         Begin VB.Label lbl_insumo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Insumo 3"
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
            Height          =   195
            Left            =   240
            TabIndex        =   107
            Top             =   1260
            Width           =   780
         End
      End
      Begin VB.TextBox txt_obs 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "observaciones"
         DataSource      =   "Ado_detalle2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Text            =   "tw_cronograma_mensual_inst.frx":754F
         Top             =   1080
         Width           =   8805
      End
      Begin VB.ComboBox cmd_campo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_cronograma_mensual_inst.frx":7551
         Left            =   1920
         List            =   "tw_cronograma_mensual_inst.frx":755E
         TabIndex        =   82
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txt_codigo01 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   81
         Text            =   "0"
         Top             =   360
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "(Los datos Registrados se concatenarán al Nombre del Edificio)"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1680
         TabIndex        =   115
         Top             =   840
         Width           =   4485
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar Horario a:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   80
         Top             =   375
         Width           =   1590
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8535
      Begin VB.OptionButton OptFilGral3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2021"
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
         Left            =   4080
         TabIndex        =   158
         Top             =   2835
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral0 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2019"
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
         Left            =   1320
         TabIndex        =   157
         Top             =   2835
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Left            =   5520
         TabIndex        =   3
         Top             =   2835
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2020"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2835
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2760
         Width           =   8385
         _ExtentX        =   14790
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
         BackColor       =   12632256
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "tw_cronograma_mensual_inst.frx":75A3
         Height          =   2490
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   4392
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "fmes_correl"
            Caption         =   "Mes"
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
            DataField       =   "observaciones"
            Caption         =   "Grupo.Piloto"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Responsable"
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
            DataField       =   "fmes_nro_dias"
            Caption         =   "Nro.Dias"
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
            DataField       =   "fmes_plan"
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
         BeginProperty Column06 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Responsable"
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
         BeginProperty Column07 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija la nueva Zona a la que se enviará el registro elegido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2160
      Left            =   9840
      TabIndex        =   54
      Top             =   7320
      Visible         =   0   'False
      Width           =   7140
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   7020
         TabIndex        =   148
         Top             =   1440
         Width           =   7020
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3720
            Picture         =   "tw_cronograma_mensual_inst.frx":75BB
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   150
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGraba3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2160
            Picture         =   "tw_cronograma_mensual_inst.frx":7EA7
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   149
            Top             =   0
            Width           =   1280
         End
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6600
         TabIndex        =   55
         Top             =   690
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "tw_cronograma_mensual_inst.frx":8695
         Height          =   315
         Left            =   240
         TabIndex        =   56
         Top             =   680
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "tw_cronograma_mensual_inst.frx":86AE
         Height          =   315
         Left            =   5880
         TabIndex        =   57
         Top             =   680
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto"
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
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   405
         Width           =   990
      End
   End
   Begin VB.PictureBox fraOpciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   37
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17760
         Picture         =   "tw_cronograma_mensual_inst.frx":86C7
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   53
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4320
         Picture         =   "tw_cronograma_mensual_inst.frx":8E89
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   45
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "tw_cronograma_mensual_inst.frx":963E
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   44
         ToolTipText     =   "Aprueba Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         Picture         =   "tw_cronograma_mensual_inst.frx":9E71
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   43
         ToolTipText     =   "Anular Cronograma"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "tw_cronograma_mensual_inst.frx":A5BD
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   42
         ToolTipText     =   "Modifica Datos Cabecera"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7200
         Picture         =   "tw_cronograma_mensual_inst.frx":AED2
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   10080
         Picture         =   "tw_cronograma_mensual_inst.frx":B691
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   9360
         Picture         =   "tw_cronograma_mensual_inst.frx":B89B
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   10
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
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
         Left            =   12255
         TabIndex        =   40
         Top             =   200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CRONOGRAMA FINAL (DESTINO)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9015
      Left            =   9960
      TabIndex        =   35
      Top             =   720
      Width           =   9255
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Solo Horarios Libres"
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
         Left            =   3480
         TabIndex        =   78
         Top             =   8760
         Width           =   2475
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Todos los Horarios"
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
         Left            =   6360
         TabIndex        =   75
         Top             =   8760
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver los Horarios Laborables"
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
         Left            =   360
         TabIndex        =   74
         Top             =   8760
         Width           =   2835
      End
      Begin VB.PictureBox fraOpciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   75
         ScaleHeight     =   660
         ScaleWidth      =   9120
         TabIndex        =   70
         Top             =   240
         Width           =   9120
         Begin VB.PictureBox BtnAñadir3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6120
            Picture         =   "tw_cronograma_mensual_inst.frx":BCDD
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   114
            ToolTipText     =   "Actualiza Insumos desde Cronograma por Contrato"
            Top             =   0
            Width           =   1200
         End
         Begin VB.PictureBox BtnAñadir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4680
            Picture         =   "tw_cronograma_mensual_inst.frx":C4F0
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   85
            ToolTipText     =   "Habilita Horario (cambia a  HORARIO LABORABLE)"
            Top             =   0
            Width           =   1320
         End
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1680
            Picture         =   "tw_cronograma_mensual_inst.frx":CDBD
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   73
            ToolTipText     =   "Cambia Estado del Horario"
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3240
            Picture         =   "tw_cronograma_mensual_inst.frx":D6D2
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   72
            ToolTipText     =   "Anula Horario (cambia a NO LABORABLE)"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7560
            Picture         =   "tw_cronograma_mensual_inst.frx":DE1E
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   71
            ToolTipText     =   "Imprime R-302 Cronograma Mensual Final (Destino)"
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label lbl_texto2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H00FFFFC0&
            Height          =   300
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   1935
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "tw_cronograma_mensual_inst.frx":E6EB
         Height          =   7785
         Left            =   75
         TabIndex        =   36
         Top             =   960
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   13732
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "fmes_plan"
            Caption         =   "Mes"
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
            DataField       =   "dia_correl"
            Caption         =   "#.Dia"
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
            DataField       =   "dia_fecha"
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
         BeginProperty Column03 
            DataField       =   "dia_nombre"
            Caption         =   "Nombre.Dia"
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
            DataField       =   "horario_codigo"
            Caption         =   "Horario"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Hora.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "hora_salida"
            Caption         =   "Hora.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nro_total_horas"
            Caption         =   "#.Horas"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Equipo"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tec.Mantenimiento"
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
            DataField       =   "beneficiario_codigo_resp2"
            Caption         =   "Tec.Emergencias"
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
            DataField       =   "estado_activo"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Edificio"
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
            DataField       =   "cantidad1"
            Caption         =   "Haipe/Trapo"
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
            DataField       =   "cantidad2"
            Caption         =   "Gasolina"
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
            DataField       =   "cantidad3"
            Caption         =   "ISO-680"
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
            DataField       =   "cantidad4"
            Caption         =   "ISO-2050"
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
            DataField       =   "cantidad5"
            Caption         =   "Grasa"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   2984.882
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
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CRONOGRAMA ELABORADO (ORIGEN)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5775
      Left            =   0
      TabIndex        =   14
      Top             =   3960
      Width           =   8565
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Solo los Horarios Pendientes"
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
         Left            =   4800
         TabIndex        =   77
         Top             =   5520
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Todos Horarios Elaborados"
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
         Left            =   720
         TabIndex        =   76
         Top             =   5520
         Width           =   3015
      End
      Begin VB.PictureBox fraOpciones3 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   8415
         TabIndex        =   67
         Top             =   240
         Width           =   8415
         Begin VB.PictureBox BtnVer2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            Picture         =   "tw_cronograma_mensual_inst.frx":E706
            ScaleHeight     =   615
            ScaleWidth      =   1575
            TabIndex        =   127
            ToolTipText     =   "Actualiza #Horas y Orden"
            Top             =   0
            Width           =   1575
         End
         Begin VB.PictureBox BtnAnlDetalle4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            Picture         =   "tw_cronograma_mensual_inst.frx":F6EF
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   69
            ToolTipText     =   "Anula Horario"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4560
            Picture         =   "tw_cronograma_mensual_inst.frx":FE3B
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   68
            ToolTipText     =   "Imprime R-302 Origen (Borrador)"
            Top             =   0
            Width           =   1400
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "tw_cronograma_mensual_inst.frx":10708
         Height          =   4560
         Left            =   75
         TabIndex        =   15
         Top             =   960
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8043
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "fmes_plan"
            Caption         =   "Mes"
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
            DataField       =   "bien_orden"
            Caption         =   "Orden"
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
            DataField       =   "dia_fecha"
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
         BeginProperty Column03 
            DataField       =   "dia_nombre"
            Caption         =   "Nombre.Dia"
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
            DataField       =   "horario_codigo"
            Caption         =   "Horario"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Hora.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "hora_salida"
            Caption         =   "Hora.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nro_total_horas"
            Caption         =   "#.Horas"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Equipo"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tec.Mantenimiento"
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
            DataField       =   "beneficiario_codigo_resp2"
            Caption         =   "Tec.Emergencias"
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
            DataField       =   "estado_activo"
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
            DataField       =   "estado_codigo"
            Caption         =   "Enviado"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Edificio"
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
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         BeginProperty Column15 
            DataField       =   "dia_correl"
            Caption         =   "#.Dia"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
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
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   524.976
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   10200
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
      Left            =   2160
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   4320
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   13320
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   11040
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   8760
      Top             =   10200
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
   Begin Crystal.CrystalReport CR01 
      Left            =   4560
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
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -1560
      Top             =   23640
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
      Caption         =   "Ado_datos23"
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
      Left            =   0
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   6480
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2280
      Top             =   10560
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
   Begin Crystal.CrystalReport CR02 
      Left            =   5040
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
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   9105
      Left            =   8640
      ScaleHeight     =   9075
      ScaleWidth      =   1245
      TabIndex        =   18
      Top             =   600
      Width           =   1275
      Begin VB.PictureBox BtnAnlDetalle3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":10723
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   137
         Top             =   7800
         Width           =   1095
      End
      Begin VB.PictureBox BtnModDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":11171
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   136
         Top             =   7080
         Width           =   1095
      End
      Begin VB.PictureBox BtnAddDetalle3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "tw_cronograma_mensual_inst.frx":11A8B
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   135
         Top             =   6120
         Width           =   1095
      End
      Begin VB.PictureBox BtnAnlDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":123E0
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   134
         Top             =   5160
         Width           =   1095
      End
      Begin VB.PictureBox BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":12D09
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   133
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton BtnImprimir3 
         BackColor       =   &H80000015&
         Caption         =   "Edif.X.Zona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":13656
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Imprime Edificios por Zonas"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   15600
      Top             =   10200
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
End
Attribute VB_Name = "tw_cronograma_mensual_inst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset
Dim rs_aux13 As New ADODB.Recordset
Dim rs_aux14 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
'Dim swnuevo As String
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES, gestion0 As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_ZONA, VAR_UNITEC As String
Dim VAR_EDIF, VAR_EQP As String
Dim VAR_OBS, VAR_EQP2 As String
Dim VAR_ANL, VAR_SW2, VAR_MSG As String
Dim VAR_DA, VAR_UORIGEN, VAR_DPTOC As String

Dim VAR_AUX, VAR_CONT2 As Double
Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Dim VAR_AUX2, VAR_COD0, CONT3 As Integer
Dim DIAS_HAB, NRO_HRS, NRO_HORARIO As Integer
Dim VAR_ORDEN, VAR_MES, VAR_FMES As Integer
Dim buscados, busca3, VAR_CONT As Integer
Dim VAR_REG, VAR_CANT1 As Integer

Dim VAR_FECH1, VAR_FECH2 As Date
Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
     If VAR_SW <> "MOD" Then
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
'                Call ABRIR_TABLA_DET3
'            Case "4"
'
'        End Select
        If Ado_datos.Recordset.RecordCount > 0 Then
            VAR_FMES = Ado_datos.Recordset!fmes_plan
            buscados = buscados + 1
            If busca3 = 1 Then
                If buscados = 1 Then
                    Call ABRIR_TABLA_DET
                    If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                        lbl_texto2.Caption = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
                        lbl_texto3.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                    End If
                    'mes2 = MonthName(Month(DTPFec_Inicio.Value))
                    buscados = buscados + 1
                End If
            Else
                Call ABRIR_TABLA_DET
                If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                    lbl_texto2.Caption = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
                    lbl_texto3.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                End If
                buscados = buscados + 1
            End If
        Else
            Set dg_det1.DataSource = rsNada
            Set dg_det2.DataSource = rsNada
        End If
        If glusuario = "MLLOSA" Then
            BtnModificar.Visible = False
            BtnEliminar.Visible = False
            BtnAprobar.Visible = False
            BtnAnlDetalle4.Visible = False
            BtnModDetalle.Visible = False
            BtnAnlDetalle.Visible = False
            BtnAñadir2.Visible = False
            BtnGraba4.Visible = False
            BtnGrabar2.Visible = False
            BtnGraba3.Visible = False
            BtnAddDetalle.Visible = False
            BtnAnlDetalle2.Visible = False
            BtnAddDetalle3.Visible = False
            BtnModDetalle2.Visible = False
            BtnAnlDetalle3.Visible = False
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    'GENERA CRONOGRAMA FINAL ITEM x ITEM (INI)
    fraOpciones.Enabled = False
    fraOpciones2.Enabled = False
    FrmABMDet.Enabled = False
    FraDet3.Visible = True
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "Select * from to_cronograma_detalle WHERE unidad_codigo_tec = '" & Ado_detalle1.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_detalle1.Recordset!tec_plan_codigo & "  and bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  ", db, adOpenStatic
    If rs_aux7.RecordCount > 0 Then
        'txtnrohrs.Text = rs_aux7!bien_cantidad_por_empaque
        'cmd_campo2.Text = rs_aux7!bien_cantidad_por_empaque
        txtnrohrs.Text = Ado_detalle1.Recordset!nro_total_horas
        cmd_campo2.Text = Ado_detalle1.Recordset!nro_total_horas
    Else
        txtnrohrs.Text = "2"
        cmd_campo2.Text = "2"
    End If
    'GENERA CRONOGRAMA FINAL ITEM x ITEM (FIN)
  Else
      MsgBox "No se puede ENVIAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnAddDetalle3_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    'INI ENVIA A OTRA ZONA
    fraOpciones.Enabled = False
    fraOpciones2.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Visible = True
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    If Ado_datos.Recordset!zpiloto_codigo = "37" Then
        rs_aux7.Open "Select * from tc_zonas_piloto WHERE zpiloto_codigo <> '" & Ado_datos.Recordset!zpiloto_codigo & "'   ", db, adOpenStatic
    Else
        rs_aux7.Open "Select * from tc_zonas_piloto WHERE zpiloto_codigo = '37'  ", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_aux7
    If rs_aux7.RecordCount > 0 Then
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        VAR_ZONA = rs_aux7!zpiloto_codigo
    Else
        VAR_ZONA = "0"
    End If
    'FIN ENVIA A OTRA ZONA
  Else
      MsgBox "No se puede ENVIAR A OTRA ZONA, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If

End Sub

Private Sub BtnAnlDetalle_Click()
  
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    If Ado_detalle2.Recordset!bien_codigo <> "" And Ado_detalle2.Recordset!edif_descripcion <> " " Then
        MsgBox "No se puede ANULAR, porque ya tiene un equipo asignado en este horario ..." & vbCrLf & "Solo puede ANULAR, horarios Libres..", vbExclamation, "Validación de Registro"
    Else
        VAR_SW2 = "ANL"
        sino = MsgBox("Elige SI: para cambiar a HORARIO NO LABORABLE SOLO el registro elegido ..." & vbCrLf & "Elija NO: Para cambiar a HORARIO NO LABORABLE, de acuerdo a los parámetros a elegir ...", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            Ado_detalle2.Recordset!estado_activo = "ANL"
            Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
            Ado_detalle2.Recordset!edif_descripcion = " "
            Ado_detalle2.Recordset.Update
        
    '      Set rs_aux6 = New ADODB.Recordset
    '      If rs_aux6.State = 1 Then rs_aux6.Close
    '      rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
    '      If rs_aux6.RecordCount > 0 Then
    '        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and bien_codigo ='' "
    '        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo = 'REG' where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' "
    '        Call ABRIR_TABLA_DET
    '      End If
        Else
            VAR_MSG = "Anular (Marcar como Honario NO Laborable) ..."
            FraDet7.Caption = FraDet7.Caption + VAR_MSG
            FraDet7.Visible = True
            
            fraOpciones.Visible = False
            FrmABMDet.Visible = False
            FraGrabarCancelar.Visible = False
            fraOpciones2.Visible = False
            
            'dia_fecha Inicial
            Set rs_aux10 = New ADODB.Recordset
            If rs_aux10.State = 1 Then rs_aux10.Close
            rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND estado_activo <> 'ANL' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
            Set Ado_datos10.Recordset = rs_aux10
            If Ado_datos10.Recordset.RecordCount > 0 Then
            End If
        
            'dia_fecha Final
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND estado_activo <> 'ANL' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
            Set Ado_datos11.Recordset = rs_aux11
            If Ado_datos11.Recordset.RecordCount > 0 Then
            End If
        
        End If
    End If
        
'    If Ado_detalle2.Recordset("estado_activo") = "REG" Or Ado_detalle2.Recordset("estado_activo") = "APC" Then
'      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma Final - Destino) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between ('" & CDate(DTPfecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "') "
'        Ado_detalle2.Recordset!estado_activo = "ANL"
'        Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
'        Ado_detalle2.Recordset!edif_descripcion = " "
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'      End If
'    Else
'        MsgBox "No se puede ANULAR, el registro ya fue Aprobado (Estado=APR) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
'    End If
  Else
      MsgBox "No se puede ANULAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If

  'WWWWWWWWWWWWWWWWWWWWW
'  If Ado_detalle2.Recordset.RecordCount > 0 Then
'    If ExisteReg(Ado_detalle2.Recordset!fmes_plan) Then MsgBox "No se puede RETORNAR TODO, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
'    sino = MsgBox("Elige SI: para cambiar a HORARIO NO LABORABLE el registro elegido ..." & vbCrLf & "Elija NO: Para cambiar a HORARIO NO LABORABLE, de acuerdo a los parámetros elegidos a continuación ...", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        Ado_detalle2.Recordset!estado_activo = "ANL"
'        Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
'        Ado_detalle2.Recordset!edif_descripcion = " "
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'    Else
'        'edif_descripcion
'        Set rs_aux9 = New ADODB.Recordset
'        If rs_aux9.State = 1 Then rs_aux9.Close
'        rs_aux9.Open "Select edif_descripcion from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND edif_descripcion <> '' group  by edif_descripcion order by edif_descripcion ", db, adOpenStatic
'        Set Ado_datos9.Recordset = rs_aux9
''        dtc_desc9.BoundText = dtc_codigo9.BoundText
'
'        'dia_fecha Inicial
'        Set rs_aux10 = New ADODB.Recordset
'        If rs_aux10.State = 1 Then rs_aux10.Close
'        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos10.Recordset = rs_aux10
'        If Ado_datos10.Recordset.RecordCount > 0 Then
'        End If
'
'        'dia_fecha Final
'        Set rs_aux11 = New ADODB.Recordset
'        If rs_aux11.State = 1 Then rs_aux11.Close
'        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos11.Recordset = rs_aux11
'        If Ado_datos11.Recordset.RecordCount > 0 Then
'        End If
'        VAR_ANL = "RET"
'        FraDet5.Caption = "Elija los parámetros para retornar al Crono. Origen..."
'        FraDet5.Visible = True
'    End If
'  Else
'        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
'  End If
  'WWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub BtnAnlDetalle2_Click()
  'If ExisteReg2(Ado_detalle2.Recordset!fmes_plan, Ado_detalle2.Recordset!bien_codigo) Then MsgBox "No se puede RETORNAR 1, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
  
  If Ado_datos.Recordset!estado_codigo = "REG" Then
   'If Ado_detalle2.Recordset!estado_codigo = "REG" And Ado_detalle2.Recordset!estado_activo = "APR" Then
   If Ado_detalle2.Recordset!estado_activo = "APR" Then
      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Final) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        'db.Execute "update to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        'db.Execute "update to_cronograma_diario_final set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        db.Execute "update to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "           'AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & "
        db.Execute "update to_cronograma_diario_final set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "         'bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
  Else
      MsgBox "No se puede RETORNAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnAnlDetalle3_Click()
  If Ado_detalle2.Recordset.RecordCount > 0 Then
    If ExisteReg(Ado_detalle2.Recordset!fmes_plan) Then MsgBox "No se puede RETORNAR TODO, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
    
    sino = MsgBox("Elige SI: para RETORNAR TODO el Cronograma DESTINO al ORIGEN..." & vbCrLf & "Elija NO: Para RETORNAR al Cronograma ORIGEN, registros de acuerdo a los parámetros elegidos...", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
      Set rs_aux6 = New ADODB.Recordset
      If rs_aux6.State = 1 Then rs_aux6.Close
      rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
      If rs_aux6.RecordCount > 0 Then
        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' "

        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo = 'REG' where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' "
      
        Call ABRIR_TABLA_DET
      End If
    Else
        'edif_descripcion
        Set rs_aux9 = New ADODB.Recordset
        If rs_aux9.State = 1 Then rs_aux9.Close
        rs_aux9.Open "Select edif_descripcion from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND edif_descripcion <> '' group  by edif_descripcion order by edif_descripcion ", db, adOpenStatic
        Set Ado_datos9.Recordset = rs_aux9
'        dtc_desc9.BoundText = dtc_codigo9.BoundText

        'dia_fecha Inicial
        Set rs_aux10 = New ADODB.Recordset
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos10.Recordset = rs_aux10
        If Ado_datos10.Recordset.RecordCount > 0 Then
        End If

        'dia_fecha Final
        Set rs_aux11 = New ADODB.Recordset
        If rs_aux11.State = 1 Then rs_aux11.Close
        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos11.Recordset = rs_aux11
        If Ado_datos11.Recordset.RecordCount > 0 Then
        End If
        VAR_ANL = "RET"
        FraDet5.Caption = "Elija los parámetros para retornar al Crono. Origen..."
        FraDet5.Visible = True
    End If
  Else
        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
  End If

'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    'to_cronograma_diario_final
'    Set rs_aux6 = New ADODB.Recordset
'    If rs_aux6.State = 1 Then rs_aux6.Close
'    rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_codigo <> '' ", db, adOpenStatic
'    If rs_aux6.RecordCount > 0 Then
'      sino = MsgBox("Está Seguro de RETORNAR TODO ? (Se Retornará TODO el Cronograma DESTINO al ORIGEN) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', estado_activo = 'REG' WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "
'
'        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo   = 'REG' where fmes_plan  = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "
'
'        Call ABRIR_TABLA_DET
'      End If
'    Else
'        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique los registros ...", vbExclamation, "Validación de Registro"
'    End If
'  Else
'      MsgBox "No se puede RETORNAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
'  End If
End Sub

Private Function ExisteReg(codigo2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM to_cronograma_diario_final  WHERE fmes_plan = " & codigo2 & " AND bien_codigo <> '' and (nro_fojas IS NOT NULL) AND (doc_numero IS NOT NULL)  "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Function ExisteReg2(codigo2 As String, codigo3 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos2 FROM to_cronograma_diario_final  WHERE fmes_plan = " & codigo2 & " AND bien_codigo = '" & codigo3 & "' and (nro_fojas IS NOT NULL or nro_fojas='0') AND (doc_numero IS NOT NULL or doc_numero ='0')  "
    GlSqlAux = "SELECT Count(*) AS Cuantos2 FROM to_cronograma_diario_final  WHERE fmes_plan = " & codigo2 & " AND bien_codigo = '" & codigo3 & "' and (nro_fojas IS NOT NULL or nro_fojas='0') AND (doc_numero IS NOT NULL or doc_numero ='0')  "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg2 = rs!Cuantos2 > 0
End Function

Private Sub BtnAnlDetalle4_Click()
 If Ado_datos.Recordset!estado_activo = "REG" Then
   If Ado_detalle1.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Elaborado - Origen) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "update to_cronograma_diario_inst set estado_activo = 'ANL', estado_codigo = 'ANL' WHERE fmes_plan = " & VAR_FMES & " AND horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & " AND  bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
        'db.Execute "update to_cronograma_diario_inst set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue ENVIADO al Cronograma Destino o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
 Else
      MsgBox "No se puede ANULAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
 End If
End Sub

Private Sub BtnAñadir2_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    VAR_SW2 = "HAB"
    
    sino = MsgBox("Elige SI: para Habilitar a HORARIO LABORABLE SOLO el registro elegido ..." & vbCrLf & "Elija NO: Para Habilitar a HORARIO LABORABLE, de acuerdo a los parámetros a elegir ...", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        If Ado_detalle2.Recordset("estado_activo") = "ANL" Or Ado_detalle2.Recordset("estado_activo") = "APC" Then
           'sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
           'If sino = vbYes Then
            Ado_detalle2.Recordset!estado_activo = "REG"
            Ado_detalle2.Recordset!observaciones = " "       '"HORARIO LABORABLE"
            Ado_detalle2.Recordset!edif_descripcion = " "
            Ado_detalle2.Recordset.Update
             'Call ABRIR_TABLA_DET
           'End If
        Else
            If (Ado_detalle2.Recordset!bien_codigo = "" And Ado_detalle2.Recordset!tec_plan_codigo = "0") Then
             'sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
             'If sino = vbYes Then
               Ado_detalle2.Recordset!estado_activo = "REG"
               Ado_detalle2.Recordset!observaciones = " "        '"HORARIO LABORABLE"
               Ado_detalle2.Recordset!edif_descripcion = " "
               Ado_detalle2.Recordset.Update
             'End If
            Else
               MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR, APC, APP) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
            End If
        End If
   Else
        VAR_MSG = "Habilitar (Marcar como Honario LABORABLE) ..."
        FraDet7.Caption = FraDet7.Caption + VAR_MSG
        FraDet7.Visible = True
        
        fraOpciones.Visible = False
        FrmABMDet.Visible = False
        FraGrabarCancelar.Visible = False
        fraOpciones2.Visible = False
        
        'dia_fecha Inicial
        Set rs_aux10 = New ADODB.Recordset
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos10.Recordset = rs_aux10
        If Ado_datos10.Recordset.RecordCount > 0 Then
        End If
    
        'dia_fecha Final
        Set rs_aux11 = New ADODB.Recordset
        If rs_aux11.State = 1 Then rs_aux11.Close
        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final where fmes_plan  = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos11.Recordset = rs_aux11
        If Ado_datos11.Recordset.RecordCount > 0 Then
        End If
   End If

'    'Actualiza Codigos de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Actualiza Cantidad de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
'    'Actualiza Carta al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
  Else
      MsgBox "No se puede HABILITAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnAñadir3_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    FraInsumos.Visible = True
'    VAR_AUX2 = VAR_FMES     ' fmes_plan
'    'Carga "Codigos de Insumos" al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_y_detalle.bien_codigo1 , to_cronograma_diario_final.bien_codigo2 = tv_cronograma_y_detalle.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_y_detalle.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_y_detalle.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_y_detalle.bien_codigo5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "
'
'    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'
'    'Actualiza Cantidad de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_y_detalle.cantidad1 , to_cronograma_diario_final.cantidad2 = tv_cronograma_y_detalle.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_y_detalle.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_y_detalle.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_y_detalle.cantidad5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "
'
'    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " "
'
'    'Quita Cantidad de Insumo3 e Insumo4 en meses pares al Crono Final
'    sino = MsgBox("Elija SI: para programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC) los insumos 3 y 4..." & vbCr & _
'             "Elija NO: para programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV) los insumos 3 y 4....", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        'PROGRAMAR en Mes PAR y QUITAR Mes IMPAR
'        db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0', to_cronograma_diario_final.cantidad4 = '0' From to_cronograma_diario_final INNER JOIN tv_cronograma_mensual_impar ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_mensual_impar.fmes_plan ) " & _
'        " where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & "  ) "
'
'
'    Else
'        'Mes PAR
'        db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0', to_cronograma_diario_final.cantidad4 = '0' From to_cronograma_diario_final INNER JOIN tv_cronograma_mensual_par ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_mensual_par.fmes_plan ) " & _
'        " where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & " ) "
'    End If
'
'    'Actualiza Cantidad de Insumos al Crono Final Bmes, Tmes, etc.
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and tv_cronograma_insumos.unimed_codigo <> 'MES' "
'
'    'Actualiza Carta al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_carta.carta " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_carta ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_carta.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final.bien_codigo <> '' "
'
'    db.Execute " update to_cronograma_diario_final set to_cronograma_diario_final.carta = tv_cronograma_y_detalle.carta from to_cronograma_diario_final inner join tv_cronograma_y_detalle on to_cronograma_diario_final.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final.bien_codigo <> '' "
'
''    'Actualiza Carta al Crono Final
''    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
''    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'
'    'Carga Codigos de Insumos al Crono Final de Otras Zonas
'    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Actualiza Cantidad de Insumos al Crono Final de Otras Zonas
'    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " "
'
'    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    'db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0', to_cronograma_diario_final.cantidad4 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan ) " & _
'    '" where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
'    'Actualiza Cantidad de Insumos al Crono Final Bmes, Tmes, etc.
'    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4 " & _
'    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and tv_cronograma_insumos.unimed_codigo <> 'MES' "
'
'    MsgBox "Se actualizaron los Insumos desde CRONOGRAMA POR CONTRATO correspondientes a la misma Gestión y Zona del CRONOGRAMA FINAL (DESTINO) ...", vbInformation, "Información"
  Else
      MsgBox "No se puede Actualizar INSUMOS, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If

End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Cronograma Nro. " + Str(VAR_FMES), vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
'        VAR_COD0 = 0
'        VAR_EQP2 = ""
'        VAR_CONT = 0
'        Set rs_aux12 = New ADODB.Recordset
'        If rs_aux12.State = 1 Then rs_aux12.Close
'        'SQL_FOR = "SELECT to_cronograma_mensual.fmes_plan, to_cronograma_mensual.ges_gestion, to_cronograma_mensual.fmes_correl, to_cronograma_mensual.zpiloto_codigo, to_cronograma_diario_final.dia_correl, to_cronograma_diario_final.horario_codigo, to_cronograma_diario_final.bien_codigo FROM to_cronograma_mensual INNER JOIN to_cronograma_diario_final ON to_cronograma_mensual.fmes_plan = to_cronograma_diario_final.fmes_plan where (to_cronograma_mensual.fmes_plan = " & VAR_FMES & " AND bien_codigo <> '') ORDER BY to_cronograma_diario_final.dia_correl, to_cronograma_diario_final.horario_codigo"
'        SQL_FOR = "SELECT * FROM tv_cronograma_mensual_y_final where (fmes_plan = " & VAR_FMES & " AND bien_codigo <> '') ORDER BY dia_correl, horario_codigo"
'        rs_aux12.Open SQL_FOR, db, adOpenStatic  'group  by bien_codigo
'        VAR_CONT = rs_aux12.RecordCount
'        Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'        SQL_FOR = "SELECT fmes_plan, ges_gestion, fmes_correl, zpiloto_codigo  FROM to_cronograma_mensual where ges_gestion = '" & rs_aux12!ges_gestion & "' AND fmes_correl = " & rs_aux12!fmes_correl & " + 1 AND zpiloto_codigo = " & rs_aux12!zpiloto_codigo & " "
'        rs_aux13.Open SQL_FOR, db, adOpenStatic  'group  by bien_codigo
'        If rs_aux13.RecordCount > 0 Then
'           db.Execute "update to_cronograma_diario_inst set bien_orden = " & VAR_CONT & " where fmes_plan = " & rs_aux13!fmes_plan & "  aND (bien_codigo <> '' AND bien_orden <= " & VAR_CONT & " ) "
'           db.Execute "update tc_zona_piloto_edif set zona_edif_orden = " & VAR_CONT & " where zpiloto_codigo = " & rs_aux12!zpiloto_codigo & " AND zona_edif_orden <= " & VAR_CONT & "  "
'           If rs_aux12.RecordCount > 0 Then
'              rs_aux12.MoveFirst
'              While Not rs_aux12.EOF
'                If VAR_EQP2 <> rs_aux12!bien_codigo Then
'                  VAR_COD0 = VAR_COD0 + 1
'                  db.Execute "update to_cronograma_diario_inst set bien_orden = " & VAR_COD0 & " where fmes_plan = " & rs_aux13!fmes_plan & " AND bien_codigo = '" & rs_aux12!bien_codigo & "' "
'                  db.Execute "update tc_zona_piloto_edif set zona_edif_orden = " & VAR_COD0 & " where zpiloto_codigo = " & rs_aux12!zpiloto_codigo & "  and edif_codigo = '" & rs_aux12!edif_codigo & "' "
'                End If
'                VAR_EQP2 = rs_aux12!bien_codigo
'                rs_aux12.MoveNext
'              Wend
'           End If
'        End If
      End If
       Ado_datos.Recordset!estado_codigo = "APR"
       Ado_datos.Recordset!estado_activo = "APR"
       Ado_datos.Recordset!fecha_registro = Date
       Ado_datos.Recordset!usr_codigo = glusuario
       Ado_datos.Recordset.Update
   End If
    
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    
    If Ado_datos.Recordset.RecordCount > 0 Then
        buscados = 0
        busca3 = 1
'        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
''        If Ado_datos.Recordset!estado_codigo = "REG" Then
''            Call OptFilGral1_Click
''        Else
''            Call OptFilGral2_Click
''        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existen registros. ", vbExclamation, "Atención!"
      'OptFilGral1.Visible = True
      'OptFilGral2.Visible = True
    End If

End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Visible = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
    End If

End Sub

Private Sub BtnCancelar2_Click()
    fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet3.Visible = False
     cmd_campo2.Text = "2"
End Sub

Private Sub BtnCancelar3_Click()
    fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet2.Visible = False
End Sub

Private Sub BtnCancelar4_Click()
    FraDet4.Visible = False
    fraOpciones.Enabled = True
    FrmABMDet.Enabled = True
    fraOpciones2.Enabled = True
End Sub

Private Sub BtnCancelar5_Click()
     fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet2.Visible = False
     FraDet5.Visible = False
End Sub

Private Sub BtnCancelar6_Click()
    FraDet6.Visible = False
End Sub

Private Sub BtnCancelar7_Click()
    FraDet7.Visible = False
    VAR_SW2 = ""
    VAR_MSG = ""
    fraOpciones.Visible = True
    FrmABMDet.Visible = True
    FraGrabarCancelar.Visible = True
    fraOpciones2.Visible = True
End Sub

Private Sub BtnCancelar8_Click()
    FraInsumos.Visible = False
End Sub

Private Sub BtnGraba3_Click()
   'CCCCCCCCCCCCCCCCCCCCCCCCCCCBBBBBBBBBBBBBBB
   VAR_ZONA = dtc_codigo5.Text
   VAR_MES = lbl_texto1.Caption
   gestion0 = txt_codigo1.Text
   
     Set rs_aux4 = New ADODB.Recordset
     If rs_aux4.State = 1 Then rs_aux4.Close
     rs_aux4.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
     If rs_aux4.RecordCount > 0 Then
        If rs_aux4!estado_codigo = "APR" Then
            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
            Exit Sub
        End If
        VAR_UNITEC = Ado_detalle1.Recordset!unidad_codigo_tec
        VAR_EQP = Ado_detalle1.Recordset!bien_codigo
'        VAR_FMES = Ado_detalle1.Recordset!fmes_plan
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "  and unidad_codigo_tec = '" & VAR_UNITEC & "'   ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
             VAR_AUX2 = rs_aux2!fmes_plan
             VAR_COD0 = 0
             'db.Execute "SELECT VAR_ORDEN = isnull(max(bien_orden),0) from to_cronograma_diario_inst WHERE     (fmes_plan = " & VAR_AUX2 & " ) "
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "select isnull(max(bien_orden),0) as bien_orden2 from to_cronograma_diario_inst WHERE fmes_plan = " & VAR_AUX2 & "  ", db, adOpenStatic
            If rs_aux5.RecordCount > 0 Then
               VAR_ORDEN = rs_aux5!bien_orden2 + 1
            End If
             Set rs_aux3 = New ADODB.Recordset
             If rs_aux3.State = 1 Then rs_aux3.Close
             rs_aux3.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux3.RecordCount > 0 Then
                 rs_aux3.MoveFirst
                 While Not rs_aux3.EOF
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then
                        db.Execute "update to_cronograma_diario_inst set bien_codigo = '" & rs_aux4!bien_codigo & "', unidad_codigo_tec = '" & rs_aux4!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux4!tec_plan_codigo & ", observaciones = '" & rs_aux4!observaciones & "', bien_orden = " & VAR_ORDEN & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'REG', estado_activo = 'REG', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', observaciones = 'HORARIO LABORABLE'  WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & rs_aux4!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                    End If
                    rs_aux3.MoveNext
                    'Habilitar .....
                    'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                 Wend
             End If
             db.Execute "update to_cronograma_diario_inst set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & VAR_EQP & "'  "
        End If
     End If
     db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_inst INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_inst.bien_codigo  = av_bienes_vs_edificios.bien_codigo "
     Call ABRIR_TABLA_DET
    fraOpciones.Enabled = True
    fraOpciones2.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Visible = False
End Sub

Private Sub BtnGraba4_Click()
    If Txt_cant1.Text = "" Then
        Txt_cant1.Text = "0"
    End If
    If Txt_cant2.Text = "" Then
        Txt_cant2.Text = "0"
    End If
    If Txt_cant3.Text = "" Then
        Txt_cant3.Text = "0"
    End If
    If Txt_cant4.Text = "" Then
        Txt_cant4.Text = "0"
    End If
    If Txt_cant5.Text = "" Then
        Txt_cant5.Text = "0"
    End If
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "select * from gv_equipos_vs_edificios WHERE Bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    If rs_aux8.RecordCount > 0 Then
        VAR_EDIF = Trim(rs_aux8!edif_descripcion)
    Else
        VAR_EDIF = ""
    End If
    If rs_aux8.RecordCount > 0 Then
        Select Case Trim(txt_codigo01.Text)
            Case "APC"      'COMPENSACION
                VAR_EDIF = ""
                VAR_OBS = "(" + Trim(cmd_campo1.Text) + ")"
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = '0', cantidad2 = '0', cantidad3 = '0', cantidad4 = '0', cantidad5 = '0' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "

                db.Execute "update to_cronograma_diario_final set bien_codigo = '', observaciones = '" & VAR_OBS & "', edif_descripcion = '" & VAR_EDIF & "', bien_orden = 0  WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
            
            Case "APP"      'HORARIO POR CONFIRMAR
                VAR_EDIF = Trim(rs_aux8!edif_descripcion)
                VAR_OBS = "(" + Trim(cmd_campo1.Text) + ")"
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = " & CDbl(Txt_cant1.Text) & ", cantidad2 = " & CDbl(Txt_cant2.Text) & ", cantidad3 = " & CDbl(Txt_cant3.Text) & ", cantidad4 = " & CDbl(Txt_cant4.Text) & ", cantidad5 = " & CDbl(Txt_cant5.Text) & " WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
                
                db.Execute "update to_cronograma_diario_final set edif_descripcion = '" & VAR_EDIF & "', observaciones = '" & VAR_OBS & "' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND bien_codigo  = '" & Ado_detalle2.Recordset!bien_codigo & "' "
                
            Case "APR"      'HORARIO LABORAL Confirmado
                VAR_EDIF = Trim(rs_aux8!edif_descripcion)
                If txt_obs.Text = "" Then
                    VAR_OBS = ""
                Else
                    VAR_OBS = "(" + Trim(txt_obs.Text) + ")"
                End If
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = " & CDbl(Txt_cant1.Text) & ", cantidad2 = " & CDbl(Txt_cant2.Text) & ", cantidad3 = " & CDbl(Txt_cant3.Text) & ", cantidad4 = " & CDbl(Txt_cant4.Text) & ", cantidad5 = " & CDbl(Txt_cant5.Text) & " WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
                
                db.Execute "update to_cronograma_diario_final set edif_descripcion = '" & VAR_EDIF & "', observaciones = '" & VAR_OBS & "' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND bien_codigo  = '" & Ado_detalle2.Recordset!bien_codigo & "' "
                
        End Select
        
    Else
        Select Case Trim(txt_codigo01.Text)
            Case "APC"  'COMPENSACION
                VAR_EDIF = ""
                VAR_OBS = cmd_campo1.Text
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = '0', cantidad2 = '0', cantidad3 = '0', cantidad4 = '0', cantidad5 = '0' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "

                db.Execute "update to_cronograma_diario_final set bien_codigo = '', observaciones = '" & VAR_OBS & "', edif_descripcion = '" & VAR_EDIF & "', bien_orden = 0  WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
                
            Case "APP"  'HORARIO POR CONFIRMAR
                VAR_EDIF = Trim(cmd_campo1.Text)
                db.Execute "update to_cronograma_diario_final set edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = " & CDbl(Txt_cant1.Text) & ", cantidad2 = " & CDbl(Txt_cant2.Text) & ", cantidad3 = " & CDbl(Txt_cant3.Text) & ", cantidad4 = " & CDbl(Txt_cant4.Text) & ", cantidad5 = " & CDbl(Txt_cant5.Text) & " WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
            Case "APR"  'HORARIO LABORAL Confirmado
                VAR_EDIF = ""
                db.Execute "update to_cronograma_diario_final set edif_descripcion = '" & VAR_EDIF & "' WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
                
                db.Execute "update to_cronograma_diario_final set estado_activo = '" & txt_codigo01.Text & "', cantidad1 = " & CDbl(Txt_cant1.Text) & ", cantidad2 = " & CDbl(Txt_cant2.Text) & ", cantidad3 = " & CDbl(Txt_cant3.Text) & ", cantidad4 = " & CDbl(Txt_cant4.Text) & ", cantidad5 = " & CDbl(Txt_cant5.Text) & " WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & Ado_detalle2.Recordset!dia_correl & " AND horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & " "
                
        End Select

    End If
    Call ABRIR_TABLA_DET
    FraDet4.Visible = False
    fraOpciones.Enabled = True
    FrmABMDet.Enabled = True
    fraOpciones2.Enabled = True


'   Else
'        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
'   End If

End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     '
     Set rs_aux5 = New ADODB.Recordset
     If rs_aux5.State = 1 Then rs_aux5.Close
     rs_aux5.Open "select dia_correl from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and estado_activo <> 'ANL' group by dia_correl", db, adOpenStatic
     If rs_aux5.RecordCount > 0 Then
        DIAS_HAB = rs_aux5.RecordCount
     End If
        
     Set rs_aux5 = New ADODB.Recordset
     If rs_aux5.State = 1 Then rs_aux5.Close
     rs_aux5.Open "select COUNT(dia_correl) as nro_horarios, SUM(nro_total_horas) as nro_horas from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and estado_activo <> 'ANL' ", db, adOpenStatic
     If rs_aux5.RecordCount > 0 Then
        NRO_HORARIO = rs_aux5!nro_horarios
        NRO_HRS = rs_aux5!nro_horas
     End If
     
'     rs_datos!fmes_fecha_registro = dtpFecha1.Value
'     rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
'     rs_datos!observaciones = Txt_campo2.Text
'
'     rs_datos!fmes_nro_dias_habiles = DIAS_HAB
'     rs_datos!fmes_nro_horarios_hab = NRO_HORARIO
'     rs_datos!fmes_nro_hrs_habiles = NRO_HRS

'     rs_datos!fecha_registro = Date     'no cambia
'     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'     rs_datos.Update    'Batch 'adAffectAll
     db.Execute "Update to_cronograma_mensual Set fecha_registro= '" & Date & "', usr_codigo ='" & glusuario & "', beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = " & VAR_FMES & "   "
     db.Execute "Update to_cronograma_diario_inst Set beneficiario_codigo_resp = " & dtc_codigo4.Text & ", beneficiario_codigo_resp2 = " & dtc_codigo4.Text & " Where fmes_plan = " & VAR_FMES & "   "
     db.Execute "Update to_cronograma_diario_final Set beneficiario_codigo_resp = " & dtc_codigo4.Text & ", beneficiario_codigo_resp2 = " & dtc_codigo4.Text & " Where fmes_plan = " & VAR_FMES & "   "
     Call OptFilGral1_Click
     'rs_datos.MoveFirst
'     mbDataChanged = False

    Fra_datos.Visible = False
    Fra_datos.Enabled = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    dg_datos.Enabled = True
        
     VAR_SW = ""
'     dtc_codigo9.Enabled = True

  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'Valida compos para editables
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo4 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (Txt_campo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  
End Sub

Private Sub BtnGrabar2_Click()
     'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
     If rs_aux2.RecordCount > 0 Then
        If rs_aux2!estado_codigo = "APR" Then
            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
            Exit Sub
        End If
         VAR_AUX2 = rs_aux2!fmes_plan
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                If cmd_campo2.Text > 2 Then
                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 2
                        CONT3 = 1
                   End If
                Else
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                   End If
                End If
'                   If cmd_campo2.Text = "4" Then
'                      rs_aux3.MoveNext
'                      If VAR_COD0 < 2 And rs_aux3!estado_activo = "REG" Then        '
'                         'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
'                         db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
'                         db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         'db.Execute "update to_cronograma_diario_final set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         'db.Execute "update to_cronograma_diario_final set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         VAR_COD0 = VAR_COD0 + 1
'                         CONT3 = 1
'                      End If
'                   End If
'                   If cmd_campo2.Text = "8" Then
'                      rs_aux3.MoveNext
'                      If VAR_COD0 < 2 And rs_aux3!estado_activo = "REG" Then        '
'                         'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
'                         db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
'                         db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         'db.Execute "update to_cronograma_diario_final set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         'db.Execute "update to_cronograma_diario_final set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                         VAR_COD0 = VAR_COD0 + 1
'                         CONT3 = 1
'                      End If
'                   End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
     End If
     db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final.bien_codigo  = av_bienes_vs_edificios.bien_codigo where to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " AND to_cronograma_diario_final.bien_codigo <>'' "
    
    'Actualiza Codigos de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " "
    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

'     Call BtnAñadir2_Click
     fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet3.Visible = False
     cmd_campo2.Text = "2"
     Call ABRIR_TABLA_DET
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
End Sub

Private Sub COPIA_TODOS()
     'WWWWW GENERA TODO EL CRONOGRAMA DIARIO FINAL DESDE ORIGEN
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and bien_codigo <> '' order by bien_orden  ", db, adOpenKeyset, adLockOptimistic       'and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "
     If rs_aux2.RecordCount > 0 Then
        'If rs_aux2!estado_codigo = "APR" Then
        '    MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
        '    Exit Sub
        'End If
       VAR_AUX2 = rs_aux2!fmes_plan
       rs_aux2.MoveFirst
       While Not rs_aux2.EOF
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final where fmes_plan = " & VAR_AUX2 & "  and estado_codigo = 'REG' ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                If cmd_campo2.Text > 2 Then
                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 2
                        CONT3 = 1
                   End If
                Else
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                   End If
                End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
        rs_aux2.MoveNext
       Wend
     End If
     db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    
    'Actualiza Codigos de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " "
    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'     Call BtnAñadir2_Click
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
    'wwwwwwwwwwwwwwwwwwwwwwwwwwwww
End Sub

Private Sub COPIA_ALGUNOS()
     'WWWWW GENERA ALGUNOS EL CRONOGRAMA DIARIO FINAL DESDE ORIGEN
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and bien_codigo <> '' AND estado_activo = 'REG' order by bien_orden  ", db, adOpenKeyset, adLockOptimistic       'and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "
     If rs_aux2.RecordCount > 0 Then
        'If rs_aux2!estado_codigo = "APR" Then
        '    MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
        '    Exit Sub
        'End If
       VAR_AUX2 = rs_aux2!fmes_plan
       rs_aux2.MoveFirst
       While Not rs_aux2.EOF
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final where fmes_plan = " & VAR_AUX2 & " and estado_codigo = 'REG' ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                If cmd_campo2.Text > 2 Then
                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 2
                        CONT3 = 1
                   End If
                Else
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                   End If
                End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
        rs_aux2.MoveNext
       Wend
     End If
     db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    
    'Actualiza Codigos de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " "
    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

'     Call BtnAñadir2_Click
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
End Sub

Private Sub BtnGrabar5_Click()
    
    If DTPfecha2.Text = "Todos" Then
        DTPfecha2.Text = "01" & "/" & Trim(Ado_datos.Recordset!fmes_correl) & "/" & Trim(Ado_datos.Recordset!ges_gestion)
    End If
    If DTPfecha3.Text = "Todos" Then
        DTPfecha3.Text = Trim(Ado_datos.Recordset!fmes_nro_dias) & "/" & Trim(Ado_datos.Recordset!fmes_correl) & "/" & Trim(Ado_datos.Recordset!ges_gestion)
    End If
    If dtc_desc9.Text = "Todos" And DTPfecha2.Text <> "Todos" Then
        VAR_FECH1 = CDate(DTPfecha2.Text)
        VAR_FECH2 = CDate(DTPfecha3.Text)
        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "

        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = '00' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '')"

        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = 'OK' FROM to_cronograma_diario_inst INNER JOIN to_cronograma_diario_final ON to_cronograma_diario_inst.fmes_plan = to_cronograma_diario_final.fmes_plan and to_cronograma_diario_inst.bien_codigo  = to_cronograma_diario_final.bien_codigo WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') "

        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.estado_activo  = 'REG', to_cronograma_diario_inst.estado_codigo  = 'REG' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') AND (to_cronograma_diario_inst.hora_registro  = '00')"

        'db.Execute "UPDATE to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG'  where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' and dia_fecha between '" & CDate(dtpFecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "' "
      
        Call ABRIR_TABLA_DET
        'cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text) & "
    Else
        VAR_FECH1 = CDate(DTPfecha2.Text)
        VAR_FECH2 = CDate(DTPfecha3.Text)
        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "

        'db.Execute "UPDATE to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG'  where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between ('" & CDate(dtpFecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "') "
        
        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = '00' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '')"

        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = 'OK' FROM to_cronograma_diario_inst INNER JOIN to_cronograma_diario_final ON to_cronograma_diario_inst.fmes_plan = to_cronograma_diario_final.fmes_plan and to_cronograma_diario_inst.bien_codigo  = to_cronograma_diario_final.bien_codigo WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') "

        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.estado_activo  = 'REG', to_cronograma_diario_inst.estado_codigo  = 'REG' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') AND (to_cronograma_diario_inst.hora_registro  = '00')"
      
        Call ABRIR_TABLA_DET
    End If
    FraDet5.Visible = False

End Sub

Private Sub BtnGrabar6_Click()
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        'MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), debe deshabilitarlos (Retornar) o utilizar el botón (Envia Uno) ...", vbExclamation, "Validación de Registro"
        MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), solo podrá procesar la opción 3. ...", vbExclamation, "Validación de Registro"
        If Option8.Value = True Then
            Call COPIA_ALGUNOS
            db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
        End If
    Else
      If Option6.Value = True Then
        Call COPIA_TODOS
        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
      End If
      If Option7.Value = True Then
        db.Execute "UPDATE to_cronograma_diario_final SET to_cronograma_diario_final.bien_orden  = to_cronograma_diario_inst.bien_orden, to_cronograma_diario_final.bien_codigo = to_cronograma_diario_inst.bien_codigo, to_cronograma_diario_final.unidad_codigo_tec = to_cronograma_diario_inst.unidad_codigo_tec, " & _
        " to_cronograma_diario_final.tec_plan_codigo = to_cronograma_diario_inst.tec_plan_codigo, to_cronograma_diario_final.edif_descripcion = to_cronograma_diario_inst.edif_descripcion, to_cronograma_diario_final.estado_activo = 'APR' FROM to_cronograma_diario_final INNER JOIN to_cronograma_diario_inst " & _
        " ON to_cronograma_diario_final.fmes_plan  = to_cronograma_diario_inst.fmes_plan AND to_cronograma_diario_final.dia_correl  = to_cronograma_diario_inst.dia_correl AND to_cronograma_diario_final.horario_codigo = to_cronograma_diario_inst.horario_codigo WHERE to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "

        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
      End If
      
'      sino = MsgBox("Está Seguro de ENVIAR TODO el Cronograma ORIGEN al DESTINO ?." & vbCrLf & " SI-->(Envía solo a los Horarios Laborales definidos en el Destino) " & vbCrLf & " NO-->(Envía todo a todos los días calendario, incluyendo días NO laborales) " & vbCrLf & " Cancelar, la Operación", vbYesNoCancel + vbQuestion, "Atención")
'      If sino = vbYes Then
'      Else
'        If sino = vbNo Then
'        End If
'      End If
'        'Call BtnAñadir2_Click
      Call ABRIR_TABLA_DET
    End If
    FraDet6.Visible = False
End Sub

Private Sub BtnGrabar7_Click()
    VAR_FECH1 = CDate(DTPfecha4.Text)
    VAR_FECH2 = CDate(DTPfecha5.Text)
    
    If VAR_SW2 = "HAB" Then
        db.Execute "UPDATE to_cronograma_diario_final SET estado_activo  = 'REG', observaciones = 'HORARIO LABORABLE', edif_descripcion = '', tec_plan_codigo = '0' WHERE fmes_plan = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
    Else
        db.Execute "UPDATE to_cronograma_diario_final SET estado_activo  = 'ANL', observaciones = 'HORARIO NO LABORABLE', edif_descripcion = '', tec_plan_codigo = '0' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo <> 'APR' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
    End If
    Call ABRIR_TABLA_DET
    FraDet7.Visible = False
    VAR_SW2 = ""
    VAR_MSG = ""
    fraOpciones.Visible = True
    FrmABMDet.Visible = True
    FraGrabarCancelar.Visible = True
    fraOpciones2.Visible = True
End Sub

Private Sub BtnGrabar8_Click()
    VAR_AUX2 = VAR_FMES     ' fmes_plan
    'Carga "Codigos de Insumos" al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1 = tv_cronograma_y_detalle.bien_codigo1 , to_cronograma_diario_final.bien_codigo2 = tv_cronograma_y_detalle.bien_codigo2, to_cronograma_diario_final.bien_codigo3 = tv_cronograma_y_detalle.bien_codigo3, to_cronograma_diario_final.bien_codigo4 = tv_cronograma_y_detalle.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_y_detalle.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "

    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_y_detalle.cantidad1 , to_cronograma_diario_final.cantidad2 = tv_cronograma_y_detalle.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_y_detalle.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_y_detalle.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_y_detalle.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "
    
'    'Quita Cantidad de Insumo3 e Insumo4 en meses pares al Crono Final
'    sino = MsgBox("Elija SI: para programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC) los insumos 3 y 4..." & vbCr & _
'             "Elija NO: para programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV) los insumos 3 y 4....", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
    If Option10.Value = True Then
        'Programar Meses IMPARES y quitar PARES
        db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0', to_cronograma_diario_final.cantidad4 = '0' From to_cronograma_diario_final INNER JOIN tv_cronograma_mensual_par ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_mensual_par.fmes_plan ) " & _
        " where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & " ) "
    Else
        'PROGRAMAR en Meses PARES y quitar Mes IMPARES
        db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0', to_cronograma_diario_final.cantidad4 = '0' From to_cronograma_diario_final INNER JOIN tv_cronograma_mensual_impar ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_mensual_impar.fmes_plan ) " & _
        " where (to_cronograma_diario_final.fmes_plan >= " & VAR_AUX2 & "  ) "
    End If
    
    'Actualiza Cantidad de Insumos al Crono Final Bmes, Tmes, etc.
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and tv_cronograma_insumos.unimed_codigo <> 'MES' "
    
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_carta.carta " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_carta ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_carta.bien_codigo) WHERE to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final.bien_codigo <> '' "
    
    db.Execute " update to_cronograma_diario_final set to_cronograma_diario_final.carta = tv_cronograma_y_detalle.carta from to_cronograma_diario_final inner join tv_cronograma_y_detalle on to_cronograma_diario_final.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_final.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final.bien_codigo <> '' "
    
    MsgBox "Se actualizaron los Insumos desde CRONOGRAMA POR CONTRATO correspondientes a la misma Gestión y Zona del CRONOGRAMA FINAL (DESTINO) ...", vbInformation, "Información"
End Sub

Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
'    'Actualiza Codigos de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Actualiza Cantidad de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
'    'Actualiza Carta al Crono Final
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_insumos.carta " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    
    'to_cronograma_diario_final
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select distinct bien_codigo  from to_cronograma_diario_final where fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and bien_codigo <>'' ", db, adOpenStatic
    If rs_datos1.RecordCount > 0 Then
        VAR_REG = rs_datos1.RecordCount
        VAR_CANT1 = rs_datos1.RecordCount
        'Actualiza Carta al Crono Final
'        db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.carta  = tv_cronograma_carta.carta " & _
'        " From to_cronograma_diario_final INNER JOIN tv_cronograma_carta ON (to_cronograma_diario_final.bien_codigo  = tv_cronograma_carta.bien_codigo)  " & _
'        " WHERE to_cronograma_diario_final.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and to_cronograma_diario_final.bien_codigo <> '' "
    Else
        VAR_REG = "0"
        VAR_CANT1 = "0"
    End If
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual_eqp.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN", "DMANS", "DMANB", "DMANC"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      'Cmb_Mes.Text = "ENERO"
      CR01.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR01.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      CR01.Formulas(3) = "TotalReg = " & VAR_REG & " "
      CR01.Formulas(4) = "CANT1 = " & VAR_CANT1 & " "
      
     CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!fmes_plan
     CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!zpiloto_codigo
     
'    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
'    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
'    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    'db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    '" where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual_origen.rpt"
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      'Cmb_Mes.Text = "ENERO"
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR02.Formulas(2) = "periodo = '" & Cmb_Mes & "' "

    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
    CR02.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir3_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    'db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    '" where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_zonas_vs_edificios.rpt"
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      'Cmb_Mes.Text = "ENERO"
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR02.Formulas(2) = "periodo = '" & Cmb_Mes & "' "

    CR02.StoredProcParam(0) = VAR_FMES
      
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized

End Sub

Private Sub BtnModDetalle_Click()
'    If Ado_detalle2.Recordset("estado_activo") = "ANL" Then  '<> "REG"
'      sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        Ado_detalle2.Recordset!estado_activo = "REG"
'        Ado_detalle2.Recordset!observaciones = "HORARIO LABORABLE"
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'      End If
'   Else
'        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
'   End If
'    Call BtnAñadir2_Click
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    FraDet4.Visible = True
    txt_obs.Visible = False
    Frame2.Visible = False
    BtnGraba4.Visible = False
    BtnCancelar4.Visible = False
    fraOpciones.Enabled = False
    FrmABMDet.Enabled = False
    fraOpciones2.Enabled = False
    Select Case Ado_detalle2.Recordset!estado_activo
        Case "APP"
            cmd_campo1.Text = "HORARIO POR CONFIRMAR"
        Case "APC"
            cmd_campo1.Text = "COMPENSACION"
        Case "APR"
            cmd_campo1.Text = "HORARIO LABORAL Confirmado"
        Case Else
            cmd_campo1.Text = ""
    End Select
    'HORARIO LABORAL Confirmado"
    If dtc_codigo6.Text <> "4211" Then
        dtc_codigo6.Text = "4211"                   'TRAPO
        dtc_desc6.BoundText = dtc_codigo6.BoundText
    End If
    If dtc_codigo6A.Text <> "479" Then
        dtc_codigo6A.Text = "479"                   'GASOLINA
        dtc_desc6A.BoundText = dtc_codigo6A.BoundText
    End If
    If dtc_codigo6B.Text <> "500" Then          '3410003 (ANTES)
        dtc_codigo6B.Text = "500"                   'ACEITE PREPARADO
        dtc_desc6B.BoundText = dtc_codigo6B.BoundText
    End If
    If dtc_codigo6C.Text <> "4529" Then
        dtc_codigo6C.Text = "4529"                  'ACEITE DELGADO 20/50
        dtc_desc6C.BoundText = dtc_codigo6C.BoundText
    End If
    If dtc_codigo6D.Text <> "3113" Then
        dtc_codigo6D.Text = "3113"                  'GRASA PARA RODAMIENTO
        dtc_desc6D.BoundText = dtc_codigo6D.BoundText
    End If
  
  Else
      MsgBox "No se puede MODIFICAR un cronograma APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    'to_cronograma_diario_final
    FraDet6.Visible = True
'    Set rs_aux6 = New ADODB.Recordset
'    If rs_aux6.State = 1 Then rs_aux6.Close
'    rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
'    If rs_aux6.RecordCount > 0 Then
'        MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), debe deshabilitarlos (Retornar) o utilizar el botón (Envia Uno) ...", vbExclamation, "Validación de Registro"
'    Else
'      sino = MsgBox("Está Seguro de ENVIAR TODO el Cronograma ORIGEN al DESTINO ?." & vbCrLf & " SI-->(Envía solo a los Horarios Laborales definidos en el Destino) " & vbCrLf & " NO-->(Envía todo a todos los días calendario, incluyendo días NO laborales) " & vbCrLf & " Cancelar, la Operación", vbYesNoCancel + vbQuestion, "Atención")
'      If sino = vbYes Then
'        Call COPIA_TODOS
'        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'      Else
'        If sino = vbNo Then
'            db.Execute "UPDATE to_cronograma_diario_final SET to_cronograma_diario_final.bien_orden  = to_cronograma_diario_inst.bien_orden, to_cronograma_diario_final.bien_codigo = to_cronograma_diario_inst.bien_codigo, to_cronograma_diario_final.unidad_codigo_tec = to_cronograma_diario_inst.unidad_codigo_tec, " & _
'            " to_cronograma_diario_final.tec_plan_codigo = to_cronograma_diario_inst.tec_plan_codigo, to_cronograma_diario_final.edif_descripcion = to_cronograma_diario_inst.edif_descripcion, to_cronograma_diario_final.estado_activo = 'APR' FROM to_cronograma_diario_final INNER JOIN to_cronograma_diario_inst " & _
'            " ON to_cronograma_diario_final.fmes_plan  = to_cronograma_diario_inst.fmes_plan AND to_cronograma_diario_final.dia_correl  = to_cronograma_diario_inst.dia_correl AND to_cronograma_diario_final.horario_codigo = to_cronograma_diario_inst.horario_codigo WHERE to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'
'            db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'        End If
'      End If
'        'Call BtnAñadir2_Click
'      Call ABRIR_TABLA_DET
'    End If
  Else
      MsgBox "No se puede ENVIAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Visible = True
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        'tc_zonas_piloto
        Set rs_aux4 = New ADODB.Recordset
        If rs_aux4.State = 1 Then rs_aux4.Close
        rs_aux4.Open "Select * from tc_zonas_piloto where zpiloto_codigo = " & dtc_codigo3.Text & " ", db, adOpenStatic
        If rs_aux4.RecordCount > 0 Then
            dtc_codigo4.Text = rs_aux4!beneficiario_codigo
            dtc_desc4.BoundText = dtc_codigo4.BoundText
        End If
    '    BtnVer.Visible = True
    Else
      MsgBox "No se puede MODIFICAR un cronograma APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer_Click()
    'ARREGLO 1
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc11 = dtc_aux41.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc21 = dtc_aux51.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc31 = IIf(IsNull(Ado_datos.Recordset!trafico_c_time_entrada_salida), 0, Ado_datos.Recordset!trafico_c_time_entrada_salida)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod11 = IIf(IsNull(Ado_datos.Recordset!trafico_d_num_paradas_probables), 0, Ado_datos.Recordset!trafico_d_num_paradas_probables)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe11 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_e_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe21 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe31 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre), 0, Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe41 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_entrada_salida), 0, Ado_datos.Recordset!trafico_e_tiempo_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof11 = IIf(IsNull(Ado_datos.Recordset!trafico_f_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_f_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof21 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_f_time_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof31 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_apertura_cierre), 0, Ado_datos.Recordset!trafico_f_time_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof41 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_entrada_salida), 0, Ado_datos.Recordset!trafico_f_time_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog11 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti), 0, Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog21 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_total_arreglo), 0, Ado_datos.Recordset!trafico_g_capacidad_total_arreglo)
    
End Sub

Private Sub BtnVer2_Click()
    Set rs_aux14 = New ADODB.Recordset
    If rs_aux14.State = 1 Then rs_aux14.Close
    rs_aux14.Open "select * from to_cronograma_diario_final where fmes_plan = '" & VAR_FMES & "'  and estado_activo = 'APR' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Sort = "bien_orden"
    If rs_aux14.RecordCount > 0 Then
        MsgBox "No se puede Actualizar el #Horas ni Orden, porque ya existen registros en el Cronograma Final de esta Zona en el Mes a procesar, Vuelva a Intentar ...", vbExclamation, "Validación"
    Else
        db.Execute " update to_cronograma_diario_inst set to_cronograma_diario_inst.nro_total_horas = tv_cronograma_y_detalle.bien_cantidad_por_empaque from to_cronograma_diario_inst inner join tv_cronograma_y_detalle on to_cronograma_diario_inst.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_inst.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " AND tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
        db.Execute " update to_cronograma_diario_inst set to_cronograma_diario_inst.bien_orden = tv_cronograma_y_detalle.zona_edif_orden from to_cronograma_diario_inst inner join tv_cronograma_y_detalle on to_cronograma_diario_inst.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_inst.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " AND tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "

        Call ABRIR_TABLA_DET
        MsgBox "Se Actualizó el <#Horas> por equipo y el <Orden> actual de la Organización de Zonas ...", vbInformation, "Información"
    End If
End Sub

Private Sub cmd_campo1_Click()
    txt_obs.Visible = True
    Frame2.Visible = True
    BtnGraba4.Visible = True
    BtnCancelar4.Visible = True
End Sub

Private Sub cmd_campo1_LostFocus()
    Select Case Trim(cmd_campo1.Text)
        Case "HORARIO POR CONFIRMAR"
            txt_codigo01.Text = "APP"
            If txt_obs.Text <> "" Then
                txt_obs.Text = cmd_campo1.Text + " - " + txt_obs.Text
            Else
                txt_obs.Text = cmd_campo1.Text
            End If
        Case "COMPENSACION"
            txt_codigo01.Text = "APC"
            txt_obs.Text = cmd_campo1.Text
            
        Case "HORARIO LABORAL Confirmado"
            txt_codigo01.Text = "APR"
            If txt_obs.Text = "" Then
                txt_obs.Text = ""
            End If
        Case Else
            txt_codigo01.Text = "REG"
    End Select
'    sino = MsgBox("Reemplazar el texto de Observaciones ?..." & vbCrLf & " SI (Reemplaza) " & vbCrLf & " NO (Aumenta al Texto existente) ", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        txt_obs.Text = cmd_campo1.Text
'    Else
'        txt_obs.Text = cmd_campo1.Text + " - " + txt_obs.Text
'    End If
End Sub

Private Sub dg_datos_ButtonClick(ByVal ColIndex As Integer)
    busca3 = 2
End Sub

Private Sub dg_datos_Click()
    buscados = 0
    busca3 = 2
End Sub

Private Sub dg_datos_GotFocus()
    busca3 = 2
End Sub

Private Sub dg_datos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    busca3 = 2
End Sub

Private Sub dg_det1_DblClick()
    cmd_campo2.Text = "2"
    Call BtnGrabar2_Click
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo6A_Click(Area As Integer)
    dtc_desc6A.BoundText = dtc_codigo6A.BoundText
End Sub

Private Sub dtc_codigo6B_Click(Area As Integer)
    dtc_desc6B.BoundText = dtc_codigo6B.BoundText
End Sub

Private Sub dtc_codigo6C_Click(Area As Integer)
    dtc_desc6C.BoundText = dtc_codigo6C.BoundText
End Sub

Private Sub dtc_codigo6D_Click(Area As Integer)
    dtc_desc6D.BoundText = dtc_codigo6D.BoundText
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

Private Sub dtc_desc6A_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6B_Click(Area As Integer)
    dtc_codigo6B.BoundText = dtc_desc6B.BoundText
End Sub

Private Sub dtc_desc6C_Click(Area As Integer)
    dtc_codigo6C.BoundText = dtc_desc6C.BoundText
End Sub

Private Sub dtc_desc6D_Click(Area As Integer)
    dtc_codigo6D.BoundText = dtc_desc6D.BoundText
End Sub

Private Sub DTPfecha3_LostFocus()
    'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)
    'If Val(Txt_descripcion) < 0 Then
    If Val(DateDiff("y", CDate(DTPfecha2), CDate(DTPfecha3))) < 0 Then
        MsgBox "La Fecha Inicial NO puede ser MAYOR a la Fecha Final, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        DTPfecha3.SetFocus
    End If
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    VAR_SW2 = ""
    busca3 = 0
    cmd_campo2.Text = "2"
    'Fra_Gestion.Visible = True
    VAR_GES = Year(Date)        'Cmb_gestion.Text
    
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux8.RecordCount > 0 Then
        usuario2 = rs_aux8!beneficiario_codigo
        VAR_DA = rs_aux8!da_codigo
        VAR_DPTOC = rs_aux8!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    VAR_UORIGEN = "DNINS"
'    If Aux = "DNINS" Then
'        Select Case VAR_DPTOC
'            Case "1"    ' Chuquisaca
'                VAR_UORIGEN = "DMANC"
'            Case "2"    'La Paz - Tecnico
'                VAR_UORIGEN = "DNMAN"
'            Case "3"    'Cochabamba
'                VAR_UORIGEN = "DMANB"
'                'VAR_DPTOC = "3"
'            Case "7"    'Santa Cruz
'                VAR_UORIGEN = "DMANS"
'                'VAR_DPTOC = "7"
'            Case "4"    'Oruro - Tecnico
'                VAR_UORIGEN = "DNMAN"
'                'VAR_DPTOC = "2"
'            Case "5"    ' Potosi
'                VAR_UORIGEN = "DMANC"
'            Case "6"    ' Tarija
'                VAR_UORIGEN = "DMANC"
'            Case "8"    ' Beni
'                VAR_UORIGEN = "DMANC"
'            Case "9"    ' Pando
'                VAR_UORIGEN = "DMANC"
'            Case Else    ' TODO
'                VAR_UORIGEN = "DNMAN"
'                VAR_DPTOC = "0"
'         End Select
'     End If
    
    parametro = Aux
    VAR_ANL = ""
'    If glusuario = "MLLOSA" Then
'        'fraOpciones.Enabled = False
'        'FrmABMDet.Enabled = False
'        'Picture1.Enabled = False
'        BtnModificar.Visible = False
'        BtnEliminar.Visible = False
'        BtnAprobar.Visible = False
'        BtnAnlDetalle4.Visible = False
'        BtnModDetalle.Visible = False
'        BtnAnlDetalle.Visible = False
'        BtnAñadir2.Visible = False
'        BtnGraba4.Visible = False
'        BtnGrabar2.Visible = False
'        BtnGraba3.Visible = False
'        BtnAddDetalle.Visible = False
'        BtnAnlDetalle2.Visible = False
'        BtnAddDetalle3.Visible = False
'        BtnModDetalle2.Visible = False
'        BtnAnlDetalle3.Visible = False
'    End If
    
'    ' Actualiza Nombre de Edificios
'    db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.edif_descripcion   = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_inst INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_inst.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
'    'Actualiza Zona Geografica
'    db.Execute "update tc_zona_piloto_edif set tc_zona_piloto_edif.zona_codigo  = gc_edificaciones.zona_codigo from tc_zona_piloto_edif inner join gc_edificaciones on tc_zona_piloto_edif.edif_codigo   = gc_edificaciones.edif_codigo"
'    'Actualiza Orden Edificio a Cronograma Diario por equipo
'    db.Execute "update ac_bienes set ac_bienes.kit  = '0' where par_codigo = '43340'"
'    db.Execute "update ac_bienes set ac_bienes.kit  = tc_zona_piloto_edif.zona_edif_orden, ac_bienes.observaciones = tc_zona_piloto_edif.observaciones from ac_bienes inner join tc_zona_piloto_edif on ac_bienes.edif_codigo  = tc_zona_piloto_edif.edif_codigo"
'
'    db.Execute "update to_cronograma_diario_inst set bien_orden = '0' where bien_orden <> '0' "
'    db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.bien_orden = ac_bienes.kit, to_cronograma_diario_inst.observaciones   = ac_bienes.observaciones from to_cronograma_diario_inst inner join ac_bienes on to_cronograma_diario_inst.bien_codigo   = ac_bienes.bien_codigo "
'    'Actualiza Observaciones de Organizacion de Zonas vs. Edificios
    'Actualiza Responsables de Zona
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    Option1.Value = True
    Option3.Value = True
    buscados = 0
    'lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
   'If Not Ado_datos.Recordset.EOF Then
            'SSTab1.Tab = 0
            'SSTab1.TabEnabled(0) = True
            ''SSTab1.TabEnabled(1) = False
            'SSTab1.TabVisible(1) = False
   'End If
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
        
    'tc_zonas_piloto
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & VAR_UORIGEN & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'INSUMOS
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select distinct * from av_bienes_vs_venta_detalle where par_codigo = '33100' or par_codigo = '34110' ORDER BY bien_descripcion ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
End Sub

'Private Sub pnivel1(codigo1 As String)
''   Dim strConsultaF As String
''   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'
'   Set dtc_codigo10.RowSource = Nothing
''   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo10.ReFill
'   dtc_codigo10.BoundText = Empty
'
'   Set dtc_desc10.RowSource = Nothing
'   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc10.ReFill
'   dtc_desc10.BoundText = Empty
'End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub OptFilGral0_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2019)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2019' ) "
        Case "2"    'La Paz - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2019' ) "
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2019' ) "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2019' ) "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2019' ) "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2019' ) "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2019' ) "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2019' ) "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2019' ) "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual where ( ges_gestion = '2019' ) "
     End Select
    'queryinicial = "Select * from to_cronograma_mensual "          'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2020)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2020' ) "
        Case "2"    'La Paz - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2020' ) "
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2020' ) "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2020' ) "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2020' ) "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2020' ) "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2020' ) "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2020' ) "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2020' ) "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual where ( ges_gestion = '2020' ) "
     End Select

    'queryinicial = "Select * from to_cronograma_mensual "          'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2021)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From to_cronograma_mensual_inst  "      'WHERE (ges_gestion = '2022' )
'    Select Case VAR_DPTOC
'        Case "1"    ' Chuquisaca
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2022' ) "
'        Case "2"    'La Paz - Tecnico
'            If glusuario = "OCOLODRO" Then
'                queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2022' ) "
'            Else
'                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' OR zpiloto_codigo='39' OR zpiloto_codigo='40')  AND ges_gestion = '2022' ) "
'            End If
'        Case "3"    'Cochabamba
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2022' ) "
'        Case "7"    'Santa Cruz
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2022' ) "
'        Case "4"    'Oruro - Tecnico
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2022' ) "
'        Case "5"    ' Potosi
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2022' ) "
'        Case "6"    ' Tarija
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2022' ) "
'        Case "8"    ' Beni
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2022' ) "
'        Case "9"    ' Pando
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2022' ) "
'        Case Else    ' TODO
'            queryinicial = "select * From to_cronograma_mensual where ( ges_gestion = '2022' ) "
'     End Select

    rs_datos.Sort = "ges_gestion, fmes_correl, zpiloto_codigo"
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from ao_solicitud_cotiza_venta where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
        
'    dtc_desc31.BoundText = dtc_codigo31.BoundText
'    dtc_desc32.BoundText = dtc_codigo31.BoundText
'    dtc_desc33.BoundText = dtc_codigo31.BoundText
'    dtc_desc34.BoundText = dtc_codigo31.BoundText
'
'    dtc_desc41.BoundText = dtc_codigo41.BoundText
'    dtc_desc42.BoundText = dtc_codigo41.BoundText
'    dtc_desc43.BoundText = dtc_codigo41.BoundText
'    dtc_desc44.BoundText = dtc_codigo41.BoundText
'
'    dtc_desc51.BoundText = dtc_codigo51.BoundText
'    dtc_desc52.BoundText = dtc_codigo51.BoundText
'    dtc_desc53.BoundText = dtc_codigo51.BoundText
'    dtc_desc54.BoundText = dtc_codigo51.BoundText
End Sub

'Private Sub Img_03_Click()
' If AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo asociado al Registro, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'   If GlServidor = "SRVPRO" Then
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   Else
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   End If
' End If
'
'End Sub

'Private Sub Img_CTO_Click()
' If Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo Asociado al Contrato, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SRVPRO" Then
'        'e = ShellExecute(Img_CTO, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    Else
'        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    End If
' End If
'End Sub

'Private Sub Img_CV_Click()
''    Dim e As Long
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_HOJAVIDA = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "C_V"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'         ' e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.AdoMovilidad.Recordset!solicitud_codigo) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "C_V"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'  End If
'  If GlServidor = "SRVPRO" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  End If
'End Sub
'
'Private Sub Img_Foto_Click()
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FOT"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FOT"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'
'    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
'        ARCH_FOTO = App.Path + "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    End If
'    If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where solicitud_codigo= '" & Ado_datos.Recordset("solicitud_codigo") & "' ", "Foto", ARCH_FOTO) Then
'        MsgBox "Se cargo la Imagen Correctamente !!"
'    Else
'        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'    End If
'  End If
'End Sub

'Private Sub SSTab1_DblClick()
'    If SSTab1.Tab = 0 Then
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  If glPersNew = "P" Then
  End If
  glPersNew = "N"
'   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub ABRIR_TABLA_DET()
On Error GoTo UpdateErr
    If Option3.Value = True Then
        Set rs_det1 = New ADODB.Recordset
        If rs_det1.State = 1 Then rs_det1.Close
        rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & VAR_FMES & "'  and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Sort = "bien_orden"
        Set Ado_detalle1.Recordset = rs_det1
        If Ado_detalle1.Recordset.RecordCount > 0 Then
            Set dg_det1.DataSource = Ado_detalle1.Recordset
        Else
            Set dg_det1.DataSource = rsNada
        End If
    End If
    If Option4.Value = True Then
        Set rs_det1 = New ADODB.Recordset
        If rs_det1.State = 1 Then rs_det1.Close
        'rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "' and estado_activo <> 'ANL' AND estado_activo <> 'APR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & VAR_FMES & "' and estado_codigo =  'REG' AND bien_codigo <> ''  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Sort = "bien_orden"
        Set Ado_detalle1.Recordset = rs_det1
        If Ado_detalle1.Recordset.RecordCount > 0 Then
            Set dg_det1.DataSource = Ado_detalle1.Recordset
        Else
            Set dg_det1.DataSource = rsNada
        End If
    End If
    If Option1.Value = True Then
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & VAR_FMES & "' and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Sort = "bien_orden"
        Set Ado_detalle2.Recordset = rs_det2
        If Ado_detalle2.Recordset.RecordCount > 0 Then
            Ado_detalle2.Recordset.MoveLast
            Set dg_det2.DataSource = Ado_detalle2.Recordset
            dg_det2.Visible = True
        Else
            Set dg_det2.DataSource = rsNada
            dg_det2.Visible = False
        End If
    End If
    If Option2.Value = True Then
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & VAR_FMES & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Sort = "bien_orden"
        Set Ado_detalle2.Recordset = rs_det2
        If Ado_detalle2.Recordset.RecordCount > 0 Then
            Ado_detalle2.Recordset.MoveLast
            Set dg_det2.DataSource = Ado_detalle2.Recordset
            dg_det2.Visible = True
        Else
            Set dg_det2.DataSource = rsNada
            dg_det2.Visible = False
        End If
    End If
    If Option5.Value = True Then
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & VAR_FMES & "' and estado_activo <> 'ANL' AND estado_activo <> 'APR' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Sort = "bien_orden"
        Set Ado_detalle2.Recordset = rs_det2
        If Ado_detalle2.Recordset.RecordCount > 0 Then
            Ado_detalle2.Recordset.MoveLast
            Set dg_det2.DataSource = Ado_detalle2.Recordset
            dg_det2.Visible = True
        Else
            Set dg_det2.DataSource = rsNada
            dg_det2.Visible = False
        End If
    End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub OptFilGral3_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2020)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2021' ) "
        Case "2"    'La Paz - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2021' ) "
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2021' ) "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2021' ) "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2021' ) "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2021' ) "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2021' ) "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2021' ) "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2021' ) "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual where ( ges_gestion = '2021' ) "
     End Select

    'queryinicial = "Select * from to_cronograma_mensual "          'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub Option1_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option2_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option3_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option4_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option5_Click()
    Call ABRIR_TABLA_DET
End Sub



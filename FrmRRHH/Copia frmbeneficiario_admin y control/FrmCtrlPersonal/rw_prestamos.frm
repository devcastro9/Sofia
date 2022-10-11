VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rw_prestamos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Tecnico - Venta de Servicios"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6.1193e6
   ScaleMode       =   0  'User
   ScaleWidth      =   9.33641e9
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_crono 
      BackColor       =   &H80000009&
      Caption         =   "Cronograma de Pagos"
      ForeColor       =   &H00FFFF00&
      Height          =   4335
      Left            =   6720
      TabIndex        =   80
      Top             =   960
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox Text17 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   280
         Left            =   1800
         TabIndex        =   111
         Top             =   1240
         Width           =   270
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H00FFFFFF&
         DataField       =   "cobranza_observaciones"
         DataSource      =   "Ado_datos1"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   109
         Top             =   2715
         Width           =   9915
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   $"rw_prestamos.frx":0000
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
         Height          =   735
         Left            =   120
         TabIndex        =   100
         Top             =   3480
         Width           =   11295
         Begin VB.TextBox Text14 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   7390
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "estado_codigo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "total_cobrado"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   103
            Text            =   "0"
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "correl_prog"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   9885
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   300
            Width           =   1065
         End
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "prestamo_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   101
            Text            =   "0"
            Top             =   300
            Width           =   1665
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   2955
            X2              =   2955
            Y1              =   840
            Y2              =   120
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Meses"
            DataField       =   "unimed_codigo"
            DataSource      =   "Ado_datos"
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
            Left            =   1440
            TabIndex        =   108
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   7005
            TabIndex        =   107
            Top             =   345
            Width           =   405
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   4800
            TabIndex        =   106
            Top             =   360
            Width           =   405
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFF80&
            X1              =   9480
            X2              =   9480
            Y1              =   840
            Y2              =   120
         End
      End
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
         ScaleWidth      =   11280
         TabIndex        =   96
         Top             =   240
         Width           =   11280
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5955
            Picture         =   "rw_prestamos.frx":0087
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   98
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4440
            Picture         =   "rw_prestamos.frx":0973
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   97
            Top             =   0
            Width           =   1280
         End
         Begin VB.Label Label7 
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
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   13095
            TabIndex        =   99
            Top             =   180
            Width           =   1005
         End
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "cobranza_programada_bs"
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
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4110
         TabIndex        =   94
         Text            =   "0"
         Top             =   2040
         Width           =   1500
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "fecha_pla"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos1"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   2400
         TabIndex        =   92
         Text            =   "0"
         Top             =   2040
         Width           =   1260
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "prestamo_prog_codigo"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos1"
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Left            =   960
         TabIndex        =   90
         Text            =   "0"
         Top             =   2040
         Width           =   1140
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   280
         Left            =   6720
         TabIndex        =   89
         Top             =   1245
         Width           =   270
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "unidad_codigo_ant"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7245
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "0"
         Top             =   1230
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "rw_prestamos.frx":1149
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   840
         TabIndex        =   81
         Top             =   1230
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "rrhh_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "rw_prestamos.frx":1162
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   2280
         TabIndex        =   83
         Top             =   1230
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "No. Cuota"
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
         Index           =   2
         Left            =   960
         TabIndex        =   91
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
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
         Left            =   960
         TabIndex        =   110
         Top             =   2400
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Monto Programado Bs"
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
         Index           =   5
         Left            =   3855
         TabIndex        =   95
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Mes PLanilla"
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
         Index           =   4
         Left            =   2370
         TabIndex        =   93
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "prestamo_codigo"
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
         Left            =   9360
         TabIndex        =   88
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correl Prestamo"
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
         Left            =   9360
         TabIndex        =   87
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite Trámite"
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
         Left            =   7320
         TabIndex        =   86
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   2325
         TabIndex        =   84
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Trámite"
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
         Index           =   1
         Left            =   900
         TabIndex        =   82
         Top             =   960
         Width           =   1110
      End
   End
   Begin VB.Frame FrmCabecera 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5070
      Left            =   6720
      TabIndex        =   39
      Top             =   720
      Width           =   11535
      Begin VB.TextBox Text3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   280
         Left            =   1080
         TabIndex        =   75
         Top             =   475
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   280
         Left            =   10920
         TabIndex        =   74
         Top             =   1215
         Width           =   270
      End
      Begin VB.TextBox Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "unidad_codigo_ant"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   466
         Width           =   1935
      End
      Begin VB.Frame Fra_Total 
         BackColor       =   &H00000000&
         Caption         =   $"rw_prestamos.frx":117C
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
         Height          =   735
         Left            =   60
         TabIndex        =   53
         Top             =   4140
         Width           =   11415
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "prestamo_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   72
            Text            =   "0"
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox TxtBstotal 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "correl_prog"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   9885
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   300
            Width           =   1065
         End
         Begin VB.TextBox TxtMontoBs 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "total_cobrado"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "0"
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox txtCantTotal 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            DataField       =   "estado_codigo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   300
            Width           =   975
         End
         Begin VB.TextBox TxtCobrado 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   7390
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   300
            Width           =   1665
         End
         Begin VB.TextBox txtTDC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            DataField       =   "venta_tipo_cambio"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   -240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   9480
            X2              =   9480
            Y1              =   840
            Y2              =   120
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   4800
            TabIndex        =   77
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   7005
            TabIndex        =   60
            Top             =   345
            Width           =   405
         End
         Begin VB.Label lbl_unimed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Meses"
            DataField       =   "unimed_codigo"
            DataSource      =   "Ado_datos"
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
            Left            =   1440
            TabIndex        =   59
            Top             =   300
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF80&
            X1              =   2955
            X2              =   2955
            Y1              =   840
            Y2              =   120
         End
      End
      Begin VB.Frame Fra_datos 
         BackColor       =   &H00000000&
         Caption         =   $"rw_prestamos.frx":1203
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
         Height          =   2325
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   11295
         Begin VB.TextBox txt_mes 
            BackColor       =   &H00000000&
            DataField       =   "mes_grupo"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFF00&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   78
            Text            =   "0"
            Top             =   1920
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txt_interes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9480
            TabIndex        =   4
            Text            =   "0"
            Top             =   270
            Width           =   540
         End
         Begin VB.TextBox txt_monto_bs 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DataField       =   "prestamo_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7365
            TabIndex        =   3
            Text            =   "0"
            Top             =   270
            Width           =   1140
         End
         Begin VB.TextBox TxtConcepto 
            BackColor       =   &H00404040&
            DataField       =   "prestamos_descripcion"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   525
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   960
            Width           =   6795
         End
         Begin VB.ComboBox cmb_mes_ini 
            DataField       =   "mes_inicio_crono"
            DataSource      =   "Ado_datos"
            Height          =   315
            ItemData        =   "rw_prestamos.frx":1295
            Left            =   840
            List            =   "rw_prestamos.frx":12BD
            TabIndex        =   6
            Top             =   1920
            Width           =   2460
         End
         Begin VB.TextBox txtCantCobr 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DataField       =   "cantidad_periodos"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5565
            TabIndex        =   7
            Text            =   "0"
            Top             =   1920
            Width           =   1140
         End
         Begin VB.ComboBox cmd_unimed2 
            DataField       =   "unimed_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            ItemData        =   "rw_prestamos.frx":1326
            Left            =   8520
            List            =   "rw_prestamos.frx":1336
            TabIndex        =   8
            Text            =   "MENSUAL"
            Top             =   1920
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker DTPFechaIni 
            Height          =   285
            Left            =   7440
            TabIndex        =   5
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Format          =   89325569
            CurrentDate     =   41791
            MaxDate         =   401768
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "rw_prestamos.frx":1360
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "rw_prestamos.frx":1379
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   45
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker DTPFechaFin 
            Height          =   285
            Left            =   9480
            TabIndex        =   46
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   89325569
            CurrentDate     =   41791
            MinDate         =   36526
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "%"
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
            Left            =   10005
            TabIndex        =   79
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
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
            Left            =   240
            TabIndex        =   52
            Top             =   645
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha De Prestamo"
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
            Left            =   7245
            TabIndex        =   51
            Top             =   765
            Width           =   1800
         End
         Begin VB.Label Label21 
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Mes de Inicio de la Cobranza Segun Planilla"
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
            Height          =   480
            Left            =   240
            TabIndex        =   50
            Top             =   1560
            Width           =   5010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin"
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
            Left            =   9720
            TabIndex        =   49
            Top             =   765
            Width           =   885
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Número de Periodos"
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
            Left            =   5280
            TabIndex        =   48
            Top             =   1575
            Width           =   1890
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
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
            Left            =   8505
            TabIndex        =   47
            Top             =   1575
            Width           =   1425
         End
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   280
         Left            =   7395
         TabIndex        =   41
         Top             =   480
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "rw_prestamos.frx":1392
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6105
         TabIndex        =   42
         Top             =   1215
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "rw_prestamos.frx":13AB
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   1215
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "rw_prestamos.frx":13C4
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   6900
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Dtc_aux2 
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6120
         TabIndex        =   63
         Top             =   2040
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483632
         ForeColor       =   -2147483624
         ListField       =   "codigo2"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "rw_prestamos.frx":13DE
         DataField       =   "unidad_codigo"
         Height          =   315
         Left            =   1740
         TabIndex        =   62
         Top             =   465
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_haber_men 
         Bindings        =   "rw_prestamos.frx":13F8
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9480
         TabIndex        =   73
         Top             =   1200
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_haber_mensual"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_sol_cod 
         Bindings        =   "rw_prestamos.frx":1411
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   76
         Top             =   465
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "rrhh_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Persona que solicita el prestamo"
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
         Left            =   180
         TabIndex        =   70
         Top             =   885
         Width           =   2940
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correl Prestamo"
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
         Left            =   9900
         TabIndex        =   69
         Top             =   195
         Width           =   1485
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   1785
         TabIndex        =   68
         Top             =   195
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Trámite"
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
         Index           =   0
         Left            =   180
         TabIndex        =   67
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Haber Básico"
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
         Index           =   3
         Left            =   8145
         TabIndex        =   66
         Top             =   1245
         Width           =   1260
      End
      Begin VB.Label txt_venta 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "prestamo_codigo"
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
         Left            =   9900
         TabIndex        =   65
         Top             =   465
         Width           =   1485
      End
      Begin VB.Label lbl_cite 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite Trámite"
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
         Left            =   7755
         TabIndex        =   64
         Top             =   195
         Width           =   1965
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      Picture         =   "rw_prestamos.frx":142A
      ScaleHeight     =   2235
      ScaleWidth      =   1875
      TabIndex        =   33
      Top             =   5970
      Width           =   1935
      Begin VB.CommandButton BtnanularDetalle2 
         BackColor       =   &H00808080&
         Height          =   525
         Left            =   240
         Picture         =   "rw_prestamos.frx":6D45C
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Modifica la Cobranza Identifiacada"
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CommandButton BtnAñadirDetalle2 
         BackColor       =   &H00808080&
         Height          =   525
         Left            =   240
         MaskColor       =   &H00404040&
         Picture         =   "rw_prestamos.frx":6DBA8
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Aprueba Registro Identificado"
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   240
         Picture         =   "rw_prestamos.frx":6E367
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cronograma para Contrato"
         Top             =   2280
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo->"
         Height          =   640
         Left            =   360
         Picture         =   "rw_prestamos.frx":6EC34
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Registra una Nueva Cobranza"
         Top             =   165
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   240
         Picture         =   "rw_prestamos.frx":6F076
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Modifica la Cobranza Identifiacada"
         Top             =   840
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   285
         Picture         =   "rw_prestamos.frx":6F98B
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   2355
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton BtnAprobar2 
         BackColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   240
         Picture         =   "rw_prestamos.frx":700D7
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Aprueba Registro Identificado"
         Top             =   2085
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20400
      TabIndex        =   17
      Top             =   0
      Width           =   20400
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "rw_prestamos.frx":7090A
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   71
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         Appearance      =   0  'Flat
         Caption         =   "Digitaliza"
         Height          =   710
         Left            =   11640
         Picture         =   "rw_prestamos.frx":710C9
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnDesAprobar 
         Appearance      =   0  'Flat
         Height          =   710
         Left            =   11280
         Picture         =   "rw_prestamos.frx":7150B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1425
         Picture         =   "rw_prestamos.frx":71715
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   25
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "rw_prestamos.frx":7202A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   24
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "rw_prestamos.frx":72776
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   23
         ToolTipText     =   "Aprueba el Registro Elegido"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "rw_prestamos.frx":72FA9
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   22
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "rw_prestamos.frx":7375E
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   21
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17640
         Picture         =   "rw_prestamos.frx":7402B
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   20
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnVer3 
         BackColor       =   &H80000015&
         Caption         =   "Tramite Provisional"
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
         Left            =   9480
         Picture         =   "rw_prestamos.frx":747ED
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cerrar Tramite y Archivarlo"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer2 
         BackColor       =   &H80000015&
         Caption         =   "Alcance"
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
         Left            =   10440
         Picture         =   "rw_prestamos.frx":74D77
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Regitra Alcance del Contrato"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   12975
         TabIndex        =   28
         Top             =   180
         Width           =   1815
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
      ScaleWidth      =   20400
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   20400
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "rw_prestamos.frx":75A41
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   31
         Top             =   0
         Width           =   1280
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "rw_prestamos.frx":76217
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   30
         Top             =   0
         Width           =   1455
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13095
         TabIndex        =   32
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTA"
      ForeColor       =   &H00FFFFC0&
      Height          =   5040
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   6465
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3600
         TabIndex        =   16
         Top             =   4635
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pendientes"
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
         TabIndex        =   15
         Top             =   4635
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   4290
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   7567
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   16777152
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "prestamo_codigo"
            Caption         =   "Prestamo"
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
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "C.I."
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
            DataField       =   "beneficiario_denominacion_resp"
            Caption         =   "Nombre Completo"
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
            DataField       =   "prestamo_total_bs"
            Caption         =   "Monto Bs"
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
         BeginProperty Column05 
            DataField       =   "fecha_inicio"
            Caption         =   "Fecha Prestamo"
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
            DataField       =   "fecha_fin"
            Caption         =   "Fecha Fin"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3089.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   780.095
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4560
         Width           =   6225
         _ExtentX        =   10980
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
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00000000&
      Caption         =   "CRONOGRAMA PARA CONTROL DEL PRESTAMO"
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
      Height          =   2385
      Left            =   2160
      TabIndex        =   11
      Top             =   5880
      Width           =   16095
      Begin VB.Frame fra_alerta 
         Caption         =   "AVISO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1935
         Left            =   12840
         TabIndex        =   113
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Label lbl_alerta 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Aviso"
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
            Height          =   1605
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   2805
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   1980
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   15855
         _ExtentX        =   27966
         _ExtentY        =   3493
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "prestamo_prog_codigo"
            Caption         =   "No. Cuota"
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
            DataField       =   "gestion"
            Caption         =   "Gestón"
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
            DataField       =   "fecha_pla"
            Caption         =   "Mes Planilla "
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
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto Programado Bs."
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
         BeginProperty Column04 
            DataField       =   "cobranza_concepto_plazo"
            Caption         =   "Plazo a Cumplir"
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Observacioon"
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
         BeginProperty Column07 
            DataField       =   "cobranza_programada_dol"
            Caption         =   "Monto a Pagar Dol."
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   2819.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   7409.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1470.047
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   0
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
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2160
      Top             =   9960
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   11280
      Top             =   10320
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
      Height          =   450
      Left            =   9000
      Top             =   10320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   794
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
      Top             =   10320
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
      Left            =   13560
      Top             =   10320
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
      Left            =   6720
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   11280
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2160
      Top             =   10320
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
      Left            =   4440
      Top             =   10320
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
      Left            =   13560
      Top             =   9960
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
      Left            =   4440
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9960
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
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9000
      Top             =   9960
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
      Left            =   480
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
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "rw_prestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Ventas

Dim rs_datos As New ADODB.Recordset     'VENTAS
Dim rs_datos1 As New ADODB.Recordset    'UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'Beneficiario Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset
Dim calculo As Double
Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
'AUXILIARES
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

Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
'OTROS
'Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim monto_cuota As Double


Dim msgSalir, ado, modif, modif_det As String
Dim queryinicial As String
Dim queryinicial2 As String
'Almacenes
Dim descri_bien As String
Dim Cant_Alm, VAR_CANT, dia, mes As Integer
Dim correlativo1, mes_grupo, gestion As Integer
'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2, CONT_MED As Integer
Dim nroventa, correlv, correldet2, corrprog As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SOL, CANTOT, var_cod5 As Integer
Dim CONT2, CONT3, CONT4, VAR_TIPO As Integer
Dim CUOTA, fdia, fmes, fanio, Dias_Mes, TimeD As Integer
Dim VAR_COBR1, VAR_COBR2, VAR_CONTR As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2, TOTAL_COB As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3, VAR_COD4 As String
Dim VAR_MED, VAR_MED2 As String
Dim VAR_TIPOV, VAR_VAL As String
Dim VAR_FEC2, MControl, VAR_MES2 As String
Dim VAR_BEN2 As String

Dim FInicio, FFin, FControl As Date


Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub


Private Sub calcular_total()
Dim accion As String
accion = "NO"
TOTAL_COB = 0
Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux9.Close
    rs_aux9.Open "select sum(cobranza_programada_bs)as SUMA_CUOTA from ro_prestamo_prog where prestamo_codigo = '" & Ado_datos.Recordset!prestamo_codigo & "' AND beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    
     
     If rs_aux9!SUMA_CUOTA > rs_datos!prestamo_total_bs Then
     TOTAL_COB = rs_aux9!SUMA_CUOTA - rs_datos!prestamo_total_bs
     If TOTAL_COB >= 1 Then
     lbl_alerta.Caption = "La suma de las cuotas es mayor a la deuda" & vbCrLf & "Sobrepasa: " & TOTAL_COB & " Bs."
     accion = "SI"
     End If
     End If
     
     If rs_aux9!SUMA_CUOTA < rs_datos!prestamo_total_bs Then
     TOTAL_COB = rs_datos!prestamo_total_bs - rs_aux9!SUMA_CUOTA
     If TOTAL_COB >= 1 Then
     lbl_alerta.Caption = "La suma de las cuotas es menor a la deuda" & vbCrLf & "Falta: " & TOTAL_COB & " Bs."
     accion = "SI"
     End If
     End If
     
     If accion = "NO" Then
     fra_alerta.Visible = False
     Else
     fra_alerta.Visible = True
     End If
     
End Sub

'Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
'        End Select
'End Sub

Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
        Else
            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            dtccodpar.Text = " "
            dtcdescripar.Text = " "
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
If Ado_datos.Recordset.RecordCount > 0 Then
FrmABMDet2.Enabled = True
If ado = "SI" Then
calculo = 0
calculo = Ado_datos.Recordset!prestamo_total_bs - IIf(IsNull(Ado_datos.Recordset!total_cobrado), "0", Ado_datos.Recordset!total_cobrado)
TxtCobrado.Text = Format(calculo, "###,###,##0.00")
Text14.Text = Format(calculo, "###,###,##0.00")
DTPFechaIni.Value = Ado_datos.Recordset!fecha_inicio
DTPFechaFin.Value = Ado_datos.Recordset!fecha_fin
If Ado_datos.Recordset!intereses_bs > 0 Then
txt_interes.Text = (Ado_datos.Recordset!intereses_bs / Ado_datos.Recordset!prestamo_bs) * 100
Else
txt_interes.Text = "0"
End If
End If
End If
End If
Call ABRIR_TABLA_DET


'Dim descri_bien As String
'Dim Cant_Alm As Integer
'If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
'   If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then
'        If buscados = 0 Then
'           OptFilGral1.Visible = True
'           OptFilGral2.Visible = True
'        Else
'           OptFilGral1.Visible = False
'           OptFilGral2.Visible = False
'        End If
'        If (Ado_datos.Recordset("estado_codigo") = "REG") Then
'            BtnAprobar.Visible = True
''                BtnDesAprobar.Visible = False
'            BtnModificar.Visible = True
'            BtnEliminar.Visible = True
''            BtnVer.Visible = False
'            BtnAñadir.Visible = False   'Cerrar Tramite
'            BtnVer3.Visible = False     'Provisional
'            lbl_cerrado.Caption = ""
'            BtnImprimir2.Visible = True
'            If IsNull(Ado_datos.Recordset("venta_tipo")) Then
'                FrmABMDet.Visible = False
'                FrmABMDet2.Visible = False
'                FrmCobranza.Visible = False
'            Else
'                FrmABMDet.Visible = True
'                FrmABMDet2.Visible = True
'                FrmCobranza.Visible = True
'            End If
'        Else
'            BtnAprobar.Visible = False
''                BtnDesAprobar.Visible = True
'            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
''            BtnVer.Visible = True
'            Select Case Ado_datos.Recordset!estado_cancelado
'                Case "S"
'                    lbl_cerrado.Caption = "TRAMITE CERRADO !!"
'                    FrmABMDet2.Visible = False
'                    BtnAñadir.Visible = False   'Cerrar Tramite
'                    BtnVer3.Visible = False     'Provisional
'                Case "P"
'                    lbl_cerrado.Caption = "TRAMITE PROVISIONAL !!"
'                    FrmABMDet2.Visible = True
'                    BtnAñadir.Visible = False   'Cerrar Tramite
'                    BtnVer3.Visible = False     'Provisional
'                Case Else
'                    BtnAñadir.Visible = True   'Cerrar Tramite
'                    BtnVer3.Visible = True     'Provisional
'                    lbl_cerrado.Caption = ""
'                    FrmABMDet2.Visible = True
'            End Select
'            FrmABMDet.Visible = False
'            FrmCobranza.Visible = True
'            BtnImprimir2.Visible = True
'        End If
'
'        If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Then
''            TxtPlazo.Visible = True
''            BtnAddDetalle2.Visible = True
'        Else
''            TxtPlazo.Visible = False
'            If Ado_datos.Recordset("venta_tipo") = "E" Then
''                BtnAddDetalle2.Visible = False
'            End If
'        End If
'
'        If Dtc_deudor2.Text = "SI" Then
'            Dtc_deudor2.backColor = &HFF&
'        Else
'            Dtc_deudor2.backColor = &H80000010
'        End If
'        'If Ado_datos.Recordset("beneficiario_codigo") <> "" And Ado_datos.Recordset("beneficiario_codigo") <> "VD" Then
'        If Ado_datos.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'            'RS_BENEF.Recordset.Requery
'            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.backColor = &HFF&
'                Else
'                    Dtc_deudor2.backColor = &H80000010
'                End If
'            End If
'        End If
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "' order by  par_codigo, bien_codigo ", db, adOpenKeyset, adLockOptimistic
'        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
'        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'        Set ado_datos14.Recordset = rs_datos14
'        ado_datos14.Recordset.Requery
'        If ado_datos14.Recordset.RecordCount > 0 Then
'            deta2 = 1
'            'TxtMontoBs.Text = Ado_datos.Recordset!monto_total_bS
'            'TxtMontoUs.Text = Ado_datos.Recordset!deuda_cobrada
'            'Text2.Text = Ado_datos.Recordset!saldo_p_cobrar
'            Call AbreAlmacen
''            If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Then
''                FrmABMDet2.Visible = True
''                FrmCobranza.Visible = True
''            Else
''                FrmABMDet2.Visible = False
''                FrmCobranza.Visible = False
''            End If
'        Else
'            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
'        End If
'
'        Set rs_datos16 = New ADODB.Recordset
'        If rs_datos16.State = 1 Then rs_datos16.Close
'        rs_datos16.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos16.Recordset = rs_datos16
'        Ado_datos16.Recordset.Requery
'        If Ado_datos16.Recordset.RecordCount > 0 Then
'            FrmCobranza.Visible = True
'            'BtnImprimir2.Visible = True
'            'BtnImprimir3.Visible = True
'        Else
'            FrmCobranza.Visible = False
'            'BtnImprimir2.Visible = False
'            'BtnImprimir3.Visible = False
'        End If
'
'        FrmDetalle.Caption = "BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
'        FrmCobranza.Caption = "CRONOGRAMA DE COBRANZAS - CONTRATO NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
''        Else
''            ' por si es nuevo
''            dtccodpoa.Text = " "
''            dtcdespoa.Text = dtccodpoa.BoundText
''            dtc_codigo4.Text = " "
''            Dtcpaternosol.Text = dtc_codigo4.BoundText
''            dtcmaternosol.Text = " "
''            dtcnombresol.Text = " "
''            dtccodpuesto.Text = " "
''            dtcdenopuesto.Text = dtccodpuesto.BoundText
''            dtccoduni.Text = " "
''            dtcdescripuni.Text = dtccoduni.BoundText
''            dtc_codigo15.Text = " "
''            dtc_desc15.Text = " "
''            TxtMonto_bolivianos.Text = 0
''            Txtobservaciones.Text = ""
''            Txtcaracteristicas.Text = ""
''            txtsolpeso.Text = 0
'        End If
'        FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
'    Else
'        'FrmABMDet.Visible = False
'        FrmABMDet2.Visible = False
'        'FrmCabecera.Enabled = True
'        'FrmDetalle.Visible = False
'        FrmCobranza.Visible = False
'        If buscados = 0 Then
'           OptFilGral1.Visible = True
'           OptFilGral2.Visible = True
'        Else
'           OptFilGral1.Visible = False
'           OptFilGral2.Visible = False
'        End If
'    End If

  '<-- Inicio                Identificación del Cliente                Fin -->
'   If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
'     'If VAR_SW = "NO" Then
''
'        If Ado_datos.Recordset.RecordCount > 0 Then
'            Call ABRIR_TABLA_DET
'            'Call ABRIR_TABLAS_AUX
'            busq = 0
'
'        End If
'        VAR_SW = ""
'    Else
'        VAR_SW = ""
'        'Set rs_det1 = New ADODB.Recordset
'        'Set dg_det1.DataSource = rsNada
'       ' Set dg_det2.DataSource = rsNada
'    'End If
'  End If


End Sub

Private Sub AbreAlmacen()
    Set rs_datos13 = New ADODB.Recordset
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh

End Sub



Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
' If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
'    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
'        'BtnModDetalle2.Visible = False
'        If (Ado_datos16.Recordset("estado_codigo") = "REG") Then
''            If (Ado_datos.Recordset("estado_codigo") = "APR") Then
''                BtnAprobar2.Visible = False
''            Else
''                BtnAprobar2.Visible = True
''            End If
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = True
''            BtnAnlDetalle2.Visible = True
'            BtnModDetalle2.Visible = True
'        End If
'        If (Ado_datos16.Recordset("estado_codigo") = "APR") Then
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = False
''            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
'        End If
'        If (Ado_datos16.Recordset("estado_codigo") = "ANL") Then
'            'BtnImprimir2.Visible = False
''            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
'            BtnAprobar2.Visible = False
'        End If
'    Else
'        BtnAprobar2.Visible = False
'        BtnImprimir2.Visible = False
''        BtnAnlDetalle2.Visible = False
'        BtnModDetalle2.Visible = False
'    End If
' Else
'    BtnAprobar2.Visible = False
'    BtnImprimir2.Visible = False
''    BtnAnlDetalle2.Visible = False
'    BtnModDetalle2.Visible = False
' End If
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    'If OptFilGral1.Value = True Then Call OptFilGral1_Click
    'If OptFilGral2.Value = True Then Call OptFilGral2_Click
'    Ado_datos.Recordset.Move marca1 - 1
    swnuevo = 1
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    FraNavega.Enabled = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
    'tipo Beneficiario
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    'rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
    rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Dtc_aux2.Text & "' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_Datos12.Recordset = rs_datos12
    Ado_Datos12.Refresh

    ado_datos14.Recordset.AddNew
  Else
    MsgBox "Los productos del registro Aprobado o Entregado, NO pueden ser cambiados !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnanularDetalle2_Click()
 If Ado_datos1.Recordset!estado_codigo = "REG" Then
 sino = MsgBox("Esta seguro de ELIMNAR este registro?", vbYesNo, "Confirmando")
 If sino = vbYes Then
 db.Execute "delete ro_prestamo_prog where prestamo_codigo = " & Ado_datos1.Recordset!prestamo_codigo & " AND prestamo_prog_codigo = " & Ado_datos1.Recordset!prestamo_prog_codigo & " AND beneficiario_codigo = '" & Ado_datos1.Recordset!beneficiario_codigo & "'"
 Call ABRIR_TABLA_DET
 End If
 Else
 sino = MsgBox("No se puede eliminar un registro ya cobrado (APR) ", vbCritical, "Confirmando")
 End If
End Sub

Private Sub BtnAñadir_Click()
 ado = "NO"
  If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    
     modif = "ADD"
    rs_datos.AddNew
    txt_interes.Text = "0"
    'lblStatus.Caption = "Agregar registro"
    Call ABRIR_TABLAS_AUX
    FrmCabecera.Enabled = True
    FraNavega.Enabled = False
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_Total.Visible = False
    cmb_mes_ini.Text = ""
    txt_mes.Text = ""
    DTPFechaIni.Value = Date
    DTPFechaFin.Value = Date
 

    'dg_datos.Enabled = False
   
End Sub

Private Sub BtnAñadirDetalle2_Click()
Fra_crono.Visible = True
modif_det = "NO"
rs_datos1.AddNew
FraNavega.Enabled = False
FrmCobranza.Enabled = False
FrmABMDet2.Enabled = False
fraOpciones.Enabled = False

End Sub

Private Sub BtnAprobar_Click()

    If Ado_datos.Recordset!estado_codigo = "REG" Then
    sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
        If sino = vbYes Then
           rs_datos!estado_codigo = "APR"
            rs_datos!usr_codigo = glusuario
            rs_datos!Fecha_Registro = Date
            rs_datos.Update
        End If
    Else
    sino = MsgBox("Nose puede aprobar un registro Aprobado (APR) o Anulado (ANL)", vbYesNo, "Confirmando")
    End If



End Sub

Private Sub GENERA_COMPRA()
'    If rs_datos!estado_cotiza = "REG" Then
'      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'        ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
'                'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
'                       correldetalle = rs_aux2!correl_negocia
'                       rs_aux2.Update
'                    End If
'                    'WWWWWWWWWWWWWWW
'                    'correlv = Ado_datos.Recordset!venta_codigo
'                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = parametro
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_COD4
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
'                        rs_aux3!venta_tipo = "E"
'                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_descripcion = "COMPRA POR: " + lbl_titulo.Caption
'                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
'                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
'                        rs_aux3!tipo_moneda = "BOB"
'                        rs_aux3!compra_monto_dol = 0        'VAR_DOL2
'                        rs_aux3!proceso_codigo = "TEC"
'                        rs_aux3!subproceso_codigo = "TEC-06"
'                        rs_aux3!etapa_codigo = "TEC-06-01"
'                        rs_aux3!clasif_codigo = "ADM"
'                        rs_aux3!doc_codigo = "R-114"
'                        rs_aux3!poa_codigo = "3.2.8"
'                        rs_aux3!estado_codigo_eqp = "REG"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'
'                        'DETALLE Carga ao_ventas_detalle
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rstdestino.RecordCount > 0 Then
'                        End If
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            VAR_REG = 1
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                              If rs_aux4!grupo_codigo = "30000" Then
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'
'                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
'                                VAR_REG = VAR_REG + 1
'                              End If
'                               rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                    End If
'                    'WWWWWWWWWW
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_cotiza = "APR"
'        rs_datos!fecha_aprueba = Date
'        rs_datos!usr_codigo_aprueba = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnAprobar2_Click()
 If IsNull(Ado_datos16.Recordset("cobranza_observaciones")) Or (Ado_datos16.Recordset("cobranza_programada_bs") = 0) Or Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
    'If Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    Exit Sub
 Else
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
        MsgBox "No se puede APROBAR el registro (Cronograma), previamente debe APROBAR la Venta (Cabecera) y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    If Ado_datos16.Recordset("estado_codigo") = "REG" Then
       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
       If sino = vbYes Then
            db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & Ado_datos.Recordset!venta_codigo & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo, venta_codigo, beneficiario_codigo, beneficiario_codigo_fac, beneficiario_codigo_resp, cobranza_programada_bs, cobranza_programada_dol, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, Literal, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo, estado_codigo, usr_codigo, fecha_registro) " & _
            "VALUES ('" & glGestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & Ado_datos16.Recordset!venta_codigo & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo_resp & "', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '0', '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_observaciones & "', 'FIN', 'FIN-01', 'FIN-01-02', 'ADM', 'R-105', '0', 'R-101', '0', '0', '3.1.2', 'REG', '" & glusuario & "', '" & Date & "')"

'            Set rs_aux1 = New ADODB.Recordset
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            rs_aux1.Open "select * from ao_ventas_cobranza where venta_codigo= " & Ado_datos16.Recordset!venta_codigo & "  and cobranza_prog_codigo= " & Ado_datos16.Recordset!cobranza_prog_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'            If rs_aux1.RecordCount <= 0 Then
'                rs_aux1.AddNew
'            End If
'                rs_aux1!ges_gestion = Ado_datos16.Recordset!ges_gestion
'                rs_aux1!cobranza_prog_codigo = Ado_datos16.Recordset!cobranza_prog_codigo
'                rs_aux1!venta_codigo = Ado_datos16.Recordset!venta_codigo
'                rs_aux1!beneficiario_codigo = Ado_datos16.Recordset!beneficiario_codigo                 'Codigo Beneficiario/Cliente
'                rs_aux1!beneficiario_codigo_resp = Ado_datos16.Recordset!beneficiario_codigo_resp       'Codigo Cobrador
'
'                rs_aux1!cobranza_programada_bs = Ado_datos16.Recordset!cobranza_programada_bs           'Monto Programado Bs
'                rs_aux1!cobranza_programada_dol = Ado_datos16.Recordset!cobranza_programada_dol         'Monto Programado en Dolares
'                rs_aux1!cobranza_deuda_bs = Ado_datos16.Recordset!cobranza_programada_bs                'Monto Cobrado
'                rs_aux1!cobranza_deuda_dol = Ado_datos16.Recordset!cobranza_programada_dol              'Monto en Dolares
'                rs_aux1!cobranza_descuento_bs = 0     'Ado_datos16.Recordset!cobranza_descuento_bs      'Descuento Bs
'                rs_aux1!cobranza_descuento_dol = 0    'Ado_datos16.Recordset!cobranza_descuento_dol     'Descuento Dol
'                rs_aux1!cobranza_total_bs = Ado_datos16.Recordset!cobranza_programada_bs                'Monto Total Bs
'                rs_aux1!cobranza_total_dol = Ado_datos16.Recordset!cobranza_programada_dol              'Monto Total Dol
'                rs_aux1!Literal = Ado_datos16.Recordset!Literal
'                rs_aux1!cobranza_fecha_prog = Ado_datos16.Recordset!cobranza_fecha_prog                 'Fecha de Programada
'                rs_aux1!cobranza_fecha_cobro = Ado_datos16.Recordset!cobranza_fecha_prog                'Fecha de Cobranza
'
'                rs_aux1!cobranza_observaciones = Ado_datos16.Recordset!cobranza_observaciones
'                rs_aux1!proceso_codigo = "COM"
'                rs_aux1!subproceso_codigo = "COM-02"
'                rs_aux1!etapa_codigo = "COM-02-04"
'                rs_aux1!clasif_codigo = "ADM"
'                rs_aux1!doc_codigo = "R-103"
'                rs_aux1!doc_numero = rs_aux1.RecordCount
'                rs_aux1!doc_codigo_fac = ""
'                rs_aux1!cobranza_nro_factura = "0"
'                rs_aux1!cobranza_nro_autorizacion = "0"
'                rs_aux1!poa_codigo = "3.1.2"
'                rs_aux1!estado_codigo = "REG"
'                rs_aux1!usr_codigo = GlUsuario
'                rs_aux1!fecha_registro = Format(Date, "dd/mm/yyyy")
'                rs_aux1!hora_registro = Format(Time, "hh:mm:ss")
'                rs_aux1.Update
            ' APRUEBA ao_ventas_cobranza_prog
            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            'db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            Ado_datos16.Refresh
       End If
    End If
 End If
End Sub

Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      buscados = 1
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
  ado = "SI"
  cmb_mes_ini.Text = ""
  txt_mes.Text = ""

  rs_datos.CancelBatch
  If Ado_datos.Recordset.RecordCount > 0 Then

  If OptFilGral1.Value = True Then
   Call OptFilGral1_Click
  End If
   If OptFilGral2.Value = True Then
   Call OptFilGral2_Click
  End If
  End If
  
  FraNavega.Enabled = True
  FrmCabecera.Enabled = False
  Fra_datos.Enabled = True
  'FrmDetalle.Visible = True
  FrmCobranza.Visible = True
  Fra_Total.Visible = True
  dg_datos.Visible = True
'  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
'  TxtCobrado.Visible = False
'  Label7.Visible = False
'  Cmd_Cliente.Visible = False
'  SSTab1.Tab = 0
 ' SSTab1.TabEnabled(0) = True
  'SSTab1.TabEnabled(1) = True
  'SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnEliminar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
          'Dim rstdestino As New ADODB.Recordset
          'Set rstdestino = New ADODB.Recordset
          'If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  ", db, adOpenDynamic, adLockOptimistic
          'If Not rstdestino.BOF Then rstdestino.MoveFirst
          'If Not rstdestino.BOF And Not rstdestino.EOF Then
          '    rstdestino("estado_codigo") = "E"
          '    rstdestino.Update
          'End If
          'If rstdestino.State = 1 Then rstdestino.Close
          marca1 = Ado_datos.Recordset.Bookmark
          'Ado_datos.Recordset.Requery
          'Ado_datos.Refresh
          Call OptFilGral1_Click
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
    End If
  Else
    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub valida_campos()
VAR_VAL = "OK"
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir Responsable CGI, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If dtc_codigo11 = "" Then
'    MsgBox "Debe Elejir el Tipo de Venta!! (Credito, pago ne Efectivo, etc.), Vuelva a Intentar ...", vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If dtc_codigo2 = "" Then
    MsgBox "Debe la Persona a La que se realizara el prestamo. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If cmd_unimed2 = "" Then
    MsgBox "Debe Elegir la forma de pago. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    
    
    If txt_monto_bs.Text = "" Then
    MsgBox "Debe ingresar el monto del prestamo. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
   End If
    
    
     If cmb_mes_ini.Text = "" Then
    MsgBox "Debe igresar el mes de inicio de la cobranza. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
    End If
    
    
     
     If dtc_desc4.Text = "" Then
    MsgBox "Debe elegir al responsable del prestamo. Consulte con el Administrador del Sistema ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
    
  End If
  If parametro = "DNMAN" Then
    If CDate(Format(DTPFechaFin.Value, "dd/mm/yyyy")) <= CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy")) Then
      MsgBox "La Fecha de Inicio debe ser MENOR a la Fecha de Fin del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
    End If
  Else
    If CDate(Format(DTPFechaFin.Value, "dd/mm/yyyy")) < CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy")) Then
      MsgBox "La Fecha de Inicio debe ser MENOR o IGUAL a la Fecha de Fin del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
      VAR_VAL = "ERR"
      Exit Sub
    End If
  End If
  Select Case RTrim(cmb_mes_ini)
        Case "ENERO"
            VAR_MES2 = 1
        Case "FEBRERO"
            VAR_MES2 = 2
        Case "MARZO"
            VAR_MES2 = 3
        Case "ABRIL"
            VAR_MES2 = 4
        Case "MAYO"
            VAR_MES2 = 5
        Case "JUNIO"
            VAR_MES2 = 6
        Case "JULIO"
            VAR_MES2 = 7
        Case "AGOSTO"
            VAR_MES2 = 8
        Case "SEPTIEMBRE"
            VAR_MES2 = 9
        Case "OCTUBRE"
            VAR_MES2 = 10
        Case "NOVIEMBRE"
            VAR_MES2 = 11
        Case "DICIEMBRE"
            VAR_MES2 = 12
  End Select
  If Month(CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy"))) <> 12 And VAR_MES2 <> 1 Then
    If Val(VAR_MES2) < Month(CDate(Format(DTPFechaIni.Value, "dd/mm/yyyy"))) Then
        MsgBox "El MES de Inicio de Cobranza NO puede ser MENOR al de la Fecha de Inicio del Contrato!! , Vuelva a Intentar ...", vbExclamation, "Atención"
        VAR_VAL = "ERR"
        Exit Sub
    End If
  End If
'  If dtc_codigo11.Text = "C" And dtc_codigo2 = "VD" Then
'        MsgBox "NO se puede realizar la Venta a Credito, Debe cambiar de Cliente ..."
'  Else

End Sub

Private Sub BtnGrabar_Click()


Call valida_campos
If VAR_VAL = "OK" Then
rs_datos!ges_gestion = Year(DTPFechaIni.Value)
rs_datos!unidad_codigo_pla = ""
rs_datos!depto_codigo = ""
rs_datos!unidad_codigo = ""
rs_datos!solicitud_codigo = dtc_sol_cod.Text
rs_datos!beneficiario_codigo = dtc_codigo2.Text
rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
rs_datos!prestamos_descripcion = TxtConcepto.Text
rs_datos!prestamo_bs = txt_monto_bs.Text
rs_datos!prestamo_dol = txt_monto_bs.Text / GlTipoCambioOficial
If txt_interes.Text = "" Then
txt_interes.Text = "0"
End If
rs_datos!intereses_bs = ((txt_monto_bs.Text * txt_interes.Text) / 100)
rs_datos!intereses_dol = rs_datos!intereses_bs / GlTipoCambioOficial
rs_datos!prestamo_total_bs = rs_datos!prestamo_bs + rs_datos!intereses_bs 'mas interes
rs_datos!prestamo_total_dol = rs_datos!prestamo_total_bs / GlTipoCambioOficial 'mas interes

rs_datos!prestamos_descripcion = "PRESTAMO PERSONAL DE Bs. " & rs_datos!prestamo_total_bs & " A NOMBRE DE: " & dtc_desc2.Text

rs_datos!total_cobrado = "0"
rs_datos!tipo_cambio = GlTipoCambioOficial
rs_datos!unidad_codigo_ant = "0"
rs_datos!correl_prog = "0"

rs_datos!fecha_inicio = DTPFechaIni.Value




rs_datos!unimed_codigo = cmd_unimed2.Text
rs_datos!cantidad_periodos = txtCantCobr
rs_datos!mes_inicio_crono = cmb_mes_ini.Text
rs_datos!mes_grupo = txt_mes.Text
rs_datos!estado_codigo = "REG"
rs_datos!Fecha_Registro = Date
rs_datos!usr_codigo = glusuario


 

'rs_datos.Update
If modif = "ADD" Then
rs_datos.Update
Call OptFilGral1_Click
Ado_datos.Recordset.MoveLast
rs_datos!unidad_codigo_ant = "PR-" & Str(rs_datos!prestamo_codigo) & "/" & rs_datos!ges_gestion '------------------>
Else
rs_datos!unidad_codigo_ant = "PR-" & Str(rs_datos!prestamo_codigo) & "/" & rs_datos!ges_gestion '------------------>

End If



db.Execute "DELETE ro_prestamo_prog where prestamo_codigo= " & rs_datos!prestamo_codigo & ""
If rs_datos1.State = 1 Then rs_datos1.Close
rs_datos1.Open "select * From ro_prestamo_prog order by prestamo_codigo", db, adOpenKeyset, adLockOptimistic


mes_grupo = rs_datos!mes_grupo
gestion = rs_datos!ges_gestion
CUOTA = 0
monto_cuota = rs_datos!prestamo_total_bs / Val(txtCantCobr.Text)

If mes_grupo < Month(rs_datos!fecha_inicio) Then
gestion = gestion + 1
End If


While (Ado_datos.Recordset!cantidad_periodos <> CUOTA)

    
rs_datos1.AddNew

CUOTA = CUOTA + 1
rs_datos1!prestamo_codigo = Ado_datos.Recordset!prestamo_codigo
rs_datos1!prestamo_prog_codigo = CUOTA
rs_datos1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
rs_datos1!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
rs_datos1!cobranza_programada_bs = monto_cuota
rs_datos1!cobranza_programada_dol = monto_cuota / GlTipoCambioOficial
rs_datos1!cobranza_descuento_bs = "0"
rs_datos1!cobranza_descuento_dol = "0"
rs_datos1!cobranza_total_bs = rs_datos!prestamo_total_bs
rs_datos1!cobranza_total_dol = rs_datos!prestamo_total_bs / GlTipoCambioOficial

dia = Day(rs_datos!fecha_inicio)
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion

If Weekday(rs_datos1!cobranza_fecha_prog, vbMonday) = 6 Then  'Pregunto si es Domingo
dia = dia + 1
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
ElseIf Weekday(rs_datos1!cobranza_fecha_prog, vbMonday) = 7 Then 'Pregunto si es Sabado
dia = dia + 1
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
Else
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
End If
rs_datos!fecha_fin = rs_datos1!cobranza_fecha_prog



'rs_datos1!cobranza_fecha_cobro = ""
rs_datos1!mes_planilla = mes_grupo
rs_datos1!cobranza_observaciones = "PAGO NUMERO " & CUOTA & " - PERIODO: " & gestion & " " & UCase(MonthName(mes_grupo))
rs_datos1!trans_codigo = "P"
rs_datos1!estado_codigo = "REG"
rs_datos1!usr_codigo = glusuario
rs_datos1!Fecha_Registro = Date

Select Case rs_datos!unimed_codigo
Case "MENSUAL"
  mes_grupo = mes_grupo + 1
Case "BIMESTRAL"
  mes_grupo = mes_grupo + 2
Case "SEMESTRAL"
  mes_grupo = mes_grupo + 6
Case "ANUAL"
   mes_grupo = mes_grupo + 12
End Select

If mes_grupo > 12 Then
mes_grupo = mes_grupo - 12
gestion = gestion + 1
End If
rs_datos1.Update
Wend
rs_datos.Update


    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    FrmCabecera.Enabled = False
    Fra_datos.Enabled = True
    dg_datos.Visible = True
'    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    Fra_Total.Visible = True
'    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
    ado = "SI"
    If modif = "ADD" Then
    Call OptFilGral1_Click
    OptFilGral1.Value = True
    Ado_datos.Recordset.MoveLast
    Else
    rs_datos1.Update
    rs_datos.Update
    Call ABRIR_TABLA_DET
    End If
    
    
End If
End Sub

Private Sub CRONO2()
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "select * from ao_ventas_cabecera where venta_codigo= " & var_cod5 & "  ", db, adOpenKeyset, adLockBatchOptimistic
    'Set AdoAux.Recordset = rsAuxDetalle
    If rs_aux5.RecordCount > 0 Then
      CONT2 = 1
      FInicio = rs_aux5!venta_fecha_inicio
'      CANTOT = rs_aux5!venta_cantidad_total
'      gestion0 = Ado_datos.Recordset("ges_gestion")
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
       VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
       VAR_COBR2 = Ado_datos.Recordset!venta_cantidad_cobr
      MControl = Ado_datos.Recordset!mes_inicio_crono
      Select Case RTrim(MControl)
        Case "ENERO"
            VAR_MES2 = 1
        Case "FEBRERO"
            VAR_MES2 = 2
        Case "MARZO"
            VAR_MES2 = 3
        Case "ABRIL"
            VAR_MES2 = 4
        Case "MAYO"
            VAR_MES2 = 5
        Case "JUNIO"
            VAR_MES2 = 6
        Case "JULIO"
            VAR_MES2 = 7
        Case "AGOSTO"
            VAR_MES2 = 8
        Case "SEPTIEMBRE"
            VAR_MES2 = 9
        Case "OCTUBRE"
            VAR_MES2 = 10
        Case "NOVIEMBRE"
            VAR_MES2 = 11
        Case "DICIEMBRE"
            VAR_MES2 = 12
      End Select
      FControl = FInicio
      CONT3 = 0
      CONT4 = 0
      Select Case VAR_MED2
        Case "MES"
            CONT_MED = 1
        Case "BMES"
            CONT_MED = 2
        Case "TMES"
            CONT_MED = 3
        Case "CMES"
            CONT_MED = 4
        Case "5MES"
            CONT_MED = 5
        Case "SMES"
            CONT_MED = 6
        Case "7MES"
            CONT_MED = 7
        Case "8MES"
            CONT_MED = 8
        Case "9MES"
            CONT_MED = 9
        Case "10MES"
            CONT_MED = 10
        Case "11MES"
            CONT_MED = 11
        Case "ANUAL"
            CONT_MED = 12
      End Select
      
      While (CONT2 <= VAR_COBR2)
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & var_cod5 & "'  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 And corrprog >= VAR_COBR2 Then
            MsgBox "El Cronograma ya fue generado... ", , "Atención"
            CONT2 = CONT2 + 1
        Else
           'wwwwwwwwwwwwwwwwwwwwww
'          Set rs_aux1 = New ADODB.Recordset
'          If rs_aux1.State = 1 Then rs_aux1.Close
'          rs_aux1.Open "select * from ao_ventas_cabecera where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'          If rs_aux1.RecordCount > 0 Then
            correldet2 = rs_aux5!correl_cobro_prog + 1
            rs_aux5!correl_cobro_prog = rs_aux5!correl_cobro_prog + 1
            corrprog = correldet2
            rs_aux5.Update
            
            rs_aux2.AddNew
            rs_aux2!ges_gestion = glGestion
            rs_aux2!venta_codigo = var_cod5 'Ado_datos.Recordset("venta_codigo")
            rs_aux2!cobranza_prog_codigo = correldet2
            rs_aux2!beneficiario_codigo = VAR_BENEF                   'Codigo Beneficiario/Cliente
            'OJO MODIFICAR COBRADOR - JQA 03-ENE-2015
            rs_aux2!beneficiario_codigo_resp = "4333735"  'dtc_codigo4A.Text                                                     'Codigo Cobrador
            'rs_aux2!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
            Set rs_aux6 = New ADODB.Recordset
            If rs_aux6.State = 1 Then rs_aux6.Close
            rs_aux6.Open "select sum(venta_precio_unitario_bs) as acumBs from ao_ventas_detalle where venta_codigo = '" & var_cod5 & "' AND (par_codigo = '99990' or par_codigo = '43340') ", db, adOpenKeyset, adLockReadOnly
            If rs_aux6.RecordCount > 0 Then
                rs_aux2!cobranza_programada_bs = rs_aux6!acumBs                     'Monto Programado Bs
            Else
                rs_aux2!cobranza_programada_bs = 0
            End If
            rs_aux2!cobranza_programada_dol = rs_aux6!acumBs / GlTipoCambioMercado  'Monto Programado en Dolares
            rs_aux2!cobranza_descuento_bs = 0                                       'Descuento Bs
            rs_aux2!cobranza_descuento_dol = 0                                      'Descuento Dol
            rs_aux2!cobranza_total_bs = rs_aux6!acumBs                              'Monto Total Bs
            rs_aux2!cobranza_total_dol = rs_aux6!acumBs / GlTipoCambioMercado       'Monto Total Dol
            'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW=
            'rs_aux2!cobranza_fecha_prog = rs_aux5!venta_fecha_inicio + (30 * CONT2)
            'Dim fdia, fmes, fanio As Integer
            'fmes = Month(FInicio) + CONT2
            
            fdia = Day(FControl)
            fanio = Year(FControl)
            'CONT3 = CONT2 * CONT_MED
            CONT3 = 1
            While (CONT3 <= CONT_MED)
                fmes = Month(FControl)
                Select Case fmes
                    Case 2
                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
                            Dias_Mes = 29
                        Else
                            Dias_Mes = 28
                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        Dias_Mes = 31
                    Case 4, 6, 9, 11
                        Dias_Mes = 30
                End Select
                If Val(VAR_MES2) = Month(FControl) Then
                    rs_aux2!cobranza_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
                    rs_aux2!cobranza_fecha_cobro = FControl + 20
                    VAR_MES2 = VAR_MES2 + CONT_MED
                    If Val(VAR_MES2) > 12 Then
                        VAR_MES2 = Val(VAR_MES2) - 12
                    End If
                End If
                FControl = FControl + Dias_Mes
                CONT3 = CONT3 + 1
                CONT4 = CONT4 + Dias_Mes
            Wend
            'FControl = Str(fdia) + "/" + Str(fmes) + "/" + Str(fanio)
            'rs_aux2!cobranza_fecha_prog = FInicio + (30 * CONT2)
            'rs_aux2!cobranza_fecha_prog = FControl
            If rs_aux2!cobranza_fecha_prog = Null Then
                rs_aux2!cobranza_fecha_prog = Date
            End If
            'VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), Date, rs_aux2!cobranza_fecha_prog)))
            
            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
            'rs_aux2!cobranza_fecha_cobro = FControl + 10 ' rs_aux2!cobranza_fecha_prog + 10
            'If VAR_MED2 = "MES" Then
            '    FControl = FControl + Dias_Mes
            'End If
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(Date)) + " - " + lbl_titulo
            'rs_aux2!cobranza_observaciones = "CUOTA Nro. " + Str(corrprog) + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + lbl_titulo
            Select Case parametro
              Case "DNMAN"
                  rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Trámite: " + VAR_CITE + "-C-" + Str(corrprog)
              Case "DNREP"
                  rs_aux2!cobranza_observaciones = "REPARACION DE EQUIPOS Y/O PROVISION E INSTALACION DE REPUESTOS - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
              Case "DVTA"
                  rs_aux2!cobranza_observaciones = "PROVISION E INSTALACION DE EQUIPOS - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
              Case Else
                  rs_aux2!cobranza_observaciones = lbl_titulo + " - Mes: " + UCase(VAR_FEC2) + "-" + Str(Year(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog))) + " - " + "Cite: " + VAR_CITE + "-C-" + Str(corrprog)
            End Select
            CONT2 = CONT2 + 1
            rs_aux2!cobranza_requisito_plazo = "S"
            rs_aux2!cobranza_concepto_plazo = "CONFORMIDAD DEL SERVICIO"
            If rs_aux2!cobranza_programada_bs <> 0 Then
                rs_aux2!Literal = Literal(CStr(rs_aux2!cobranza_programada_bs)) + " BOLIVIANOS"
            End If
            rs_aux2!proceso_codigo = "TEC"
            rs_aux2!subproceso_codigo = "TEC-02"
            rs_aux2!etapa_codigo = "TEC-02-02"
            rs_aux2!clasif_codigo = "TEC"
            rs_aux2!doc_codigo = "R-105"    ' R-307 Certificado de Mantenimiento ' Colocar en la conformidad
            rs_aux2!doc_numero = "0"        'var_cod5
            rs_aux2!poa_codigo = "3.2.3"
            rs_aux2!estado_codigo = "REG"
            rs_aux2!usr_codigo = glusuario
            rs_aux2!Fecha_Registro = Format(Date, "dd/mm/yyyy")
            rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
            rs_aux2.Update
        End If
      Wend
      MsgBox "El Cronograma fue Generado Exitosamente... ", , "Atención"
      If corrprog > 0 Then
        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & corrprog & "' "
        db.Execute "update ao_ventas_cabecera set venta_plazo_dias_calendario = " & CONT4 & " "
      End If
    Else
       MsgBox "Error Verifique la Venta de Productos..."
    End If
End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        Select Case Me.Ado_datos.Recordset!unidad_codigo
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
          Case "DVTA"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub


Private Sub BtnImprimir1_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
      If ado_datos14.Recordset.RecordCount > 0 Then
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_nota_de_venta.rpt"
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
     Else
        MsgBox "No se puede Imprimir. Debe registrar datos... " & FrmDetalle.Caption, , "Atención"
     End If
   Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
   End If

End Sub

Private Sub BtnModificar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
    modif = "MOD"
        FrmCabecera.Enabled = True
        'FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        'DTPfechasol.SetFocus
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        Fra_datos.Enabled = True
        Fra_Total.Visible = False
    '    If Ado_datos.Recordset!venta_tipo = "E" Then
    '        TxtCobrado.Visible = True
    '        Label7.Visible = True
    '    Else
    '        TxtCobrado.Visible = False
    '        Label7.Visible = False
    '    End If
    '    Cmd_Cliente.Visible = True
'        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False

        swgrabar = 0
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
 '       SSTab1.TabEnabled(1) = False
  '      SSTab1.TabEnabled(2) = False
    Else
      MsgBox "NO se puede MODIFICAR, porque el registro ya fue Aprobado, Anulado o Cerrado.", , "Atencion"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
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


Private Sub BtnVer2_Click()
    NumComp = Ado_datos.Recordset!venta_codigo
    frm_ao_ventas_alcance.Show vbModal
End Sub

Private Sub BtnVer3_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_cancelado = "N" And Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Esta seguro marcar como PROVISIONAL, posteriormente debe modificarlo... ", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_cancelado = 'P' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  "
          marca1 = Ado_datos.Recordset.Bookmark
          If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
          Else
            Call OptFilGral2_Click
          End If
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede procesar el TRAMITE ya fue CERRADO...", , "Atencion"
    End If
  Else
    MsgBox "NO se puede procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub Chk_plazo_Click()
    If Chk_plazo.Value = 1 Then
        lbl_plazo.Visible = True
        txt_plazo.Visible = True
    Else
        lbl_plazo.Visible = False
        txt_plazo.Visible = False
    End If
End Sub

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub

Private Sub CmdCancelaCobro_Click()
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroVenta.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  'Call cerea
  swnuevo = 0
  'cmdElige.Enabled = False
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    'FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
  'ado_datos14.Refresh
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnAnlDetalle2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
      db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & Ado_datos16.Recordset("cobranza_codigo") & " "
      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_deuda_bs = '0', ao_ventas_cobranza_prog.cobranza_deuda_dol = '0'  Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & ado_datos16.Recordset("cobranza_codigo") & " "

     'ado_ventas_COBRANZAS.Recordset.Delete
     'ado_ventas_COBRANZAS.Recordset.Update
     'ado_ventas_COBRANZAS.Requery
     'ado_ventas_COBRANZAS.Refresh
     ''cerea
     'ado_ventas_COBRANZAS.Refresh
   End If
  Else
    MsgBox "Los productos del registro sin Aprobar, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  'If Ado_datos.Recordset!venta_tipo <> "E" And Ado_datos16.Recordset!estado_codigo = "REG" Then
  If Ado_datos1.Recordset!estado_codigo = "REG" Then
  modif_det = "SI"
  Fra_crono.Visible = True
  FraNavega.Enabled = False
FrmCobranza.Enabled = False
FrmABMDet2.Enabled = False
fraOpciones.Enabled = False
  
'    marca1 = Ado_datos16.Recordset.Bookmark
'    FraNavega.Enabled = False
'    fraOpciones.Enabled = False
'    FrmDetalle.Visible = False
'    FrmCobranza.Visible = False
'    VAR_COBR1 = Ado_datos16.Recordset!cobranza_prog_codigo
'    'swgrabar = 0
'    swnuevo = 2
'    TxtCobrador.Visible = False
'    'TxtMonto.SetFocus
'    'TxtNroVenta.Enabled = False
'    'marca1 = ado_datos14.Recordset.BookMark
'    'txt_descripcion_venta.Enabled = True
'    'TxtNroVenta.Text = txt_venta.Text
'    'lbltipoVenta.Caption = dtc_desc11.Text
'    'lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'    FrmCobros.Visible = True
'    FrmCobros.Enabled = True
'    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
'    'If Ado_datos.Recordset!estado_codigo = "APR" Then
'        'sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
'        'If sino = vbYes Then
'        '    DTPFechaProg.Visible = False
'        '    DTPFechaCobro.Visible = True
'        '    Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'        '    lbl_fechas.Caption = "Fecha Efectiva de Cobranza"
'        '    Txt_parche.Visible = False      '&H80000013&
'        '    'dtc_desc2A.BackColor = &H80000013
'        'Else
'        '    DTPFechaProg.Visible = True
'        '    DTPFechaCobro.Visible = False
'        '    Lbl_nombre_fac.Caption = "Cliente :"
'        '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
'        '    Txt_parche.Visible = True       '&H80000005&
'        '    'dtc_desc2A.BackColor = &H80000005
'        'End If
'    'Else
'        DTPFechaProg.Visible = True
'        DTPFechaCobro.Visible = False
'        DTPFechaConf.Visible = True
'        Lbl_nombre_fac.Caption = "Cliente :"
'        lbl_fechas.Caption = "Fecha Programada de Cobranza"
''        Txt_parche.Visible = True       '&H80000005&
'        'dtc_desc2A.BackColor = &H80000005
'    'End If
'    VAR_MBS2 = Ado_datos16.Recordset!cobranza_programada_bs
'    TxtMonto.SetFocus
''    Call ABRIR_TABLA_DET
''    Ado_datos16.Recordset.Move marca1 - 1


  Else
    MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos16.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 2
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = False
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        'Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
'        Dim thisDate As Date
'        Dim thisMonth As Integer
'        thisDate = #2/12/1969#
'        thisMonth = Month(thisDate)
'        ' thisMonth now contains 2.
'
'
'        Dim thisMonth As Integer
'        Dim name As String
'        thisMonth = 4
'        ' Set Abbreviate to True to return an abbreviated name.
'        name = MonthName(thisMonth, True)
'        ' name now contains "Apr".
    Else
        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
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

Private Sub cmdElige_Click()
  With ALFrmMateriales
        .ALPrincipal
        If .QResp Then
            TxtCodigo.Text = .QCodigo
            txtDesc.Text = .QItem
        End If
    End With
    Txtcant_alm = 0
    Cant_Alm = 0
    DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
    Txtcant_alm = Cant_Alm
    If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

Private Sub Contabiliza_venta()
    Call graba_proyecto
    Call graba_ingreso
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
              'cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
              'Subcta_deb11 = rstdestino!Subcta_cred1
              'Subcta_deb21 = rstdestino!Subcta_cred2

              'cta_credito1 = rstdestino2!cta_deb
              'Subcta_cred11 = rstdestino2!Subcta_deb1
              'Subcta_cred21 = rstdestino2!Subcta_deb2
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "REC"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close

        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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

    'If DtCCta_codigo.Text <> "01" Then
    '  If rstdestino.State = 1 Then rstdestino.Close
    '  rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
    '  If Not rstFc_cuenta_bancaria.EOF Then
    '    fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
    '  Else
    '  End If
    'Else
    '    fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    'If VAR_CODTIPO = "DEI" Or VAR_CODTIPO = "DES" Then
    '  fte_codigo1 = Me.DtCFte_codigo.Text
    'End If

'    fte_codigo1 = VAR_FTE
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    v_Tipo_Comp(1, 1) = VAR_CODTIPO

'    If VAR_CODTIPO = "DYR" Then
'      'j = 2
'      'v_Tipo_Comp(1, 1) = "CAD"
'      'v_Tipo_Comp(1, 2) = "CAR"
'      j = 2
'      v_Tipo_Comp(1, 1) = "DYR"
'    Else
'      j = 1
'      v_Tipo_Comp(1, 1) = IIf(VAR_CODTIPO = "DEI", "DEI", IIf(VAR_CODTIPO = "REC", "REC", IIf(VAR_CODTIPO = "DES", "DES", IIf(VAR_CODTIPO = "ANI", "ANI", ""))))
'    End If
'
'    If VAR_CODTIPO = "DVI" Then
'      j = 1
'      v_Tipo_Comp(1, 1) = "DVI"
'    End If

'    For i = 1 To j
'      If rstdestino.State = 1 Then rstdestino.Close
'      If v_Tipo_Comp(1, i) = "DEI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'        Exit Sub
'      End If

    ' INI CORRECCION 18-JUN-2014
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' 02/07/2014 VERIFICAR
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rs_aux2.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'          Exit Sub
'        End If
'      End If
'
'      If rs_aux2.RecordCount < 1 Then
'        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'        Exit Sub
'      End If
'    Next

    'If rstdestino.State = 1 Then rstdestino.Close

    fte_codigo1 = VAR_FTE
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
      If rs_aux2.RecordCount > 0 Then
        yacontabilizo = 1
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
        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
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
      rs_aux2("ges_gestion") = glGestion    'Year(Date)
      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
      If yacontabilizo = 0 Then
        rs_aux2("Fecha_transacion") = Date
      End If
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4       'Ado_datos.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = Ado_datos.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE

      rs_aux2("proceso_codigo") = "FIN"
      rs_aux2("subproceso_codigo") = "FIN-02"
      Select Case VAR_CODTIPO
        Case "DEI"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "REC"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DYR"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "DES"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "ANI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DVI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
      End Select

      rs_aux2("clasif_codigo") = "ADM"
      rs_aux2("doc_codigo") = "R-128"
      rs_aux2("doc_numero") = Var_Comp
      rs_aux2("pro_codigo_det") = VAR_PROY2

      rs_aux2("estado_codigo") = "APR"

      If yacontabilizo = 0 Then
        rs_aux2("usr_codigo") = glusuario
        rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
        rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      rs_aux2.Update
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
'      If v_Tipo_Comp(1, i) = "DEI" Then     'Devengado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If

'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
'          cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'          Subcta_deb11 = rstdestino!Subcta_cred1
'          Subcta_deb21 = rstdestino!Subcta_cred2
'
'          cta_credito1 = rstdestino2!cta_deb
'          Subcta_cred11 = rstdestino2!Subcta_deb1
'          Subcta_cred21 = rstdestino2!Subcta_deb2
'        Else
'          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''          Exit Sub
'        End If
'      End If
'
'      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'        cta_credito1 = rstdestino("cta_cred")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
'      Else
'        'MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'        'Exit Sub
'
'      End If
      '2115
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")

        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2

        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
      End If

      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = rs_aux1("NombreCta")
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = rs_aux1("NombreCta")
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
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
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = Trim(d_cta_nombre_1) ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        ' para Aux1
'        Select Case d_aux1_1
'                Case "01"
'                    VAR_COD1 = VAR_BENEF
'                Case "02"
'                    VAR_COD1 = VAR_CTA
'                Case "03"
'                    VAR_COD1 = VAR_PROY2
'                Case "04"
'                    VAR_COD1 = Ado_datos.Recordset("unidad_codigo")
'                Case "05"
'                    VAR_COD1 = ""
'                Case "06"
'                    VAR_COD1 = ""
'                Case "07"
'                    VAR_COD1 = ""
'                Case "08"
'                    VAR_COD1 = ""
'                Case "09"
'                    VAR_COD1 = VAR_ORG
'                Case "10"
'                    VAR_COD1 = ""
'                Case "11"
'                    VAR_COD1 = ""
'                Case "12"
'                    VAR_COD1 = ""
'        End Select
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
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
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

        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
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

        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = "2"
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
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        ' CORREGIR MONTOS JQA 2014-JUL-08
        If j > 1 Then
            If i = 1 Then
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = Trim(h_cta_nombre_1) ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
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

        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = "2"
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

        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
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

'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If j > 1 Then
            If i = 1 Then
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
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
        If i = 1 Then
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
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
        If i = 1 Then
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado
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

      If yacontabilizo = 0 Then
        rstdestino2("Usr_codigo") = glusuario
        rstdestino2("Fecha_registro") = Date
        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If

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
  'marca1 = Ado_datos.Recordset.Bookmark
  'rs_datos.Update
  'rs_datos.Requery
  Call OptFilGral1_Click
  'Set Ado_datos.Recordset = rs_datos
  'If rs_datos.RecordCount > 0 Then
    Ado_datos.Recordset.Move marca1 - 1
  'End If
  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

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
    Select Case Ado_datos.Recordset!unidad_codigo
        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
            VAR_PROY = 12
        Case "UCOM"
            VAR_PROY = 17
        Case "DVTA"
            VAR_PROY = 18

    End Select

    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
    Else
        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & glGestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
    End If
    '
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
   VAR_COD4 = Ado_datos.Recordset!unidad_codigo
   Select Case VAR_COD4
        Case "DVTA"              'INI COMERCIAL
            VAR_ORG = "111"
            VAR_PARTIDA = "11200"
        Case "COMEX"            'INI COMEX
            VAR_ORG = "111"
            VAR_PARTIDA = "11310"
        Case "DNINS"            'INI INSTALACIONES
            VAR_ORG = "111"
            VAR_PARTIDA = "11310"
        Case "DNAJS"            'INI AJUSTE
            VAR_ORG = "113"
            VAR_PARTIDA = "11310"
        Case "DNMAN"            'INI MANTENIMIENTO
            VAR_ORG = "112"
            VAR_PARTIDA = "11320"
        Case "DNREP"            'INI REPARACIONES
            VAR_ORG = "113"
            VAR_PARTIDA = "11330"
        Case "DNMOD"            'INI MODERNIZACION
            VAR_ORG = "114"
            VAR_PARTIDA = "11340"
        Case "DNEME"            'INI EMERGENCIAS
            VAR_ORG = "113"
            VAR_PARTIDA = "11350"
        Case Else               'INI CREDITO
            VAR_ORG = "311"
   End Select
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False

      'If v_añadir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         Call add_correl
         Set rstdestino = New ADODB.Recordset
         rstdestino.Open "select * from fo_ingresos_cabecera order by org_codigo, ingreso_codigo   ", db, adOpenDynamic, adLockOptimistic
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = glGestion      'Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         VAR_CODANT = correlativo1
         'CAMBIAR org_codigo
         rstdestino("org_codigo") = VAR_ORG
         'CAMBIAR org_codigo
         'CAMBIAR COD ingreso_codigo_anterior
         rstdestino("ingreso_codigo_anterior") = correlativo1
         'CAMBIAR COD ingreso_codigo_anterior
         'CAMBIAR DEI O REC
         rstdestino("Codigo_tipo") = "DEI"
         VAR_CODTIPO = "DEI"
         'CAMBIAR DEI O REC
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-01"
         rstdestino("etapa_codigo") = "FIN-01-01"
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = "R-110"
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos.Recordset("unidad_codigo")
         rstdestino("solicitud_codigo") = Ado_datos.Recordset("solicitud_codigo")
         rstdestino("solicitud_tipo") = "10"

         rstdestino("beneficiario_codigo") = Ado_datos.Recordset("beneficiario_codigo")
         VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
         rstdestino("tipo_moneda") = "BOB"
         VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = "INGRESO POR: " + Ado_datos.Recordset("venta_descripcion")
         VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("venta_descripcion")
         If Ado_datos.Recordset("venta_tipo") = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = "DEI"
         End If
         'CAMBIAR FTE
         Select Case VAR_ORG
             Case "111"              'INI SERVICIOS DE PROVISION E INSTALACION
                 VAR_FTE = "10"
             Case "112"            'INI SERVICIO DE MANTENIMIENTO - MANTENIMIENTO PREVENTIVO
                 VAR_FTE = "10"
             Case "113"            'INI SERVICIO DE REPARACIONES - MANTENIMIENTO CORRECTIVO
                 VAR_FTE = "10"
             Case "114"            'INI SERVICIO DE MODERNIZACION
                 VAR_FTE = "10"
             Case "211"            'INI APORTES DE CAPITAL
                 VAR_FTE = "20"
             Case "311"            'INI BANCO MERCANTIL SANTA CRUZ
                 VAR_FTE = "30"
             Case "312"            'INI BANCO DE CREDITO
                 VAR_FTE = "30"
             Case "411"            'INI AMT - REPOSICION DE PIEZAS Y PARTES
                 VAR_FTE = "40"
             Case Else               'INI OTROS
                 VAR_FTE = "10"
        End Select
         rstdestino("fte_codigo") = VAR_FTE
         'CAMBIAR FTE
         'CAMBIAR RUBROS    'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww ya pues
         'rstdestino("rubro_codigo") = "11200"
         'VAR_PARTIDA = "11200"
         'VAR_PARTIDA = "11320"
         rstdestino("rubro_codigo") = VAR_PARTIDA
         'CAMBIAR RUBROS
         rstdestino("cheque_o_trf") = ""
         rstdestino("Bco_codigo") = "NN"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = "NN"
         VAR_CTA = "NN"
         'CAMBIAR CTA
         rstdestino("numero_documento") = "0"
         rstdestino("unidad_codigo_ant") = Ado_datos.Recordset("unidad_codigo_ant")
         VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
         rstdestino("monto_dolares") = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         rstdestino("monto_bolivianos") = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         rstdestino("monto_recaudado_dolares") = 0
         rstdestino("monto_recaudado_bolivianos") = 0
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = Ado_datos.Recordset("edif_codigo")
         VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
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
'   Else
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
  'FALTAAAAA!! org_codigo JQA 2014-07-10
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     rstcorrel_ing.AddNew
     rstcorrel_ing("org_codigo") = VAR_ORG
     rstcorrel_ing("ges_gestion") = glGestion       'Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
     'rstcorrel_ing("correlativo") = 1
     rstcorrel_ing("correlativo_ingreso") = 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
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
                rstpagos("ges_gestion") = glGestion     'Ado_datos.Recordset("ges_gestion")
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
  If dtc_codigo4A = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
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
  If txtDoc = "" Then
    MsgBox "Debe Registrar el " + lblccertif.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If DTPFechaConf = "" Then
    MsgBox "Debe Registrar la " + lblfechaCertif.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "select sum(cobranza_programada_bs) as totbs2, sum (cobranza_programada_dol) as totdl2 from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
    If IsNull(rs_aux3!totbs2) Then
        If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
            'MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
            MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el Monto total del Contrato . !! Vuelva a Intentar ...", vbExclamation, "Atención"
            If rs_aux3.State = 1 Then rs_aux3.Close
            Exit Sub
        End If
    Else
        If swnuevo = 1 Then
            If (rs_aux3!totbs2) + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                'MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
                MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el Monto total del Contrato . !! Vuelva a Intentar ...", vbExclamation, "Atención"
                If rs_aux3.State = 1 Then rs_aux3.Close
                Exit Sub
            End If
        Else
'            If (rs_aux3!totbs2) - VAR_MBS2 + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
'                MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
'                If rs_aux3.State = 1 Then rs_aux3.Close
'                Exit Sub
'            End If
        End If
    End If
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    If swnuevo = 1 Then
      Set rs_aux1 = New ADODB.Recordset
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from ao_ventas_cabecera where venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
      If rs_aux1.RecordCount > 0 Then
         correldet2 = rs_aux1!correl_cobro_prog + 1
         If rs_aux1!correl_cobro_prog > 1 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "Select * from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & " and cobranza_prog_codigo = " & rs_aux1!correl_cobro_prog & " ", db, adOpenStatic
            If rs_aux2.RecordCount > 0 Then
                If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
                    MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
                    If rs_aux1.State = 1 Then rs_aux1.Close
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    db.CommitTrans
                    Exit Sub
                End If
            End If

         End If
         rs_aux1!correl_cobro_prog = rs_aux1!correl_cobro_prog + 1
         rs_aux1.Update
      End If
      'Ado_datos16.Recordset.AddNew
      Ado_datos16.Recordset!cobranza_prog_codigo = correldet2
      Ado_datos16.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
      Ado_datos16.Recordset!ges_gestion = glGestion      'Ado_datos.Recordset("ges_gestion")
    End If
    If swnuevo = 2 Then
      If Ado_datos16.Recordset!cobranza_prog_codigo > 1 Then
        correldet2 = Ado_datos16.Recordset!cobranza_prog_codigo - 1
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "Select * from ao_ventas_cobranza_prog where venta_codigo=" & Ado_datos.Recordset!venta_codigo & " and cobranza_prog_codigo = " & correldet2 & " ", db, adOpenStatic
        If rs_aux2.RecordCount > 0 Then
          If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then 'DTPFechaProg.Value
          'If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
              MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
              If rs_aux2.State = 1 Then rs_aux2.Close
              db.CommitTrans
              Exit Sub
          End If
        End If
      End If
    End If
      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
      Ado_datos16.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      'Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos16.Recordset!cobranza_programada_bs = CDbl(TxtMonto.Text)                                  'Monto Programado Bs
      Ado_datos16.Recordset!cobranza_programada_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto Programado en Dolares
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'      Ado_datos16.Recordset!cobranza_deuda_bs = 0   'CDbl(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!cobranza_deuda_dol = 0  'CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
      'If TxtDscto.Text = "" Or TxtDscto.Text = "0" Or TxtDscto.Text = "0.00" Then
        Ado_datos16.Recordset!cobranza_descuento_bs = 0                                 'Descuento Bs
        Ado_datos16.Recordset!cobranza_descuento_dol = 0                                    'Descuento Dol
      'Else
      '  Ado_datos16.Recordset!cobranza_descuento_bs = CDbl(TxtDscto.Text)                                 'Descuento Bs
      '  Ado_datos16.Recordset!cobranza_descuento_dol = CDbl(TxtDscto.Text) / GlTipoCambioMercado        'Descuento Dol
      'End If
      Ado_datos16.Recordset!cobranza_total_bs = 0   'Ado_datos16.Recordset!cobranza_deuda_bs - Ado_datos16.Recordset!cobranza_descuento_bs               'Monto Total Bs
      Ado_datos16.Recordset!cobranza_total_dol = 0  'Ado_datos16.Recordset!cobranza_deuda_dol - Ado_datos16.Recordset!cobranza_descuento_dol               'Monto Total Dol
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos16.Recordset!cobranza_programada_bs <> 0 Then
            Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos16.Recordset!cobranza_programada_bs)) + " BOLIVIANOS"
            'Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      End If
      Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaConf.Value                                'Fecha de Cobranza cobranza_fecha_conformidad
      Ado_datos16.Recordset!cobranza_fecha_conformidad = DTPFechaConf.Value                                'Fecha de Cobranza
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))

      If Chk_plazo.Value = 1 Then
        lbl_plazo.Visible = True
        txt_plazo.Visible = True
        Ado_datos16.Recordset!cobranza_requisito_plazo = "S"
        Ado_datos16.Recordset!cobranza_concepto_plazo = "CERTIFICADO DE MANTENIMIENTO R-307 Nro. " + txtDoc
      Else
        lbl_plazo.Visible = False
        txt_plazo.Visible = False
        Ado_datos16.Recordset!cobranza_requisito_plazo = "N"
        Ado_datos16.Recordset!cobranza_concepto_plazo = txt_plazo.Text
      End If
      Ado_datos16.Recordset!nro_fojas = IIf(txt_fojas.Text = "", "1", txt_fojas.Text)
      Ado_datos16.Recordset!cobranza_observaciones = TxtObs.Text
      Ado_datos16.Recordset!proceso_codigo = "TEC"
      Ado_datos16.Recordset!subproceso_codigo = "TEC-02"
      Ado_datos16.Recordset!etapa_codigo = "TEC-02-02"
      Ado_datos16.Recordset!clasif_codigo = "TEC"
      Ado_datos16.Recordset!doc_codigo = "R-307"
      Ado_datos16.Recordset!doc_numero = IIf(txtDoc = "", "0", txtDoc)
      Ado_datos16.Recordset!doc_codigo_crono = "R-360"
      Ado_datos16.Recordset!doc_numero_crono = Ado_datos.Recordset("venta_codigo")
'      Ado_datos16.Recordset!doc_codigo_fac = ""
'      Ado_datos16.Recordset!cobranza_nro_factura = "0"       'Trim(TxtCmpbte)
'      Ado_datos16.Recordset!cobranza_nro_autorizacion = "0"       'Trim(TxtCmpbte)
      Ado_datos16.Recordset!poa_codigo = "3.1.2"
      'If DTPFechaProg.Visible = False Then
      '  Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value         'Fecha de Cobranza
      'Else
        'Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaProg.Value         'Fecha de Cobranza
        Ado_datos16.Recordset!cobranza_fecha_prog = DTPFechaProg.Value           'Fecha Programada de Cobranza
      'End If
      Ado_datos16.Recordset!estado_codigo = "REG"
      Ado_datos16.Recordset!usr_codigo = glusuario
      Ado_datos16.Recordset!Fecha_Registro = Format(Date, "dd/mm/yyyy")
      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos16.Recordset.Update
    db.CommitTrans
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos16.Refresh
    'Ado_datos16.Recordset.MoveLast
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    FrmCobros.Enabled = False
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    swnuevo = 0
    gestion0 = glGestion
    'gestion0 = Ado_datos.Recordset("ges_gestion")
    'correlv = Ado_datos.Recordset("correl_venta")
    nroventa = Ado_datos.Recordset("venta_codigo")

'  Set rstacumdet = New ADODB.Recordset
'  If rstacumdet.State = 1 Then rstacumdet.Close
'  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranza_prog where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
'
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & gestion0 & "' and venta_codigo = " & nroventa, db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount > 0 Then
'    rstdestino!deuda_cobrada = rstacumdet!Cobrobs
'    rstdestino!saldo_p_cobrar = (rstdestino!monto_total_Bs - rstdestino!monto_cobrado - rstdestino!deuda_cobrada)
'    rstdestino.Update
'  End If
'  If rstdestino.State = 1 Then rstdestino.Close
'  If rstacumdet.State = 1 Then rstacumdet.Close

  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If
End Sub

Private Sub CmdGrabaDet_Click()
'If dtc_desc12 = "" Then
'    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
  If dtc_codigo15 = "" Then
     MsgBox "Debe Elejir un Producto para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
'  If dtc_desc13 = "" Then
'    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
    '    VAR_PARTIDA = "OK"
    If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Or Dtc_partida15.Text = "99990" Then
          'fraOpciones.Visible = True
          'FraGrabarCancelar.Visible = False
          'TxtNroVenta.Enabled = True
          FrmEdita.Enabled = False
        '  DtGListaN.Enabled = True
          'cmdElige.Enabled = False
        '  dtc_codigo15.Visible = False
        '  dtc_desc15.Visible = False
          'txt_descripcion_venta.Enabled = False
        If swnuevo = 1 Then
          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
          ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
          ado_datos14.Recordset!ges_gestion = glGestion     'Ado_datos.Recordset("ges_gestion")
          ado_datos14.Recordset!bien_codigo = Trim(dtc_codigo15.Text)                       'Codigo Bien (Equipo, Producto, etc)
        End If
          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
          ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
          ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
          If TxtCantidad.Text = "" Then
            TxtCantidad.Text = "1"
          End If
          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
          If CDbl(TxtDescuento) > 0 Then
            ado_datos14.Recordset!venta_descuento_bs = CDbl(TxtDescuento.Text)      'Dcto por producto CON DESCUENTO
            ado_datos14.Recordset!venta_descuento_dol = Val(TxtDescuento) / GlTipoCambioMercado
          Else
            'ado_datos14.Recordset!descuento_venta = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) * (CDbl(Dtc_aux12)) 'Dcto por producto DE LA TABLA
            TxtDescuento.Text = "0"
            ado_datos14.Recordset!venta_descuento_bs = 0
            ado_datos14.Recordset!venta_descuento_dol = 0
          End If
          ado_datos14.Recordset!venta_precio_total_bs = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) - (CDbl(TxtDescuento)) 'Precio Total Producto
          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
          ado_datos14.Recordset!venta_precio_total_dol = (ado_datos14.Recordset!venta_precio_total_bs) / GlTipoCambioMercado
          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
          If OpMod1.Value = True Then
            ado_datos14.Recordset!modelo_elegido = "S"
            ado_datos14.Recordset!modelo_elegido_h = "N"
            ado_datos14.Recordset!modelo_elegido_x = "N"
          End If
          If OpMod2.Value = True Then
            ado_datos14.Recordset!modelo_elegido_h = "S"
            ado_datos14.Recordset!modelo_elegido = "N"
            ado_datos14.Recordset!modelo_elegido_x = "N"
          End If
          If OpMod2.Value = True Then
            ado_datos14.Recordset!modelo_elegido_x = "S"
            ado_datos14.Recordset!modelo_elegido = "N"
            ado_datos14.Recordset!modelo_elegido_h = "N"
          End If
          ado_datos14.Recordset!estado_codigo = "REG"
          ado_datos14.Recordset!usr_codigo = glusuario
          ado_datos14.Recordset!Fecha_Registro = Format(Date, "dd/mm/yyyy")
          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
          ado_datos14.Recordset.Update
        'db.CommitTrans
        'actualiza MODELO del equipo
        db.Execute "update ac_bienes set modelo_codigo = '" & ado_datos14.Recordset!modelo_codigo & "' Where grupo_codigo = '" & ado_datos14.Recordset!grupo_codigo & "' And subgrupo_codigo = '" & ado_datos14.Recordset!subgrupo_codigo & "'  And bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' "
        'Acumula MONTOS
        Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        FraNavega.Enabled = True
        FrmDetalle.Enabled = True
        'FrmDetalle.Visible = True
        FrmCobranza.Visible = True
        FrmABMDet.Visible = True
        FrmABMDet2.Visible = True
        Call OptFilGral1_Click
        If swnuevo = 1 Then
          'Call abre_ventas_det
          'rs_datos14.Requery
          'ado_datos14.Refresh
          'ado_datos14.Recordset.MoveLast

        End If
        swnuevo = 0
    Else
        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
    End If
  'Else
  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
  'End If
End Sub

Private Sub BtnImprimir2_Click()
  If Ado_datos16.Recordset.RecordCount > 0 Then
    Dim iResult As Variant  ', i%, y%
    'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-105_kardex.rpt"
    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
    CryR01.WindowShowRefreshBtn = True
    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
    CryR01.StoredProcParam(2) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
    'Literal por el Total de la Compra
    var_literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
    CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
    'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos16.Recordset!cobranza_prog_codigo & "' "
    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
    iResult = CryR01.PrintReport
    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
'     ado_datos14.Recordset.Delete
'     ado_datos14.Recordset.Update
'     rs_datos14.Requery
'     ado_datos14.Refresh
'     'cerea
'     ado_datos14.Refresh
      'db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
      db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
   End If
  Else
    MsgBox "Los Bienes del registro Aprobado o Anulado, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    FraNavega.Enabled = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    swgrabar = 0
    swnuevo = 2
    marca1 = Ado_datos.Recordset.Bookmark
    'txt_descripcion_venta.Enabled = True
    TxtNroVenta.Text = Ado_datos.Recordset!venta_codigo  'txt_venta.Text
    TxtNroVenta.Enabled = False
    'lbltipoVenta.Caption = dtc_desc11.Text
'    lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(2) = False
    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False

    If ado_datos14.Recordset!modelo_elegido = "S" Then
        OpMod1.Value = True
        OpMod2.Value = False
        OpMod3.Value = False
    End If
    If ado_datos14.Recordset!modelo_elegido_h = "S" Then
        OpMod1.Value = False
        OpMod2.Value = True
        OpMod3.Value = False
    End If
    If ado_datos14.Recordset!modelo_elegido_x = "S" Then
        OpMod1.Value = False
        OpMod2.Value = False
        OpMod3.Value = True
    End If
    'dtc_codigo13.Text
    If ado_datos14.Recordset!par_codigo = "43340" Then
        dtc_codigo13.Text = "0"
        dtc_desc13.BoundText = dtc_codigo13.BoundText
        dtc_desc13.backColor = &H80000013
        dtc_desc13.ForeColor = &HFFFFFF
    Else
        dtc_desc13.backColor = &HFFFFFF
        dtc_desc13.ForeColor = &H80000008
    End If

    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "select * from Gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
    Set Ado_Datos12.Recordset = rs_datos12
    Ado_Datos12.Refresh
    Dtc_aux12.BoundText = dtc_codigo12.BoundText
    dtc_desc12.BoundText = dtc_codigo12.BoundText
  Else
    MsgBox "Los productos del registro Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmb_mes_ini_Click()
txt_mes.Text = cmb_mes_ini.ListIndex
txt_mes.Text = Val(txt_mes.Text) + 1
End Sub

Private Sub cmb_mes_ini_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub



Private Sub cmd_unimed2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 0 Then
KeyAscii = 0
Else
Exit Sub
End If
End Sub

'Private Sub CmdDetCabeza_Click()
'    fraOpciones.Visible = False
'    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
'    FraNavega.Enabled = False
'    If Not (adoDetalleSolicitud.Recordset.BOF) Then adoDetalleSolicitud.Recordset.MoveFirst
'End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_aux2.BoundText
    dtc_desc2.BoundText = Dtc_aux2.BoundText
    Dtc_deudor2.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_aux4.BoundText
    dtc_desc4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    dtc_haber_men.BoundText = dtc_codigo2.BoundText
        dtc_sol_cod.BoundText = dtc_codigo2.BoundText
'    Dtc_aux2.BoundText = dtc_codigo2.BoundText
'    Dtc_deudor2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    dtc_haber_men.BoundText = dtc_desc2.BoundText
    dtc_sol_cod.BoundText = dtc_desc2.BoundText
'    Dtc_aux2.BoundText = dtc_desc2.BoundText
'    Dtc_deudor2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    
'    dtc_aux4.BoundText = dtc_desc4.BoundText
    VAR_BEN2 = dtc_codigo4.Text
End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.Text = VAR_BEN2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Dtc_deudor2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_deudor2.BoundText
    Dtc_aux2.BoundText = Dtc_deudor2.BoundText
    dtc_desc2.BoundText = Dtc_deudor2.BoundText
End Sub

Private Sub dtc_codigo13_Click(Area As Integer)
    dtc_desc13.BoundText = dtc_codigo13.BoundText
    Dtc_Stock13.BoundText = dtc_codigo13.BoundText
End Sub

Private Sub dtc_desc13_Click(Area As Integer)
    dtc_codigo13.BoundText = dtc_desc13.BoundText
    Dtc_Stock13.BoundText = dtc_desc13.BoundText
End Sub

Private Sub dtc_codigo2A_Click(Area As Integer)
    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc11_LostFocus()
    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "V" Then
        'TxtCobrado.Visible = False
        'Label7.Visible = False
        TxtConcepto.Text = Right(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text), Len(lbl_titulo.Caption + " -Edificio: " + RTrim(dtc_desc3.Text)) - 6)
'        TxtPlazo.Visible = True
    Else
        If dtc_codigo11.Text = "E" Then
            TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)
            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
        Else
        'dtc_codigo2.Text = "VD"
        'dtc_desc2.Text = "VENTA DIRECTA"
        'TxtCobrado.Visible = True
        'Label7.Visible = True
            'TxtConcepto.Text = lbl_titulo.Caption + " - " + RTrim(dtc_desc3.Text)   '"VENTA DIRECTA AL CLIENTE"
            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
        End If
    End If
End Sub

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

Private Sub dtc_codigo15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_codigo15.BoundText
    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
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

Private Sub dtc_precioventabase15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_precioventabase15.BoundText
    dtc_unimed15.BoundText = dtc_precioventabase15.BoundText
    dtc_stocktotal15.BoundText = dtc_precioventabase15.BoundText
    dtc_grupo15.BoundText = dtc_precioventabase15.BoundText
    dtc_subgrupo15.BoundText = dtc_precioventabase15.BoundText
    Dtc_partida15.BoundText = dtc_precioventabase15.BoundText
    dtc_precioventafinal15.BoundText = dtc_precioventabase15.BoundText
    dtc_codigo15.BoundText = dtc_precioventabase15.BoundText
    dtc_preciocompra15.BoundText = dtc_precioventabase15.BoundText
End Sub

Private Sub dtc_subgrupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_subgrupo15.BoundText
    dtc_desc15.BoundText = dtc_subgrupo15.BoundText
    dtc_unimed15.BoundText = dtc_subgrupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_subgrupo15.BoundText
    dtc_grupo15.BoundText = dtc_subgrupo15.BoundText
    Dtc_partida15.BoundText = dtc_subgrupo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_subgrupo15.BoundText
    dtc_precioventabase15.BoundText = dtc_subgrupo15.BoundText
    dtc_preciocompra15.BoundText = dtc_subgrupo15.BoundText
End Sub

Private Sub dtc_partida15_Click(Area As Integer)
    dtc_desc15.BoundText = Dtc_partida15.BoundText
    dtc_unimed15.BoundText = Dtc_partida15.BoundText
    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
    dtc_grupo15.BoundText = Dtc_partida15.BoundText
    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
    dtc_codigo15.BoundText = Dtc_partida15.BoundText
    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
End Sub

Private Sub dtc_desc15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_desc15.BoundText
    dtc_unimed15.BoundText = dtc_desc15.BoundText
    dtc_stocktotal15.BoundText = dtc_desc15.BoundText
    dtc_grupo15.BoundText = dtc_desc15.BoundText
    dtc_subgrupo15.BoundText = dtc_desc15.BoundText
    Dtc_partida15.BoundText = dtc_desc15.BoundText
    dtc_precioventafinal15.BoundText = dtc_desc15.BoundText
    dtc_precioventabase15.BoundText = dtc_desc15.BoundText
    dtc_preciocompra15.BoundText = dtc_desc15.BoundText
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

Private Sub dtc_desc2_LostFocus()
    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.backColor = &HFF&
'    Else
'        Dtc_deudor2.backColor = &H80000010
'    End If

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

Private Sub dtc_desc15_LostFocus()
    txt_descripcion_venta.Text = dtc_desc15.Text
    TxtDescuento.Text = "0"
    TxtPrecioU.Text = dtc_precioventabase15.Text
    Call AbreAlmacen
End Sub

Private Sub dtc_codigo12_Click(Area As Integer)
    Dtc_aux12.BoundText = dtc_codigo12.BoundText
    dtc_desc12.BoundText = dtc_codigo12.BoundText
End Sub

Private Sub dtc_grupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_grupo15.BoundText
    dtc_desc15.BoundText = dtc_grupo15.BoundText
    dtc_unimed15.BoundText = dtc_grupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_grupo15.BoundText
    dtc_subgrupo15.BoundText = dtc_grupo15.BoundText
    Dtc_partida15.BoundText = dtc_grupo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_grupo15.BoundText
    dtc_precioventabase15.BoundText = dtc_grupo15.BoundText
    dtc_preciocompra15.BoundText = dtc_grupo15.BoundText
End Sub

Private Sub dtc_stocktotal15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_stocktotal15.BoundText
    dtc_desc15.BoundText = dtc_stocktotal15.BoundText
    dtc_unimed15.BoundText = dtc_stocktotal15.BoundText
    dtc_grupo15.BoundText = dtc_stocktotal15.BoundText
    dtc_subgrupo15.BoundText = dtc_stocktotal15.BoundText
    Dtc_partida15.BoundText = dtc_stocktotal15.BoundText
    dtc_precioventafinal15.BoundText = dtc_stocktotal15.BoundText
    dtc_precioventabase15.BoundText = dtc_stocktotal15.BoundText
    dtc_preciocompra15.BoundText = dtc_stocktotal15.BoundText
End Sub

Private Sub Dtc_aux12_Click(Area As Integer)
    dtc_codigo12.BoundText = Dtc_aux12.BoundText
    dtc_desc12.BoundText = Dtc_aux12.BoundText
End Sub

Private Sub dtc_precioventafinal15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_precioventafinal15.BoundText
    dtc_desc15.BoundText = dtc_precioventafinal15.BoundText
    dtc_unimed15.BoundText = dtc_precioventafinal15.BoundText
    dtc_grupo15.BoundText = dtc_precioventafinal15.BoundText
    dtc_subgrupo15.BoundText = dtc_precioventafinal15.BoundText
    Dtc_partida15.BoundText = dtc_precioventafinal15.BoundText
    dtc_stocktotal15.BoundText = dtc_precioventafinal15.BoundText
    dtc_precioventabase15.BoundText = dtc_precioventafinal15.BoundText
    dtc_preciocompra15.BoundText = dtc_precioventafinal15.BoundText
End Sub

Private Sub dtc_preciocompra15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_preciocompra15.BoundText
    dtc_desc15.BoundText = dtc_preciocompra15.BoundText
    dtc_unimed15.BoundText = dtc_preciocompra15.BoundText
    dtc_stocktotal15.BoundText = dtc_preciocompra15.BoundText
    dtc_grupo15.BoundText = dtc_preciocompra15.BoundText
    dtc_subgrupo15.BoundText = dtc_preciocompra15.BoundText
    Dtc_partida15.BoundText = dtc_preciocompra15.BoundText
    dtc_precioventafinal15.BoundText = dtc_preciocompra15.BoundText
    dtc_precioventabase15.BoundText = dtc_preciocompra15.BoundText
End Sub

Private Sub dtc_stock13_Click(Area As Integer)
    dtc_codigo13.BoundText = Dtc_Stock13.BoundText
    dtc_desc13.BoundText = Dtc_Stock13.BoundText
End Sub

Private Sub dtc_desc12_Click(Area As Integer)
    Dtc_aux12.BoundText = dtc_desc12.BoundText
    dtc_codigo12.BoundText = dtc_desc12.BoundText
End Sub

Private Sub dtc_desc12_LostFocus()
'  If GlSistema = "A" Then       'Or GlSistema = "Z"
'    If dtc_codigo12.Text = "10" Then
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    End If
'  Else
'    'If lblventa_tipo.Caption = "E" Then
'    '    TxtPrecioU.Text = dtc_precioventafinal15.Text
'    'Else
'    '    TxtPrecioU.Text = dtc_precioventabase15.Text
'    'End If
'    If Val(dtc_codigo12.Text) > 19 Then
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 100 Then
'        TxtPrecioU.Text = dtc_preciocompra15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 200 Then
'        TxtPrecioU.Text = "0"
'    End If
'  End If

End Sub

Private Sub dtc_unimed15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_unimed15.BoundText
    dtc_desc15.BoundText = dtc_unimed15.BoundText
    dtc_stocktotal15.BoundText = dtc_unimed15.BoundText
    dtc_grupo15.BoundText = dtc_unimed15.BoundText
    dtc_subgrupo15.BoundText = dtc_unimed15.BoundText
    Dtc_partida15.BoundText = dtc_unimed15.BoundText
    dtc_precioventafinal15.BoundText = dtc_unimed15.BoundText
    dtc_precioventabase15.BoundText = dtc_unimed15.BoundText
    dtc_preciocompra15.BoundText = dtc_unimed15.BoundText
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
End Sub


Private Sub DTPFechaIni_Change()
cmb_mes_ini.Text = UCase(MonthName(Month(DTPFechaIni.Value)))
txt_mes.Text = Month(DTPFechaIni.Value)
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
'    'Ado_datos4.Refresh
'
'    DTPfechaIni.Value = DTPfechasol.Value
''    'validar fecha solicitud OJO JQA 31/12/2014
''    Set rs_TipoCambio = New ADODB.Recordset
''    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
''    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
''    If rs_TipoCambio.RecordCount > 0 Then
''        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
''    End If
'End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    modif = ""
    VAR_SW = ""
    ado = "SI"
    lbl_cerrado = ""
    parametro = Aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'Call ABRIR_TABLA_DET
    'Call ABRIR_TABLA
    'Call ABRIR_TABLA_AUX2
    'Call ABRIR_TABLA_DET3
    'txt_codigo.Enabled = True
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
    BtnImprimir2.Visible = False
'    BtnImprimir3.Visible = False

'    FrmEdita.Enabled = False
'    FrmCobros.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
    If parametro = "DNREP" Or parametro = "DNEME" Then
        lbl_cite = "Cite / O.S."
    Else
        lbl_cite = "Cite Trámite"
    End If
'    Chk_plazo.Value = 0
End Sub

Private Sub ABRIR_TABLA_DET()
   ' Dim rs_datos1 As New ADODB.Recordset
   If rs_datos.RecordCount > 0 Then
   If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select *,upper(datename(month,cobranza_fecha_prog))as fecha_pla, year(cobranza_fecha_prog) as gestion from ro_prestamo_prog where prestamo_codigo = '" & Ado_datos.Recordset!prestamo_codigo & "' AND beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ORDER BY prestamo_prog_codigo ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_datos1.Recordset = rs_datos1
    Set dg_det1.DataSource = Ado_datos1.Recordset
    'If Ado_datos1.Recordset.RecordCount > 0 Then
     If rs_datos1.RecordCount > 0 Then
    Call calcular_total
    End If
   End If
    Else
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select * FROM ro_prestamo_prog where  beneficiario_codigo = '" & "WWW" & "' ORDER BY prestamo_prog_codigo ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_datos1.Recordset = rs_datos1
    Set dg_det1.DataSource = Ado_datos1.Recordset
    End If

End Sub
Private Sub ABRIR_TABLAS_AUX()
    'UNIDAD EJECUTORA
    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "Select * from gc_unidad_ejecutora WHERE unidad_codigo = 'DRRHH'", db, adOpenStatic
    'rs_datos18.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos16.Recordset = rs_datos18
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_desc1.Text = rs_datos18!unidad_descripcion
    DataCombo2.Text = rs_datos18!unidad_descripcion

    'Beneficiario Personas Nat. y Juridicas
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    rs_datos2.Open "av_ro_peronal_vs_gc_beneficiario where estado_codigo <> 'ANL' order by beneficiario_denominacion", db, adOpenStatic
    If rs_datos2.RecordCount > 0 Then
    End If
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText

    'Beneficiario Funcionario - Vendedor
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & "DRRHH" & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    If rs_datos4.RecordCount > 0 Then
    dtc_desc4.Text = rs_datos4!beneficiario_denominacion
    dtc_codigo4.Text = rs_datos4!beneficiario_codigo
    End If

'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'    dtc_aux4.BoundText = dtc_codigo4.BoundText
'
'    'Beneficiario Funcionario - Cobrador por Pago
'    Set rs_datos4A = New ADODB.Recordset
'    If rs_datos4A.State = 1 Then rs_datos4A.Close
'    rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
'    Set ado_datos4A.Recordset = rs_datos4A
''    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
'
'    'Beneficiario Funcionario - Cobrador por Contrato
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    rs_datos5.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' order by beneficiario_denominacion", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'
''    Set rs_datos6 = New ADODB.Recordset
''    If rs_datos6.State = 1 Then rs_datos6.Close
''    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
''    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
''    Set Ado_datos6.Recordset = rs_datos6
''    dtc_desc6.BoundText = dtc_codigo6.BoundText
''
''    Set rs_datos7 = New ADODB.Recordset
''    If rs_datos7.State = 1 Then rs_datos7.Close
''    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
''    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
''    Set Ado_datos7.Recordset = rs_datos7
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
''
''    Set rs_datos8 = New ADODB.Recordset
''    If rs_datos8.State = 1 Then rs_datos8.Close
''    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
''    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
''    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
''
''    Set rs_datos9 = New ADODB.Recordset
''    If rs_datos9.State = 1 Then rs_datos9.Close
''    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
''    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
''    Set Ado_datos9.Recordset = rs_datos9
''    dtc_desc9.BoundText = dtc_codigo9.BoundText
''
''    Set rs_datos10 = New ADODB.Recordset
''    If rs_datos10.State = 1 Then rs_datos10.Close
''    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
''    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
''    Set Ado_datos10.Recordset = rs_datos10
''    dtc_desc10.BoundText = dtc_codigo10.BoundText
'
'    'ac_tipo_compra_venta
'    Set rs_datos11 = New ADODB.Recordset
'    If rs_datos11.State = 1 Then rs_datos11.Close
'    rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo <> 'L' and venta_tipo <> 'V' ", db, adOpenStatic
'    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'
'    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh
'
'    'Solo para Equipos (*)
'    Set rs_datos15 = New ADODB.Recordset
'    If rs_datos15.State = 1 Then rs_datos15.Close
'    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
'    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
'    Set ado_datos15.Recordset = rs_datos15
'    ado_datos15.Refresh
'
'   'wwwwwwwwwwwwwwwwwwww
'    'db.Execute "DELETE ao_ventas_cabecera where venta_c=odigo = 0 "
'    'Call ABREVENTAS
'
''    Set rs_Dsctos  New ADODB.Recordset
''    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
''    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
''    Set AdoDsctos.Recordset = rs_Dsctos
''    AdoDsctos.Refresh
'
'    Set rs_datos17 = New ADODB.Recordset
'    If rs_datos17.State = 1 Then rs_datos17.Close
'    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
'    Set ado_datos17.Recordset = rs_datos17
'    ado_datos17.Refresh
'
''WWWWWWWWWWWWWWWWWWWWWWWWWWWW

End Sub



Private Sub grabar()
  'db.BeginTrans
    If swgrabar = 1 Then
'      Dim rstdestino As New ADODB.Recordset
'      Set rstdestino = New ADODB.Recordset
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select tipo_tramite, numero_correlativo from fc_correl WHERE tipo_tramite='ventas'", db, adOpenDynamic, adLockOptimistic
'      If rstdestino.RecordCount <> 0 Then
'        Ado_datos.Recordset("venta_codigo") = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino!numero_correlativo = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino.Update
'      Else
'        Ado_datos.Recordset("venta_codigo") = 1
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      'Ado_datos.Recordset("venta_codigo") = Ado_datos.Recordset.RecordCount
'      'rstdestino.AddNew
    End If
       Ado_datos.Recordset("ges_gestion") = glGestion       'CStr(Year(DTPfechasol.Value))
       Ado_datos.Recordset("unidad_codigo") = dtc_codigo1.Text
       Ado_datos.Recordset("solicitud_codigo") = txt_codigo.Caption
       Ado_datos.Recordset("edif_codigo") = dtc_codigo3.Text
       Ado_datos.Recordset("depto_codigo") = Left(dtc_codigo3.Text, 1)
       Ado_datos.Recordset("venta_fecha") = Format(Date, "dd,mm,yyyy") 'DTPfechasol.Value
       Ado_datos.Recordset("venta_fecha_inicio") = DTPFechaIni.Value
       Ado_datos.Recordset("venta_fecha_fin") = DTPFechaFin.Value
       Ado_datos.Recordset("venta_tipo") = dtc_codigo11.Text                'E=Efectivo, C=Credito
       Ado_datos.Recordset("beneficiario_codigo") = dtc_codigo2.Text        'CLIENTE
       Ado_datos.Recordset("beneficiario_codigo_resp") = dtc_codigo4.Text   'Vendedor
       Ado_datos.Recordset("beneficiario_codigo_resp") = dtc_codigo5.Text   'Cobrador
       Ado_datos.Recordset("venta_descripcion") = Trim(TxtConcepto.Text)
       'CONT2 = 365 / 30 * Ado_datos.Recordset!venta_cantidad_total
       'Ado_datos.Recordset("venta_plazo_dias_calendario") = IIf(TxtPlazo.Text = "", CONT2, TxtPlazo.Text)
       Ado_datos.Recordset("venta_tipo_cambio") = GlTipoCambioMercado        'Val(txtTDC.Text)venta_tipo_cambio
        'GlTipoCambioOficial As Currency        'GlTipoCambioMercado As Currency        'GlTipoCambioGestion As Currency
       Ado_datos.Recordset("tipoben_codigo") = IIf(Dtc_aux2.Text = "", "2", Dtc_aux2.Text)      'Tipo de Beneficiario
        
       Ado_datos.Recordset("unimed_codigo_cobr") = cmd_unimed2.Text
       Ado_datos.Recordset("venta_cantidad_cobr") = txtCantCobr.Text
       Ado_datos.Recordset("mes_inicio_crono") = RTrim(cmb_mes_ini.Text)
       VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
       VAR_COBR2 = Ado_datos.Recordset!venta_cantidad_cobr
       MControl = Ado_datos.Recordset!mes_inicio_crono
       
       Ado_datos.Recordset("proceso_codigo") = "TEC"
       Ado_datos.Recordset("subproceso_codigo") = "TEC-02"
       Ado_datos.Recordset("etapa_codigo") = "TEC-02-02"
       Ado_datos.Recordset("clasif_codigo") = "TEC"
       Ado_datos.Recordset("doc_codigo") = "R-302"
       'Ado_datos.Recordset("doc_numero") = "0"
       Ado_datos.Recordset("poa_codigo") = "3.2.3"
     Select Case dtc_codigo2.Text
        Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL - Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "CMX-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
            rs_datos!proceso_codigo = "CMX"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "CMX-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "CMX-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-XXX"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-02"    '3. COMPRA-VENTA BB Y SS - COMERCIAL -         Case "2"    'SOLO VENTA DE BIENES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-03"    'VENTA DE SERVICIOS INSTTALACIONES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-04" '5       'VENTA DE SERVICIOS AJUSTE
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "TEC-01"    '6. SERVICIO MANTENIMIENTO GRATUITO
            rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "TEC-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "TEC-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "TEC-02"    '10. SERVICIO MANTENIMIENTO PREVENTIVO
            'If VAR_UNI = "DNMAN" Then
            'rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = Left(dtc_codigo2.Text, 3) ' "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = Trim(dtc_codigo2.Text) + "-01"  '"TEC-02-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = Left(dtc_codigo2.Text, 3)  '"TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-302"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'End If
        Case "TEC-03" '10 REPARACION    If VAR_UNI = "DNREP" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TRC-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
                rs_datos!unidad_codigo_ant = Txt_campo2.Text
                VAR_CITE = Txt_campo2.Text
        Case "TEC-04" '10 EMERGENCIAS   If VAR_UNI = "DNEME" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
                rs_datos!unidad_codigo_ant = Txt_campo2.Text
                VAR_CITE = Txt_campo2.Text
        Case "TEC-05"    '5. SERVICIO MODERNIZACION -If VAR_UNI = "DNMOD" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
     End Select

'       'If Ado_datos.Recordset("venta_tipo") = "E" Then
'       '     Ado_datos.Recordset("monto_cobrado") = IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'       '     Ado_datos.Recordset("deuda_cobrada") = IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'       '  Else
'            Ado_datos.Recordset("monto_cobrado") = "0"
'            Ado_datos.Recordset("deuda_cobrada") = "0"
'       'End If
'       If swgrabar = 1 Then
'         Ado_datos.Recordset("cantidad_total_vendida") = 0
'         Ado_datos.Recordset("monto_total_bS") = 0  'IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'         Ado_datos.Recordset("monto_total_Us") = 0
'       End If
'       Ado_datos.Recordset("saldo_p_cobrar") = Ado_datos.Recordset("monto_total_bS") - Ado_datos.Recordset("deuda_cobrada")

       Ado_datos.Recordset("estado_codigo") = "REG"

       Ado_datos.Recordset("usr_codigo") = glusuario
       Ado_datos.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
       Ado_datos.Recordset("hora_registro") = Format(Time, "hh/mm/ss")
       'Ado_datos.Recordset("usuario_aprueba") = ""
        'Ado_datos.Recordset("fecha_aprueba") = ""

    Ado_datos.Recordset.Update
        'GENERA CORREL CONTRATO POR DEPTO INI
        Set rs_aux7 = New ADODB.Recordset
        If rs_aux7.State = 1 Then rs_aux7.Close
        rs_aux7.Open "Select correl_contrato as Codigo from gc_departamento where depto_codigo = '" & Left(Ado_datos.Recordset!edif_codigo, 1) & "'    ", db, adOpenStatic        'VAR_PROY2
        If Not rs_aux7.EOF Then
            'VAR_CONTR = IIf(IsNull(rs_aux7!Codigo), 1, CDbl(rs_aux7!Codigo) + 1)
            If IsNull(rs_aux7!Codigo) Then
                VAR_CONTR = 1
            Else
                VAR_CONTR = IIf(IsNull(rs_aux7!Codigo), 1, CDbl(rs_aux7!Codigo) + 1)
            End If
        End If
        db.Execute "update ao_ventas_cabecera set venta_codigo_new = " & VAR_CONTR & " Where ao_ventas_cabecera.venta_codigo = " & var_cod5 & "  And ao_ventas_cabecera.ges_gestion = " & gestion0 & " "
        db.Execute "update gc_departamento set correl_contrato = " & VAR_CONTR & " Where depto_codigo = '" & Left(VAR_PROY2, 1) & "' "
        'GENERA CORREL CONTRATO POR DEPTO FIN

    'Ado_datos.Recordset.Requery
    'If rstdestino.State = 1 Then rstdestino.Close
    'db.CommitTrans
    If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       If Ado_datos.Recordset("venta_tipo") = "E" Then
           db.Execute "INSERT INTO ao_ventas_cobranza_prog (venta_codigo, ges_gestion, beneficiario_codigo, beneficiario_codigo_resp, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, literal, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, factura_impresa, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
           "VALUES ('" & Ado_datos.Recordset!venta_codigo & "', '" & glGestion & "', '" & Ado_datos.Recordset!beneficiario_codigo & "', '" & Ado_datos.Recordset!beneficiario_codigo_resp & "', " & Ado_datos.Recordset!venta_monto_total_bs & ", '" & Ado_datos.Recordset!venta_monto_total_dol & "', '0', '0', " & Ado_datos.Recordset!venta_monto_total_bs & ", " & Ado_datos.Recordset!venta_monto_total_dol & ", '" & Date & "', '" & Date & "', 'CANCELADO', 'CERO', 'COM', 'COM-02', 'COM-02-02', 'ADM', 'R-103', '0', 'R-101', '0', '0', 'N', '3.1.2', 'REG', '" & glusuario & "', '" & Date & "', '09:00')"
           '  cobranza_codigo       'Especif. de Identidad
       End If
       'Call OptFilGral1_Click
       'Ado_datos.Refresh
       'Ado_datos.Recordset.Move marca1 - 1
        If swgrabar = 1 Then
            Ado_datos.Refresh
            Ado_datos.Recordset.MoveLast
        End If
    End If

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
    TxtDescuento.Text = "0"
    Txt_modelo.Text = Txt_modelo1.Text
    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    If rs_datos18.RecordCount > 0 Then
        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
    End If
    'Set ado_datos17.Recordset = rs_datos18
    'ado_datos17.Refresh
End Sub

Private Sub OpMod2_Click()
    TxtDescuento.Text = "0"
    Txt_modelo.Text = Txt_modelo2.Text
    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    If rs_datos18.RecordCount > 0 Then
        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_h
    End If
End Sub

Private Sub OpMod3_Click()
    TxtDescuento.Text = "0"
    Txt_modelo.Text = Txt_modelo3.Text
    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    If rs_datos18.RecordCount > 0 Then
        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_x
    End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From rv_prestamos_cebacera WHERE estado_codigo = 'REG'"
    
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    
    rs_datos.Sort = "prestamo_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    If rs_datos.RecordCount > 0 Then
    Ado_datos.Recordset.MoveFirst
    End If
    
    Call ABRIR_TABLA_DET
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From rv_prestamos_cebacera"
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    
    rs_datos.Sort = "prestamo_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    If rs_datos.RecordCount > 0 Then
    Ado_datos.Recordset.MoveFirst
    End If
    
    Call ABRIR_TABLA_DET
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
  txtTDC.Text = GlTipoCambioOficial

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
  'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as VAR_COBR2 from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo = '43340'", db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!cantot0
  End If

  'rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "

  TxtMontoBs.Text = VAR_AUX
  TxtCobrado.Text = Cobrobs
  TxtBstotal.Text = VAR_Bs

'  If IsNull(Ado_datos.Recordset!venta_monto_cobrado_bs) Then
'    Ado_datos.Recordset!venta_monto_cobrado_bs = 0
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs
'  Else
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs
'  End If
'  If VAR_AUX > 0 Then
'        VAR_AUX2 = VAR_AUX / Ado_datos.Recordset!venta_tipo_cambio
'  Else
'        VAR_AUX2 = 0
'  End If
'  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas_cabecera.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas_cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas_cabecera.saldo_p_cobrar = ao_ventas_cabecera.monto_total_Bs - ao_ventas_cabecera.deuda_cobrada Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.venta_monto_total_dol = " & rstacumdet!totdl & ", ao_ventas_cabecera.venta_cantidad_total = " & rstacumdet!cantot & ", ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_AUX & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_AUX2 & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'
'  TxtMontoBs = rstacumdet!totbs
'  TxtCobrado = rs_datos19!totbs2    'IIf(IsNull(Ado_datos.Recordset("venta_monto_cobrado_bs")), 0, Ado_datos.Recordset("venta_monto_cobrado_bs"))
'  If IsNull(Ado_datos.Recordset("venta_saldo_p_cobrar_bs")) Then
'    Text2 = VAR_AUX 'Ado_datos.Recordset("venta_monto_total_bs") - Ado_datos.Recordset("venta_monto_cobrado_bs")
'    Ado_datos.Recordset("venta_saldo_p_cobrar_bs") = VAR_AUX
'  Else
'    Text2 = Ado_datos.Recordset("venta_saldo_p_cobrar_bs")
'  End If

  If rstacumdet.State = 1 Then rstacumdet.Close

  'Print ado_datos14.Recordset!ges_gestion
  'Print ado_datos14.Recordset!correl_venta
  'Print ado_datos14.Recordset!venta_codigo
  'ado_datos14.Recordset!monto_Bolivianos = rstacumdet!totbs
  'ado_datos14.Recordset!monto_dolares = rstacumdet!totdl
  'ado_datos14.Recordset.Update
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & ges & "' and correl_venta = '" & corr & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount > 0 Then
'    rstdestino!monto_total_Bs = rstacumdet!totbs
'    rstdestino!monto_cobrado = rstacumdet!totbs
'    rstdestino!monto_total_Us = rstacumdet!totdl
'    rstdestino!cantidad_total_vendida = rstacumdet!cantot
'    rstdestino!saldo_p_cobrar = 0
'    rstdestino.Update
'  End If
'  'Set Ado_datos.Recordset = rstdestino
'  If rstdestino.State = 1 Then rstdestino.Close
'  If rstacumdet.State = 1 Then rstacumdet.Close
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = False
    Else
'           FrmEditaDet.Visible = False
'           DtGLista.Visible = False
'           adoao_solicitud_lista.Visible = False
    End If

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

Private Sub Picture2_Click()
If Text8.Text = "" Or Text8.Text = "0.00" Then
sino = MsgBox("Tiene que ingresar el " & lblLabels(5).Caption, vbCritical, "SOFIA")
Exit Sub
End If
If modif_det = "SI" Then
rs_datos1!cobranza_programada_bs = Text8.Text
rs_datos1!cobranza_observaciones = Text16.Text
rs_datos1.Update

Else

    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "select  max(prestamo_prog_codigo)as ULTIMA_CUOTA from ro_prestamo_prog where prestamo_codigo = '" & Ado_datos.Recordset!prestamo_codigo & "' AND beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    
    CUOTA = 0
    CUOTA = rs_aux8!ULTIMA_CUOTA + 1
    
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "select * from ro_prestamo_prog where prestamo_codigo = '" & Ado_datos.Recordset!prestamo_codigo & "' AND beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' AND prestamo_prog_codigo = " & (CUOTA - 1), db, adOpenKeyset, adLockOptimistic, adCmdText


    
    
    

'rs_datos1.AddNew


'rs_datos!cantidad_periodos = cuota

mes_grupo = Month(rs_aux8!cobranza_fecha_prog)
gestion = Year(rs_aux8!cobranza_fecha_prog)
Select Case rs_datos!unimed_codigo
Case "MENSUAL"
  mes_grupo = mes_grupo + 1
Case "BIMESTRAL"
  mes_grupo = mes_grupo + 2
Case "SEMESTRAL"
  mes_grupo = mes_grupo + 6
Case "ANUAL"
   mes_grupo = mes_grupo + 12
End Select

If mes_grupo > 12 Then
mes_grupo = mes_grupo - 12
gestion = gestion + 1
End If


'If mes_grupo < Month(rs_datos!fecha_inicio) Then
'gestion = gestion + 1
'End If




rs_datos1!prestamo_codigo = Ado_datos.Recordset!prestamo_codigo
rs_datos1!prestamo_prog_codigo = CUOTA
rs_datos1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
rs_datos1!beneficiario_codigo_resp = Ado_datos.Recordset!beneficiario_codigo_resp
rs_datos1!cobranza_programada_bs = Text8.Text
rs_datos1!cobranza_programada_dol = rs_datos1!cobranza_programada_bs / GlTipoCambioOficial
rs_datos1!cobranza_descuento_bs = "0"
rs_datos1!cobranza_descuento_dol = "0"
rs_datos1!cobranza_total_bs = rs_datos!prestamo_total_bs
rs_datos1!cobranza_total_dol = rs_datos!prestamo_total_bs / GlTipoCambioOficial

dia = Day(rs_datos!fecha_inicio)
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion

If Weekday(rs_datos1!cobranza_fecha_prog, vbMonday) = 6 Then  'Pregunto si es Domingo
dia = dia + 1
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
ElseIf Weekday(rs_datos1!cobranza_fecha_prog, vbMonday) = 7 Then 'Pregunto si es Sabado
dia = dia + 1
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
Else
rs_datos1!cobranza_fecha_prog = dia & "/" & mes_grupo & "/" & gestion
End If
rs_datos!fecha_fin = rs_datos1!cobranza_fecha_prog



'rs_datos1!cobranza_fecha_cobro = ""
rs_datos1!mes_planilla = mes_grupo
If Text16.Text = "" Then
rs_datos1!cobranza_observaciones = "PAGO NUMERO " & CUOTA & " - PERIODO: " & gestion & " " & UCase(MonthName(mes_grupo))
Else
rs_datos1!cobranza_observaciones = Text16.Text
End If
End If
rs_datos1!trans_codigo = "P"
rs_datos1!estado_codigo = "REG"
rs_datos1!usr_codigo = glusuario
rs_datos1!Fecha_Registro = Date



rs_datos1.Update
rs_datos.Update
Fra_crono.Visible = False
Call ABRIR_TABLA_DET
FraNavega.Enabled = True
FrmCobranza.Enabled = True
FrmABMDet2.Enabled = True
fraOpciones.Enabled = True

End Sub

Private Sub Picture3_Click()
Fra_crono.Visible = False
rs_datos1.CancelBatch
FraNavega.Enabled = True
FrmCobranza.Enabled = True
FrmABMDet2.Enabled = True
fraOpciones.Enabled = True
If modif_det = "NO" Then
rs_datos1.MoveLast
End If

End Sub

Private Sub Text1_Change()

'TxtCobrado.Text = Val(Text1.Text) - Val(TxtMontoBs.Text)

End Sub


Private Sub txt_interes_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub txt_monto_bs_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub txtCantCobr_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Then      '(KeyAscii = 8) Or '(0..9)
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoBs_Change()
'TxtCobrado.Text = Val(Text1.Text) - Val(TxtMontoBs.Text)
End Sub

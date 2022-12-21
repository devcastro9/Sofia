VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form mw_solicitud_calculo_trafico_mod_DET 
   Caption         =   "Formulario Datos Tecnicos para Modernización"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":0000
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20520
      TabIndex        =   112
      Top             =   0
      Width           =   20520
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5640
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   120
         ToolTipText     =   "Imprime Lista de Cronogramas"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":12CF
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   119
         ToolTipText     =   "Busca un Registro"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":1A84
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   118
         ToolTipText     =   "Aprueba el Cronograma"
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":22B7
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   117
         ToolTipText     =   "Anula el Registro Activo"
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
         Left            =   105
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":2A03
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   116
         ToolTipText     =   "Editar Datos de ""Cabecera Cronograma"""
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8760
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":3318
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   115
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":3AD7
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   114
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7200
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":4299
         ScaleHeight     =   615
         ScaleWidth      =   1440
         TabIndex        =   113
         ToolTipText     =   "Ver Cálculos de Tráfico"
         Top             =   0
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   12600
         TabIndex        =   121
         Top             =   180
         Width           =   885
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
      TabIndex        =   108
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2955
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":4DA5
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   110
         Top             =   0
         Width           =   1395
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         Picture         =   "mw_solicitud_calculo_trafico_mod_DET.frx":5691
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   109
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   12600
         TabIndex        =   111
         Top             =   195
         Width           =   1035
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DATOS GENERALES"
      TabPicture(0)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5E67
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MAQUINA DE TRACCION"
      TabPicture(1)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5E83
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CABINA"
      TabPicture(2)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5E9F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PUERTAS"
      TabPicture(3)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5EBB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CONTROL"
      TabPicture(4)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5ED7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "C.O.P."
      TabPicture(5)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5EF3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "SEÑALIZACION DE PISO"
      TabPicture(6)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F0F
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   85
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text31 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   107
            Text            =   "0"
            Top             =   5280
            Width           =   2145
         End
         Begin VB.TextBox Text30 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   106
            Text            =   "0"
            Top             =   4680
            Width           =   2145
         End
         Begin VB.TextBox Text29 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   105
            Text            =   "0"
            Top             =   4080
            Width           =   2145
         End
         Begin VB.TextBox Text28 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   104
            Text            =   "0"
            Top             =   3480
            Width           =   2145
         End
         Begin VB.TextBox Text27 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   103
            Text            =   "0"
            Top             =   2880
            Width           =   2145
         End
         Begin VB.TextBox Text26 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   102
            Text            =   "NO"
            Top             =   5880
            Width           =   705
         End
         Begin VB.TextBox Text25 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   87
            Text            =   "NO"
            Top             =   1680
            Width           =   705
         End
         Begin VB.TextBox Text24 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   86
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo19 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F2B
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5640
            TabIndex        =   88
            Top             =   480
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo20 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F46
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   89
            Top             =   480
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo21 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F60
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6000
            TabIndex        =   90
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo22 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F7B
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   91
            Top             =   2280
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Localizacion Señalizaciòn de Piso"
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
            TabIndex        =   101
            Top             =   5280
            Width           =   3060
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Gongo"
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
            TabIndex        =   100
            Top             =   5880
            Width           =   615
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Señalización Piso"
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
            TabIndex        =   99
            Top             =   4680
            Width           =   2070
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Localizacion Botonera"
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
            TabIndex        =   98
            Top             =   4080
            Width           =   1995
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Color Iliminacion Boton Llamada"
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
            TabIndex        =   97
            Top             =   3480
            Width           =   2880
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Boton de Llamada"
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
            TabIndex        =   96
            Top             =   2880
            Width           =   2400
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Display"
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
            TabIndex        =   95
            Top             =   2280
            Width           =   1440
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Botonería con Display Integrado"
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
            TabIndex        =   94
            Top             =   1680
            Width           =   2880
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de Carreras"
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
            TabIndex        =   93
            Top             =   1080
            Width           =   1920
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Estetica Botonería y Señalización"
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
            TabIndex        =   92
            Top             =   480
            Width           =   2985
         End
      End
      Begin VB.Frame Frame6 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text23 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   675
            Left            =   3240
            TabIndex        =   84
            Text            =   "0"
            Top             =   3480
            Width           =   5145
         End
         Begin VB.TextBox Text21 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   83
            Text            =   "0"
            Top             =   2880
            Width           =   2145
         End
         Begin VB.TextBox Text22 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   70
            Text            =   "0"
            Top             =   480
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo13 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5F95
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   71
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo14 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5FB0
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   72
            Top             =   2280
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo15 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5FCA
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   73
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo16 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5FE5
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   74
            Top             =   1080
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo17 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":5FFF
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   79
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo18 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":601A
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   80
            Top             =   1680
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Otros Opcionales POC"
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
            TabIndex        =   82
            Top             =   3480
            Width           =   2025
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Color Iliminacion Boton Llamada"
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
            TabIndex        =   81
            Top             =   2880
            Width           =   2880
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad COP's"
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
            TabIndex        =   78
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Señalizacion del POC"
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
            TabIndex        =   77
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Insertos de POC"
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
            TabIndex        =   76
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Boton de Llamada"
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
            TabIndex        =   75
            Top             =   2280
            Width           =   2400
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   58
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text20 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   675
            Left            =   2160
            TabIndex        =   60
            Text            =   "0"
            Top             =   2280
            Width           =   5145
         End
         Begin VB.TextBox Text17 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2160
            TabIndex        =   59
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo9 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6034
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4800
            TabIndex        =   61
            Top             =   480
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo10 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":604F
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2160
            TabIndex        =   62
            Top             =   480
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo11 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6069
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4800
            TabIndex        =   63
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo12 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6084
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2160
            TabIndex        =   64
            Top             =   1680
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Opcionales"
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
            TabIndex        =   68
            Top             =   2280
            Width           =   1035
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Comando"
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
            TabIndex        =   67
            Top             =   1680
            Width           =   1635
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo"
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
            TabIndex        =   66
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tecnología Control"
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
            TabIndex        =   65
            Top             =   480
            Width           =   1710
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   47
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text19 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   49
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin VB.TextBox Text18 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   48
            Text            =   "0"
            Top             =   1680
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":609E
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5640
            TabIndex        =   50
            Top             =   480
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":60B9
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   51
            Top             =   480
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":60D3
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5640
            TabIndex        =   56
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":60EE
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3000
            TabIndex        =   57
            Top             =   2280
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo Operador"
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
            TabIndex        =   55
            Top             =   480
            Width           =   1605
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Apertura Libre de Puerta (mm)"
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
            TabIndex        =   54
            Top             =   1080
            Width           =   2670
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Altura Libre de Puerta (mm)"
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
            TabIndex        =   53
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Puerta Piso"
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
            TabIndex        =   52
            Top             =   2280
            Width           =   1500
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   29
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text16 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   45
            Text            =   "NO"
            Top             =   4320
            Width           =   2145
         End
         Begin VB.TextBox Text15 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   555
            Left            =   3480
            TabIndex        =   35
            Text            =   "0"
            Top             =   2880
            Width           =   3705
         End
         Begin VB.TextBox Text14 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   34
            Text            =   "0"
            Top             =   2280
            Width           =   2145
         End
         Begin VB.TextBox Text13 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   33
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin VB.TextBox Text12 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   32
            Text            =   "0"
            Top             =   1680
            Width           =   2145
         End
         Begin VB.TextBox Text11 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   31
            Text            =   "NO"
            Top             =   3720
            Width           =   2145
         End
         Begin VB.TextBox Text10 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   30
            Text            =   "0"
            Top             =   480
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6108
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   36
            Top             =   4920
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6123
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   37
            Top             =   4920
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Indicador de Sentido de la Cabina"
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
            TabIndex        =   46
            Top             =   4920
            Width           =   3045
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Pasamanos: Color - Fondo - Lado Opuesto COP - Lado COP"
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
            Height          =   480
            Left            =   240
            TabIndex        =   44
            Top             =   2880
            Width           =   3075
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo del Subtecho"
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
            TabIndex        =   43
            Top             =   2280
            Width           =   1920
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Esquinas Redondeadas"
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
            TabIndex        =   42
            Top             =   1680
            Width           =   2205
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Acabado de la Cabina"
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
            TabIndex        =   41
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Adaptado para Accesibilidad (D13)"
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
            TabIndex        =   40
            Top             =   480
            Width           =   3165
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tiene Espejo ?"
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
            TabIndex        =   39
            Top             =   3720
            Width           =   1365
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Ventilador"
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
            TabIndex        =   38
            Top             =   4320
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6495
         Left            =   -74760
         TabIndex        =   13
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text9 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   28
            Text            =   "0"
            Top             =   4080
            Width           =   2145
         End
         Begin VB.TextBox Text8 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   26
            Text            =   "0"
            Top             =   3480
            Width           =   2145
         End
         Begin VB.TextBox Text7 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   17
            Text            =   "0"
            Top             =   1680
            Width           =   2145
         End
         Begin VB.TextBox Text6 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   16
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin VB.TextBox Text5 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   15
            Text            =   "0"
            Top             =   2280
            Width           =   2145
         End
         Begin VB.TextBox Text4 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   14
            Text            =   "0"
            Top             =   2880
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":613D
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5400
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6158
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   19
            Top             =   480
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Corriente (A)"
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
            TabIndex        =   27
            Top             =   4080
            Width           =   1110
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Potencia en KW"
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
            TabIndex        =   25
            Top             =   3480
            Width           =   1425
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo Maquina Traccion"
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
            TabIndex        =   24
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de Cables"
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
            TabIndex        =   23
            Top             =   1080
            Width           =   1785
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Cables de Traccion"
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
            TabIndex        =   22
            Top             =   1680
            Width           =   1770
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Polea de Traccion (mm)"
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
            TabIndex        =   21
            Top             =   2280
            Width           =   2160
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Suspension"
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
            TabIndex        =   20
            Top             =   2880
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text3 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   12
            Text            =   "0"
            Top             =   2880
            Width           =   2145
         End
         Begin VB.TextBox Text2 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   10
            Text            =   "0"
            Top             =   2280
            Width           =   2145
         End
         Begin VB.TextBox Text1 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   5
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin VB.TextBox Txt_campo21 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   2
            Text            =   "0"
            Top             =   480
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo dtc_codigo11 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":6172
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5400
            TabIndex        =   7
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_codigo"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc11 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":618D
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   8
            Top             =   1680
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Piso Principal"
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
            TabIndex        =   11
            Top             =   2880
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Marcación de Pisos"
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
            TabIndex        =   9
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Entradas"
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
            TabIndex        =   6
            Top             =   1680
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad de Entradas"
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
            TabIndex        =   4
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Ultima Altura (mm)"
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
            TabIndex        =   3
            Top             =   480
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "mw_solicitud_calculo_trafico_mod_DET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_datos As New Recordset

Private Sub Form_Load()
    Set rs_datos = New ADODB.Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    ' order by unidad_descripcion
    rs_datos.Open "Select * from av_bienes_eqp_caracteristicas_y_venta ", db, adOpenStatic
    Set Ado_datos.Recordset = rs_datos
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub


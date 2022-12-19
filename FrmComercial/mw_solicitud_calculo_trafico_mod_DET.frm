VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DATOS GENERALES"
      TabPicture(0)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MAQUINA TRACCION"
      TabPicture(1)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CABINA"
      TabPicture(2)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "PUERTAS"
      TabPicture(3)   =   "mw_solicitud_calculo_trafico_mod_DET.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Tab 6"
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.Frame Frame4 
         Height          =   6495
         Left            =   240
         TabIndex        =   49
         Top             =   720
         Width           =   15135
         Begin VB.TextBox Text20 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   53
            Text            =   "0"
            Top             =   480
            Width           =   2145
         End
         Begin VB.TextBox Text19 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   52
            Text            =   "0"
            Top             =   1080
            Width           =   2145
         End
         Begin VB.TextBox Text18 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   51
            Text            =   "0"
            Top             =   2280
            Width           =   2145
         End
         Begin VB.TextBox Text17 
            DataField       =   "trafico_num_paradas"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   50
            Text            =   "0"
            Top             =   2880
            Width           =   2145
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":0070
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5400
            TabIndex        =   54
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":008B
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2760
            TabIndex        =   55
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
         Begin VB.Label Label25 
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
            TabIndex        =   60
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label24 
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
            TabIndex        =   59
            Top             =   1080
            Width           =   1950
         End
         Begin VB.Label Label23 
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
            TabIndex        =   58
            Top             =   1680
            Width           =   1560
         End
         Begin VB.Label Label22 
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
            TabIndex        =   57
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Pisom Principal"
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
            TabIndex        =   56
            Top             =   2880
            Width           =   1395
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
            TabIndex        =   47
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
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":00A5
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   10200
            TabIndex        =   36
            Top             =   5520
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":00C0
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   37
            Top             =   5520
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "ctrlmaq_descripcion"
            BoundColumn     =   "ctrlmaq_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":00DA
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6120
            TabIndex        =   38
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":00F5
            DataField       =   "ctrlmaq_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3480
            TabIndex        =   39
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
            TabIndex        =   48
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":010F
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":012A
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
         Left            =   -74760
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":0144
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
            Bindings        =   "mw_solicitud_calculo_trafico_mod_DET.frx":015F
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
            Caption         =   "Pisom Principal"
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
            Width           =   1395
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_datos As New Recordset

Private Sub Form_Load()
    Set rs_datos = New ADODB.Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    ' order by unidad_descripcion
    rs_datos.Open "Select * from ac_bienes_eqp_caracteristicas ", db, adOpenStatic
    Set Ado_datos.Recordset = rs_datos
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub


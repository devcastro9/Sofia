VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_contratacion_calificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Contratación Personal - Avaluación y Calificación de Proponentes"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_contratacion_calificacion.frx":0000
   ScaleHeight     =   9480
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_contratacion_calificacion.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   9795
      TabIndex        =   24
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   360
         Picture         =   "frm_ao_contratacion_calificacion.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1320
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_contratacion_calificacion.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CALIFICACION DE PROPONENTES"
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
         Left            =   3375
         TabIndex        =   25
         Top             =   240
         Width           =   5190
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   8295
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   9855
      Begin VB.Frame Frame4 
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
         Height          =   3135
         Left            =   165
         TabIndex        =   34
         Top             =   1440
         Width           =   9555
         Begin VB.TextBox Text3 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   9090
            TabIndex        =   76
            Top             =   1280
            Width           =   270
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   7660
            TabIndex        =   75
            Top             =   2650
            Width           =   270
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   9090
            TabIndex        =   72
            Top             =   540
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D6864
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.Ado_detalle2"
            Height          =   315
            Left            =   2040
            TabIndex        =   19
            Top             =   525
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.TextBox txtDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "benef_direccion_domicilio"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6840
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtTelefono 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "benef_telefonos_ref"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8160
            MaxLength       =   20
            TabIndex        =   38
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCI 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D687E
            DataField       =   "ocup_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   3840
            TabIndex        =   68
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "ocup_codigo"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D6898
            DataField       =   "ocup_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   1950
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "ocup_descripcion"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D68B2
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   4920
            TabIndex        =   14
            Top             =   1950
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "nivel_educ_descripcion"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D68CC
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   8640
            TabIndex        =   69
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "nivel_educ_codigo"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D68E6
            DataField       =   "munic_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   3840
            TabIndex        =   70
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D6900
            DataField       =   "munic_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   2655
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D691A
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.Ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   71
            Top             =   525
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D6934
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.Ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   73
            Top             =   1260
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_domicilio_legal"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "frm_ao_contratacion_calificacion.frx":D694E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.Ado_detalle2"
            Height          =   315
            Left            =   4920
            TabIndex        =   74
            Top             =   2640
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_telefono_Cel"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Profesion Principal del Proponente"
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
            TabIndex        =   67
            Top             =   1680
            Width           =   3105
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nivel Educacional (mayor importancia)"
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
            Left            =   4920
            TabIndex        =   66
            Top             =   1680
            Width           =   3465
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Lugar de Postulación"
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
            TabIndex        =   65
            Top             =   2370
            Width           =   1890
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Dirección Postulante"
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
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   41
            Top             =   1010
            Width           =   2115
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Teléfonos Postulante"
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
            Height          =   195
            Index           =   11
            Left            =   4920
            TabIndex        =   39
            Top             =   2370
            Width           =   2010
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Doc.de Identidad "
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
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Apellidos y Nombres del Postulante"
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
            Left            =   2160
            TabIndex        =   36
            Top             =   240
            Width           =   3210
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Elija el Postulante"
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
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   9375
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   9320
         TabIndex        =   64
         Top             =   970
         Width           =   280
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_contratacion_calificacion.frx":D6968
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         Height          =   315
         Left            =   8280
         TabIndex        =   32
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "puesto_codigo"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "apertura_selecionado"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         Height          =   315
         ItemData        =   "frm_ao_contratacion_calificacion.frx":D6982
         Left            =   8280
         List            =   "frm_ao_contratacion_calificacion.frx":D698C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   7200
         Width           =   735
      End
      Begin VB.TextBox txt_obs 
         DataField       =   "observaciones"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         Height          =   405
         Left            =   1680
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   7680
         Visible         =   0   'False
         Width           =   7815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         Caption         =   "Calificacion / Ponderación de Proponentes"
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
         Height          =   1095
         Left            =   165
         TabIndex        =   55
         Top             =   6000
         Visible         =   0   'False
         Width           =   9555
         Begin VB.ComboBox ComboB1 
            DataField       =   "apertura_califica_sobre_b"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6998
            Left            =   1920
            List            =   "frm_ao_contratacion_calificacion.frx":D69BD
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox ComboA1 
            DataField       =   "apertura_califica_sobre_a"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D69E3
            Left            =   240
            List            =   "frm_ao_contratacion_calificacion.frx":D6A08
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox ComboC1 
            DataField       =   "apertura_califica_sobre_c"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6A2E
            Left            =   3720
            List            =   "frm_ao_contratacion_calificacion.frx":D6A53
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox Combo11 
            DataField       =   "apertura_califica_total"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6A79
            Left            =   5640
            List            =   "frm_ao_contratacion_calificacion.frx":D6A9E
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   640
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txtFecha2 
            DataField       =   "cotiza_fecha"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   7200
            TabIndex        =   9
            Top             =   640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   100335617
            CurrentDate     =   41640
            MinDate         =   2
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Fecha.Calificacion.Prop."
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
            Left            =   6960
            TabIndex        =   62
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Perfil Academico"
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
            Left            =   5220
            TabIndex        =   59
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Otras Habilidades"
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
            Left            =   3360
            TabIndex        =   58
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Capacidad Técnica"
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
            Left            =   1440
            TabIndex        =   57
            Top             =   360
            Width           =   1785
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Experiencia"
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
            TabIndex        =   56
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00404040&
         Caption         =   "Recepción y Apertura de Docs. de Proponentes"
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
         Height          =   1095
         Left            =   165
         TabIndex        =   50
         Top             =   4800
         Visible         =   0   'False
         Width           =   9555
         Begin VB.ComboBox ComboA 
            DataField       =   "apertura_sobre_a"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6AC4
            Left            =   240
            List            =   "frm_ao_contratacion_calificacion.frx":D6ACE
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox ComboB 
            DataField       =   "apertura_sobre_b"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6ADA
            Left            =   1920
            List            =   "frm_ao_contratacion_calificacion.frx":D6AE4
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox ComboC 
            DataField       =   "apertura_sobre_c"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6AF0
            Left            =   3720
            List            =   "frm_ao_contratacion_calificacion.frx":D6AFA
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   640
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "apertura_sobre_todos"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            ItemData        =   "frm_ao_contratacion_calificacion.frx":D6B06
            Left            =   5640
            List            =   "frm_ao_contratacion_calificacion.frx":D6B10
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   640
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txtFecha1 
            DataField       =   "cotiza_fecha"
            DataSource      =   "frm_ao_contratacion.ado_detalle1"
            Height          =   315
            Left            =   7200
            TabIndex        =   4
            Top             =   640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   100335617
            CurrentDate     =   41640
            MinDate         =   2
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Fecha.Apertura.Prop."
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
            Left            =   7080
            TabIndex        =   61
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Carta Solicitud"
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
            TabIndex        =   54
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Fotografía"
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
            Left            =   1800
            TabIndex        =   53
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Pretencion Salarial"
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
            Left            =   3240
            TabIndex        =   52
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Curriculum Vitae"
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
            Left            =   5220
            TabIndex        =   51
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5720
         MaxLength       =   80
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000040C0&
         Caption         =   "Ver datos del Postulante"
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
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4800
         MaxLength       =   80
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         MaxLength       =   80
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_contratacion_calificacion.frx":D6B1C
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txtFecha3 
         DataField       =   "cotiza_fecha"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   7200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100335617
         CurrentDate     =   41640
         MinDate         =   2
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
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
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   3000
         Visible         =   0   'False
         Width           =   9555
         Begin VB.CommandButton BtnNo 
            BackColor       =   &H00C0C000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   5400
            MaskColor       =   &H00000000&
            Picture         =   "frm_ao_contratacion_calificacion.frx":D6B36
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Cancelar"
            Top             =   240
            Width           =   765
         End
         Begin VB.CommandButton BtnOk 
            BackColor       =   &H00C0C000&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   4080
            Picture         =   "frm_ao_contratacion_calificacion.frx":D70C0
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9840
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Pre-Seleccionado"
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
         Left            =   6240
         TabIndex        =   63
         Top             =   7200
         Width           =   1635
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Observaciones"
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
         TabIndex        =   60
         Top             =   7680
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.de File"
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
         Left            =   7200
         TabIndex        =   49
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Correl.Convoca"
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
         Left            =   8280
         TabIndex        =   48
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl_convoca 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
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
         Left            =   8400
         TabIndex        =   47
         Top             =   520
         Width           =   1215
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha.de Entrevista"
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
         TabIndex        =   46
         Top             =   7200
         Width           =   1785
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Puesto al que postula"
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
         TabIndex        =   44
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label txtBenef 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "rrhh_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
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
         Left            =   7200
         TabIndex        =   30
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Trámite"
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lbl_campo_des 
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
         Index           =   0
         Left            =   1545
         TabIndex        =   28
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle1"
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
         Left            =   240
         TabIndex        =   27
         Top             =   520
         Width           =   1215
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1560
         TabIndex        =   26
         Top             =   525
         Width           =   5535
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   9480
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
      Caption         =   "Ado_clasif1"
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
   Begin MSAdodcLib.Adodc Ado_clasif2 
      Height          =   330
      Left            =   2520
      Top             =   9480
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
      Caption         =   "Ado_clasif2"
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
   Begin MSAdodcLib.Adodc Ado_clasif3 
      Height          =   330
      Left            =   4680
      Top             =   9480
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
      Caption         =   "Ado_clasif3"
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
   Begin MSAdodcLib.Adodc Ado_clasif4 
      Height          =   330
      Left            =   360
      Top             =   9720
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
      Caption         =   "Ado_clasif4"
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
   Begin MSAdodcLib.Adodc Ado_clasif5 
      Height          =   330
      Left            =   2520
      Top             =   9720
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
      Caption         =   "Ado_clasif5"
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
   Begin MSAdodcLib.Adodc Ado_clasif6 
      Height          =   330
      Left            =   4680
      Top             =   9720
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
      Caption         =   "Ado_clasif6"
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
Attribute VB_Name = "frm_ao_contratacion_calificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset

Dim nomb2 As String

Dim VAR_TIME, VAR_TIME2 As Integer

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas
nomb2 = "CONTRATADO: " & dtc_desc5.Text
If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      frm_ao_contratacion.Ado_detalle1.Recordset("ges_gestion") = Year(frm_ao_contratacion.dtpFecha1.Value) 'glGestion
      frm_ao_contratacion.Ado_detalle1.Recordset("rrhh_codigo").Value = txtBenef
      'frm_ao_contratacion.Ado_detalle1.Recordset("cotiza_codigo").Value = lbl_convoca
      frm_ao_contratacion.Ado_detalle1.Recordset("unidad_codigo") = Txt_campo1.Text
      frm_ao_contratacion.Ado_detalle1.Recordset("solicitud_codigo") = txt_codigo
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
   End If
    frm_ao_contratacion.Ado_detalle1.Recordset("puesto_codigo").Value = GlPuesto 'dtc_codigo1.Text
'    frm_ao_contratacion.Ado_detalle2.Recordset("benef_primer_apellido") = txtPat.Text
'    frm_ao_contratacion.Ado_detalle2.Recordset("benef_segundo_apellido").Value = txtMat.Text
'    frm_ao_contratacion.Ado_detalle2.Recordset("benef_nombres").Value = txtNom.Text
'    frm_ao_contratacion.Ado_detalle1.Recordset("benef_direccion_domicilio").Value = txtDireccion.Text
'    frm_ao_contratacion.Ado_detalle1.Recordset("benef_telefonos_ref").Value = txtTelefono.Text
'    nomb2 = Trim(txtPat) + " " + Trim(txtMat) + " " + Trim(txtNom)
'    Set rs_aux2 = New ADODB.Recordset
'    If rs_aux2.State = 1 Then rs_aux2.Close
'    rs_aux2.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & txtCI.Text & "'   ", db, adOpenStatic
'    If rs_aux2.RecordCount > 0 Then
'        frm_ao_contratacion.Ado_detalle2.Recordset("benef_id").Value = Ado_clasif5.Recordset!beneficiario_id
'    Else
'        frm_ao_contratacion.Ado_detalle2.Recordset("benef_id").Value = 0
'    End If
    frm_ao_contratacion.Ado_detalle1.Recordset("beneficiario_codigo").Value = dtc_codigo5.Text
    
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_sobre_a").Value = ComboA
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_sobre_b").Value = ComboB
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_sobre_c").Value = ComboC
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_sobre_todos").Value = Combo1
    
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_califica_sobre_a").Value = ComboA1
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_califica_sobre_b").Value = ComboB1
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_califica_sobre_c").Value = ComboC1
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_califica_total").Value = Combo11
    
    frm_ao_contratacion.Ado_detalle1.Recordset("fecha_apertura_propuestas") = TxtFecha1.Value
    frm_ao_contratacion.Ado_detalle1.Recordset("fecha_calificacion_propuesta") = TxtFecha2.Value
    frm_ao_contratacion.Ado_detalle1.Recordset("fecha_entrevista_proponente") = TxtFecha3.Value
    
    frm_ao_contratacion.Ado_detalle1.Recordset("apertura_selecionado").Value = Combo2
    
    frm_ao_contratacion.Ado_detalle1.Recordset("ocup_codigo").Value = dtc_codigo2.Text
    frm_ao_contratacion.Ado_detalle1.Recordset("munic_codigo").Value = IIf(dtc_codigo4.Text = "", "20101", dtc_codigo4.Text)
    frm_ao_contratacion.Ado_detalle1.Recordset("nivel_educ_codigo").Value = dtc_codigo3.Text
    If txt_obs.Text = "" Then
        frm_ao_contratacion.Ado_detalle1.Recordset("observaciones") = Trim(dtc_desc5.Text) + txt_obs.Text   'IIf(txt_obs.Text = "", dtc_desc5.Text, txt_obs.Text)
    Else
        frm_ao_contratacion.Ado_detalle1.Recordset("observaciones") = txt_obs.Text
    End If
    frm_ao_contratacion.Ado_detalle1.Recordset("usr_codigo") = glusuario 'frmLogin.txtUserName.Text
    frm_ao_contratacion.Ado_detalle1.Recordset("fecha_registro") = Date
    frm_ao_contratacion.Ado_detalle1.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    
    sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        If Combo2.Text = "SI" Then
            frm_ao_contratacion.Ado_detalle1.Recordset("estado_codigo") = "APR"
            Call GRABA_ADJUDICA
            MsgBox "El Postulante ha sido elegido, a continuación registre los datos para la Contratación ...", vbExclamation, "Atención"
        Else
            frm_ao_contratacion.Ado_detalle1.Recordset("estado_codigo") = "ANL"
            MsgBox "El postulante NO fue elegido, el proceso Finaliza para este ...", vbExclamation, "Atención"
        End If
    Else
        frm_ao_contratacion.Ado_detalle1.Recordset("estado_codigo") = "REG"
    End If
    frm_ao_contratacion.Ado_detalle1.Recordset.Update
    Para_Aceptado = "S"
   'frm_ao_solicitud_rrhh.ado_detalle2.Refresh '.Recordset.Requery
'   txtSW = "0"
   frm_ao_contratacion.ABRIR_TABLA_DET
   Unload Me
End If
End Sub

Private Sub GRABA_ADJUDICA()
    db.Execute "Insert INTO ro_rrhh_adjudica_personas (ges_gestion, rrhh_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, observaciones, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & Year(frm_ao_contratacion.dtpFecha1.Value) & "', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & parametro & "', '" & txt_codigo.Caption & "', '" & nomb2 & "', '" & GlPuesto & "', 'REG', '" & glusuario & "',  '" & Date & "')"
    '('" & glGestion & "', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & parametro & "', '" & txt_codigo.Caption & "', '" & txt_obs & "', '" & glpuesto & "', 'REG', '" & glusuario & "',  '" & Date & "')"
    
' beneficiario_monto_adjudica_bs, beneficiario_monto_adjudica_dol, tipo_moneda, beneficiario_fecha_inicio, beneficiario_fecha_fin, beneficiario_fecha_adjudica, beneficiario_fecha_contrato,
' proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, cite_tramite,                      hora_registro
End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
    Valida = True
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
  If (dtc_codigo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
'  If (dtc_codigo4.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If txtPat = "" Then
'        Valida = False
'    End If
'    If txtNom = "" Then
'        Valida = False
'    End If
End Function

Private Sub BtnNo_Click()
'    Frame2.Visible = False
'    'Frame3.Visible = False
End Sub

Private Sub BtnOk_Click()
'    txtCI.Text = dtc_codigo5.Text
'    txtPat.Text = Trim(dtc_aux1.Text)
''    txtMat.Text = Trim(dtc_aux2.Text)
''    txtNom.Text = Trim(dtc_aux3.Text)
'    txtTelefono.Text = Trim(dtc_aux4.Text)
'    txtDireccion.Text = Trim(dtc_aux5.Text)
'    Call abrir_tablas
'    Frame2.Visible = False
'    Frame4.Visible = True
'    Frame5.Visible = True
'    Frame6.Visible = True
'    txtFecha1.Value = Date
'    txtFecha2.Value = Date
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
'    dtc_aux1.BoundText = dtc_aux4.BoundText
'    dtc_aux2.BoundText = dtc_aux4.BoundText
'    dtc_aux3.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
'    dtc_aux1.BoundText = dtc_aux5.BoundText
'    dtc_aux2.BoundText = dtc_aux5.BoundText
'    dtc_aux3.BoundText = dtc_aux5.BoundText
    dtc_aux4.BoundText = dtc_aux5.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_aux1.BoundText = dtc_codigo5.BoundText
'    dtc_aux2.BoundText = dtc_codigo5.BoundText
'    dtc_aux3.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
'    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & Txt_campo1 & "' ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
'    If txtSW = "IDIR" Then
'        Option1.Visible = False
'        Option2.Visible = False
''        Frame3.Visible = True
'    Else
'        Option1.Visible = True
'        Option2.Visible = True
''        Frame3.Visible = False
'    End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
'    dtc_aux1.BoundText = dtc_desc5.BoundText
'    dtc_aux2.BoundText = dtc_desc5.BoundText
'    dtc_aux3.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Activate()
    parametro = Aux
    Call abrir_tablas
    txtCI.Text = dtc_codigo5.Text
    Frame5.Visible = True
    Frame6.Visible = True
    TxtFecha1.Value = Date
    TxtFecha2.Value = Date
    TxtFecha3.Value = Date
End Sub

Private Sub Form_Load()
'If glProceso = "CONSULTORIA" Then
'    Me.Caption = "Consultoría - Captura de datos personales"
'Else
'    Me.Caption = "Recursos Humanos - Captura de datos personales"
'End If
'Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
    'txtSW = "0"
    parametro = Aux
    'Call abrir_tablas
	Call SeguridadSet(Me)
End Sub

Private Sub abrir_tablas()
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
    'rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' ORDER BY puesto_descripcion ", db, adOpenStatic
    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_clasif2 = New ADODB.Recordset
    If rs_clasif2.State = 1 Then rs_clasif2.Close
    rs_clasif2.Open "SELECT * FROM gc_ocupacion_profesion ORDER BY ocup_descripcion ", db, adOpenStatic
    Set Ado_clasif2.Recordset = rs_clasif2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_clasif3 = New ADODB.Recordset
    If rs_clasif3.State = 1 Then rs_clasif3.Close
    rs_clasif3.Open "SELECT * FROM rc_nivel_educacional ORDER BY nivel_educ_descripcion ", db, adOpenStatic
    Set Ado_clasif3.Recordset = rs_clasif3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_clasif4 = New ADODB.Recordset
    If rs_clasif4.State = 1 Then rs_clasif4.Close
    rs_clasif4.Open "SELECT * FROM gc_municipio WHERE region_codigo = 'SI' ORDER BY munic_descripcion ", db, adOpenStatic
    Set Ado_clasif4.Recordset = rs_clasif4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_clasif5.Open "SELECT * FROM rv_beneficiario_invitacion where puesto_codigo  = '" & GlPuesto & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText


End Sub

Private Sub Option1_Click()
'    Frame4.Visible = True
'    Frame2.Visible = False
''    Frame3.Visible = False
''    txtSW = "1"
    
End Sub

Private Sub Option2_Click()
'    Frame2.Visible = True
'    Frame4.Visible = False
''    Frame3.Visible = False
''    txtSW = "2"
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    Option2.Visible = False
End Sub

Private Sub txtFecha2_LostFocus()
    VAR_TIME = DateDiff("y", TxtFecha1, TxtFecha2)
    If Val(VAR_TIME) < 0 Then
        MsgBox "La Fecha Calificacion Propuesta NO puede ser MENOR a la Fecha Apertura Propuesta, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        TxtFecha2.SetFocus
    End If
End Sub

Private Sub txtFecha3_LostFocus()
    VAR_TIME = DateDiff("y", TxtFecha2, TxtFecha3)
    If Val(VAR_TIME) < 0 Then
        MsgBox "La Fecha de Entrevista NO puede ser MENOR a la Fecha Calificacion Propuesta, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        TxtFecha3.SetFocus
    End If
End Sub

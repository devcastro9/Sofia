VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_p_gc_beneficiario_aux 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Auxiliar de Personas"
   ClientHeight    =   6495
   ClientLeft      =   420
   ClientTop       =   1830
   ClientWidth     =   9060
   Icon            =   "gw_p_gc_beneficiario_aux.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9060
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   240
      Picture         =   "gw_p_gc_beneficiario_aux.frx":0A02
      ScaleHeight     =   1140
      ScaleWidth      =   8580
      TabIndex        =   76
      Top             =   5280
      Width           =   8640
      Begin VB.ComboBox dtc_codigo11 
         Height          =   315
         ItemData        =   "gw_p_gc_beneficiario_aux.frx":6CA34
         Left            =   120
         List            =   "gw_p_gc_beneficiario_aux.frx":6CA47
         TabIndex        =   18
         Text            =   "CALLE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt_descripcion11 
         DataField       =   "calle_denominacion"
         Height          =   645
         Left            =   2160
         TabIndex        =   19
         Text            =   "-"
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H80000018&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   6585
         MaskColor       =   &H00000000&
         Picture         =   "gw_p_gc_beneficiario_aux.frx":6CA68
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancelar"
         Top             =   315
         Width           =   1005
      End
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H80000018&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   7560
         Picture         =   "gw_p_gc_beneficiario_aux.frx":6CC72
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label lbl_descripcion11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación Via de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   2280
         TabIndex        =   78
         Top             =   105
         Width           =   2895
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Vía de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   77
         Top             =   105
         Width           =   1785
      End
   End
   Begin VB.OptionButton OptFilGral1 
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
      Left            =   3720
      TabIndex        =   64
      Top             =   6960
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.PictureBox fraDatos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6315
      ScaleWidth      =   8805
      TabIndex        =   32
      Top             =   120
      Width           =   8865
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   60
         Picture         =   "gw_p_gc_beneficiario_aux.frx":6CE7C
         ScaleHeight     =   915
         ScaleWidth      =   8580
         TabIndex        =   74
         Top             =   5280
         Width           =   8640
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H80000015&
            Height          =   675
            Left            =   2640
            Picture         =   "gw_p_gc_beneficiario_aux.frx":D8EAE
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1365
         End
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H80000015&
            Height          =   675
            Left            =   4200
            MaskColor       =   &H00000000&
            Picture         =   "gw_p_gc_beneficiario_aux.frx":D9684
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   1365
         End
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   435
            Left            =   10425
            TabIndex        =   75
            Top             =   300
            Width           =   135
         End
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   80
         Text            =   "-"
         Top             =   2400
         Width           =   4320
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   720
         MaxLength       =   30
         TabIndex        =   79
         Text            =   "-"
         Top             =   2400
         Width           =   3180
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00C0C0C0&
         Height          =   1545
         Left            =   60
         TabIndex        =   33
         Top             =   -15
         Width           =   8670
         Begin VB.TextBox txt_codigo 
            BackColor       =   &H00FFFFFF&
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   4000
            MaxLength       =   15
            TabIndex        =   1
            Top             =   480
            Width           =   2205
         End
         Begin VB.TextBox Txt_descripcion 
            DataField       =   "beneficiario_denominacion"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Visible         =   0   'False
            Width           =   8430
         End
         Begin VB.TextBox txt_campo1 
            DataField       =   "beneficiario_primer_apellido"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   105
            TabIndex        =   4
            Top             =   1125
            Width           =   2910
         End
         Begin VB.TextBox txt_campo2 
            DataField       =   "beneficiario_segundo_apellido"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   3105
            TabIndex        =   5
            Top             =   1125
            Width           =   2550
         End
         Begin VB.TextBox txt_campo3 
            DataField       =   "beneficiario_nombres"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   5745
            TabIndex        =   6
            Top             =   1125
            Width           =   2790
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9F70
            DataField       =   "tipodoc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6360
            TabIndex        =   2
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "tipodoc_codigo"
            BoundColumn     =   "tipodoc_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9F89
            DataField       =   "depto_sigla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7440
            TabIndex        =   3
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_sigla"
            BoundColumn     =   "depto_sigla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9FA2
            DataField       =   "tipodoc_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6600
            TabIndex        =   59
            Top             =   720
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "tipodoc_descripcion"
            BoundColumn     =   "tipodoc_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9FBB
            DataField       =   "depto_sigla"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7680
            TabIndex        =   60
            Top             =   720
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483628
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_sigla"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9FD4
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "tipoben_descripcion"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":D9FED
            DataField       =   "tipoben_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3120
            TabIndex        =   68
            Top             =   120
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483637
            ListField       =   "tipoben_codigo"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Persona"
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
            TabIndex        =   67
            Top             =   225
            Width           =   1515
         End
         Begin VB.Label lbl_campo6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Expedido en"
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
            Left            =   7440
            TabIndex        =   63
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc."
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
            Left            =   6360
            TabIndex        =   62
            Top             =   225
            Width           =   885
         End
         Begin VB.Label lbl_titulo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Documento Identidad"
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
            Left            =   3945
            TabIndex        =   37
            Top             =   225
            Width           =   2280
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
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
            Left            =   5745
            TabIndex        =   36
            Top             =   855
            Width           =   840
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Primer Apellido"
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
            Left            =   105
            TabIndex        =   35
            Top             =   855
            Width           =   1380
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Segundo Apellido"
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
            Left            =   3105
            TabIndex        =   34
            Top             =   855
            Width           =   1620
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lugar donde Radica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2535
         Left            =   60
         TabIndex        =   51
         Top             =   2760
         Width           =   8670
         Begin VB.TextBox txt_campo10 
            DataField       =   "beneficiario_edif_nro"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7120
            TabIndex        =   14
            Top             =   2040
            Width           =   1380
         End
         Begin VB.CommandButton BtnAux1 
            BackColor       =   &H80000015&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5520
            Picture         =   "gw_p_gc_beneficiario_aux.frx":DA006
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Nueva Via (Calle, Av, etc)"
            Top             =   1800
            Width           =   1020
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAADE
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6360
            TabIndex        =   52
            Top             =   195
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAAF7
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   53
            Top             =   1035
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_codigo"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB10
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB29
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_codigo"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB42
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4320
            TabIndex        =   10
            Top             =   600
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB5B
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   55
            Top             =   360
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_codigo"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB74
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAB8D
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5280
            TabIndex        =   27
            Top             =   120
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DABA6
            DataField       =   "zona_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4320
            TabIndex        =   12
            Top             =   1320
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "zona_denominacion"
            BoundColumn     =   "zona_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc9 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DABBF
            DataField       =   "calle_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "calle_denominacion"
            BoundColumn     =   "calle_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DABD8
            DataField       =   "zona_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   71
            Top             =   960
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "zona_codigo"
            BoundColumn     =   "zona_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo9 
            Bindings        =   "gw_p_gc_beneficiario_aux.frx":DABF1
            DataField       =   "calle_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3960
            TabIndex        =   72
            Top             =   1800
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "calle_codigo"
            BoundColumn     =   "calle_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label LlbCargo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Vivienda"
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
            TabIndex        =   73
            Top             =   1800
            Width           =   1185
         End
         Begin VB.Label lbl_calle 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Via de Acceso (Calle, Av, etc.)"
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
            TabIndex        =   70
            Top             =   1800
            Width           =   2685
         End
         Begin VB.Label lbl_zona 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Zona / Barrio"
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
            Left            =   4320
            TabIndex        =   69
            Top             =   1060
            Width           =   1155
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Municipio"
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
            Index           =   3
            Left            =   120
            TabIndex        =   58
            Top             =   1060
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            TabIndex        =   57
            Top             =   340
            Width           =   1290
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia"
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
            Left            =   4320
            TabIndex        =   56
            Top             =   340
            Width           =   1080
         End
      End
      Begin VB.TextBox txt_campo7 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_telefono_Cel"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   5880
         MaxLength       =   30
         TabIndex        =   25
         Text            =   "-"
         Top             =   855
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txt_campo5 
         DataField       =   "beneficiario_telefono_Cel"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   720
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "-"
         Top             =   1800
         Width           =   3180
      End
      Begin VB.TextBox txt_campo6 
         DataField       =   "beneficiario_telefono_Of"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   24
         Text            =   "-"
         Top             =   855
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox txt_campo11 
         DataField       =   "beneficiario_edif_piso_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6285
         TabIndex        =   29
         Top             =   5535
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txt_campo9 
         DataField       =   "beneficiario_email_of"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4260
         MaxLength       =   50
         TabIndex        =   26
         Text            =   "-"
         Top             =   1080
         Visible         =   0   'False
         Width           =   4150
      End
      Begin VB.TextBox txt_campo12 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_edif_depto_nro"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7395
         TabIndex        =   30
         Top             =   5535
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txt_campo8 
         DataField       =   "beneficiario_email"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   8
         Text            =   "-"
         Top             =   1800
         Width           =   4320
      End
      Begin VB.TextBox txt_campo4 
         BackColor       =   &H00FFFFFF&
         DataField       =   "beneficiario_nit"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   2220
      End
      Begin MSComCtl2.DTPicker DTP_Fecha1 
         DataField       =   "beneficiario_fecha_nacimiento"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6300
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   116391937
         CurrentDate     =   40179
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAC0A
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2700
         TabIndex        =   47
         Top             =   5400
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAC24
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   5535
         Visible         =   0   'False
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux4 
         Bindings        =   "gw_p_gc_beneficiario_aux.frx":DAC3E
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "pais_cod_telefonico"
         BoundColumn     =   "pais_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Interno"
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
         Left            =   4200
         TabIndex        =   82
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Número Corto"
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
         Left            =   720
         TabIndex        =   81
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label LblInicial 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "REG"
         DataField       =   "beneficiario_iniciales"
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
         ForeColor       =   &H00FFFF80&
         Height          =   315
         Left            =   7320
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Iniciales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   14
         Left            =   6480
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Depto."
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
         Left            =   7380
         TabIndex        =   50
         Top             =   5280
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Piso"
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
         Left            =   6315
         TabIndex        =   49
         Top             =   5280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl_campo17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
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
         TabIndex        =   48
         Top             =   5280
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   9
         Left            =   6960
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "REG"
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
         ForeColor       =   &H00FFFF80&
         Height          =   315
         Left            =   7800
         TabIndex        =   45
         Top             =   180
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Teléfono Celular Personal"
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
         Left            =   5880
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos de Referencia"
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
         Left            =   720
         TabIndex        =   43
         Top             =   1560
         Width           =   2235
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Teléfono(s) Oficina:"
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
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Nacimiento"
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
         Left            =   4320
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico Personal"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   40
         Top             =   1560
         Width           =   2880
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Correo Electrónico Institucional"
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
         Left            =   4275
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Número de NIT"
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
         Left            =   120
         TabIndex        =   38
         Top             =   1215
         Visible         =   0   'False
         Width           =   1380
      End
   End
   Begin Crystal.CrystalReport CR01 
      Left            =   2400
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   6480
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   6240
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   8640
      Top             =   6600
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
      Left            =   2160
      Top             =   6240
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   0
      Top             =   6600
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4320
      Top             =   6240
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
      Left            =   2160
      Top             =   6600
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6480
      Top             =   6240
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
      Left            =   4320
      Top             =   6600
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   8640
      Top             =   6240
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   0
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   2880
      Top             =   6960
      Width           =   3465
      _ExtentX        =   6112
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
Attribute VB_Name = "gw_p_gc_beneficiario_aux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Mantenimiento de Beneficiarios
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset

Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim var_cod, VAR_PAIS As String
Dim VAR_VAL As String
Dim VAR_SW, VAR_AUX As String
Dim NombreCarpeta, e As String
Dim SQL_FOR As String
Dim RUTA1 As String
Dim VAR_PWD As String
Dim CodBenef As String
Dim sino As String
Dim queryinicial As String

Dim VAR_COD2 As Double

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  If Ado_datos.Recordset.EOF Or Ado_datos.Recordset.BOF Then
'      BtnModificar.Enabled = False
'     ' BtnEliminar.Enabled = False
'      'TxtTipo.Text = Empty
'      txtCodigo.Text = Empty
'      Text1.Text = Empty
'      Text2.Text = Empty
'      Text3.Text = Empty
'      txtDenominacion.Text = Empty
'      Exit Sub
'  End If
  If Ado_datos.Recordset.RecordCount > 0 Then
'    Select Case Ado_datos.Recordset.EditMode
'      Case adEditInProgress
'        Frame2.Enabled = False            'Verif. Nombre Proveedor JQA NOV-2009
'
'      Case adEditNone
'      Case adEditDelete
'      Case adEditAdd
'        Frame2.Enabled = True            'Verif. Nombre Proveedor JQA NOV-2009
'    End Select

    If VAR_SW = "ADD" Then
      Txt_descripcion.Visible = False
      txt_campo1.Visible = True
      Txt_campo2.Visible = True
      Txt_campo3.Visible = True
    Else
      Txt_descripcion.Visible = True
      txt_campo1.Visible = False
      Txt_campo2.Visible = False
      Txt_campo3.Visible = False
    End If
    'Ado_datos.Caption = Ado_datos.Recordset!beneficiario_codigo + " - " + CStr(Ado_datos.Recordset!calle_codigo)
    'Ado_datos.Caption = CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & CStr(Ado_datos.Recordset.RecordCount)
    '  <-- Inicio                   Viviendas - Edificaciones                   Fin -->
  End If
End Sub
   
Private Sub BtnAux1_Click()
    'Validacion 1
    If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
    End If
    FraGrabarCancelar.Visible = False
    Fra_aux1.Visible = True
    fraDatos.Enabled = False
'para enlazar Formulario: frm_gc_calles
'    frm_gc_calles.lbl_titulo = frmMain.Mnu_ViasAcceso.Caption
'    frm_gc_calles.FraNavega = frmMain.Mnu_ViasAcceso.Caption
'    frm_gc_calles.lbl_titulo2 = frmMain.Mnu_ViasAcceso.Caption
'    frm_gc_calles.Show
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
   If glPersNew <> "FICHA" Then
    If VAR_SW = "ADD" Then
       Set rs_aux1 = New ADODB.Recordset
       SQL_FOR = "select * from gc_beneficiario where beneficiario_codigo = '" & txt_codigo.Text & "'  "
       rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic        ', adCmdText
       If rs_aux1.RecordCount > 0 Then
'                SW = True
                MsgBox " CODIGO DUPLICADO"
                txt_codigo.SetFocus
                Exit Sub
       End If
   End If
       Ado_datos.Recordset!beneficiario_codigo = txt_codigo.Text
       glBenef = txt_codigo.Text
       If glPersNew = "NEWC" Then
            mw_solicitud.txt_ci.Text = txt_codigo.Text
        End If
        If glPersNew = "NEWF" Then
            tw_identificacion_cliente.txt_ci.Text = txt_codigo.Text
        End If
       Ado_datos.Recordset!estado_codigo = "REG"
        'ado_datos.recordset!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
        'ado_datos.recordset!archivo_foto_cargado = "N"
        'ado_datos.recordset!ges_gestion = Year(Date)
        'ado_datos.recordset!correl_da = 0
        'db.Execute "Update gc_municipio Set correl_edif = CAST('" & dtc_aux2.Text & "' AS INT) + 1 Where munic_codigo= '" & dtc_codigo2.Text & "' "
     End If
     If Txt_campo2.Text = "" Then Txt_campo2.Text = "-"
     LblInicial.Caption = Trim(Left(txt_campo1.Text, 1)) + Trim(Left(Txt_campo2.Text, 1)) + Trim(Left(Txt_campo3.Text, 1))
     var_cod = IIf(txt_campo1.Text = "", "", txt_campo1.Text + " ") + IIf(Txt_campo2.Text = "", "", Txt_campo2.Text + " ") + IIf(Txt_campo3.Text = "", "", Txt_campo3.Text)
     Ado_datos.Recordset!depto_sigla = dtc_codigo3.Text
     Ado_datos.Recordset!beneficiario_iniciales = LblInicial.Caption
     Ado_datos.Recordset!tipodoc_codigo = dtc_codigo2.Text
     Ado_datos.Recordset!tipoben_codigo = dtc_codigo1.Text
     Ado_datos.Recordset!beneficiario_nit = IIf(Txt_campo4.Text = "", txt_codigo, Txt_campo4.Text)
     Ado_datos.Recordset!beneficiario_primer_apellido = Trim(txt_campo1.Text)
     Ado_datos.Recordset!beneficiario_segundo_apellido = Trim(Txt_campo2.Text)
     Ado_datos.Recordset!beneficiario_nombres = Trim(Txt_campo3.Text)
     Ado_datos.Recordset!beneficiario_denominacion = var_cod
     Ado_datos.Recordset!beneficiario_fecha_nacimiento = DTP_Fecha1.Value  'IIF(ISNULL(DTP_Fecha1.Value),DATE,DTP_Fecha1.Value)
     'Ado_datos.Recordset!beneficiario_telefono_fijo = IIf(txt_campo5.Text = "", "0", txt_campo5.Text)
     Ado_datos.Recordset!beneficiario_telefono_Of = IIf(Txt_campo6.Text = "", "0", Txt_campo6.Text)
     Ado_datos.Recordset!beneficiario_telefono_Cel = IIf(Txt_campo5.Text = "", "0", Txt_campo5.Text)
     Ado_datos.Recordset!beneficiario_email = IIf(Txt_campo8.Text = "", "-", Txt_campo8.Text)
     Ado_datos.Recordset!beneficiario_email_of = IIf(Txt_campo9.Text = "", "-", Txt_campo9.Text)
     Ado_datos.Recordset!beneficiario_domicilio_legal = "Z. " + dtc_desc8.Text + " C. " + dtc_desc9.Text + " # " + Txt_campo10.Text
     Ado_datos.Recordset!pais_codigo = IIf(dtc_codigo4.Text = "", "BOL", dtc_codigo4.Text)
     Ado_datos.Recordset!depto_codigo = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)
     Ado_datos.Recordset!prov_codigo = IIf(dtc_codigo6.Text = "", "0", dtc_codigo6.Text)
     Ado_datos.Recordset!munic_codigo = IIf(dtc_codigo7.Text = "", "0", dtc_codigo7.Text)
     Ado_datos.Recordset!zona_codigo = IIf(dtc_codigo8.Text = "", "0", dtc_codigo8.Text)
     Ado_datos.Recordset!calle_codigo = IIf(dtc_codigo9.Text = "", "0", dtc_codigo9.Text)
     Ado_datos.Recordset!edif_codigo = IIf(dtc_codigo10.Text = "", "10101-0", dtc_codigo10.Text)
     
     Ado_datos.Recordset!beneficiario_edif_nro = IIf(Txt_campo10.Text = "", "0", Txt_campo10.Text)
     Ado_datos.Recordset!beneficiario_edif_piso_nro = IIf(Txt_campo11.Text = "", "0", Txt_campo11.Text)
     Ado_datos.Recordset!beneficiario_edif_depto_nro = IIf(Txt_campo12.Text = "", "0", Txt_campo12.Text)
     
'     If ado_datos.recordset!ARCHIVO_Foto = ".JPG" Or ado_datos.recordset!ARCHIVO_Foto = "" Then
'        ado_datos.recordset!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
'     End If
     
     Ado_datos.Recordset!fecha_registro = Date
     Ado_datos.Recordset!usr_codigo = glusuario
        
        
'    If glPersNew <> "FICHA" Then
'        RUTA1 = "CLIENTES\" + Trim(Ado_datos.Recordset("beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
'        MsgBox RUTA1
'        MkDir RUTA1
'        MkDir RUTA1 + "\CONTRATOS"
'        MkDir RUTA1 + "\RESPALDOS"
'        MkDir RUTA1 + "\HOJA_VIDA"
'        MkDir RUTA1 + "\OTROS"
'    End If
       'glPersNew = "NEWC"
        Ado_datos.Recordset!estado_codigo = "APR"
         'rs_datos!fecha_registro = Date
         'rs_datos!usr_codigo = glusuario
         'rs_datos.Update
     Ado_datos.Recordset.UpdateBatch adAffectAll
     If glPersNew <> "FICHA" Then
        RUTA1 = "CLIENTES\" + Trim(Ado_datos.Recordset("beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
        MsgBox RUTA1
        MkDir RUTA1
        MkDir RUTA1 + "\CONTRATOS"
        MkDir RUTA1 + "\RESPALDOS"
        MkDir RUTA1 + "\HOJA_VIDA"
        MkDir RUTA1 + "\OTROS"
    End If
     If glPersNew = "FICHA" Then
        db.Execute "update ro_personal_contratado set codigo_interno = '" & Text2.Text & "', codigo_corto = '" & Text1.Text & "' where beneficiario_codigo = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "'"
        rw_ficha_rrhh.Ado_datos.Recordset.UpdateBatch adAffectAll
     
        If rw_ficha_rrhh.OptFilGral1.Value = True Then
            rw_ficha_rrhh.OptFilGral2.Value = True
            rw_ficha_rrhh.OptFilGral1.Value = True
        End If
     
        If rw_ficha_rrhh.OptFilGral2.Value = True Then
            rw_ficha_rrhh.OptFilGral1.Value = True
            rw_ficha_rrhh.OptFilGral2.Value = True
        End If
     
         rw_ficha_rrhh.dtc_buscar_ci = txt_codigo.Text
         rw_ficha_rrhh.dtc_buscar_desc.BoundText = rw_ficha_rrhh.dtc_buscar_ci.BoundText
     End If
'     'Call ABRIR_TABLA
'     Select Case Ado_datos.Recordset!tipoben_codigo
'        Case Is < 20
''          Call OptFilGral1_Click        'TODOS
'
''        Case Is < 2
''          Call OptFilGral2_Click        'PERSONAL CGI
''
''        Case 3 Or 5 Or 0
''          Call OptFilGral3_Click        'PROVEEDORES
''
''        Case 2 Or 4 Or 0
''          Call OptFilGral4_Click        'CLIENTES
'
'     End Select
'     Ado_datos.Recordset.MoveLast
''     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
'     fraDatos.Enabled = False
'     dg_datos.Enabled = True
     Txt_descripcion.Visible = True
     txt_campo1.Visible = False
     Txt_campo2.Visible = False
     Txt_campo3.Visible = False
     txt_codigo.Enabled = True
     If glPersNew = "NEWC" Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & glBenef & "' ", db, adOpenStatic
        Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
        mw_solicitud.txt_ci = txt_codigo
        mw_solicitud.txt_nombre.Visible = True
        mw_solicitud.txt_nombre.Text = rs_aux1!beneficiario_denominacion
        'Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
        mw_solicitud.dtc_codigo4.Text = txt_codigo
        mw_solicitud.dtc_desc4.BoundText = mw_solicitud.dtc_codigo4.BoundText
        mw_solicitud.txt_obs = txt_codigo.Text + " - " + rs_aux1!beneficiario_denominacion + " - Telef. " + IIf(IsNull(rs_aux1!beneficiario_telefono_fijo), "0", rs_aux1!beneficiario_telefono_Cel)
     End If
     If glPersNew = "NEWF" Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & glBenef & "' ", db, adOpenStatic
        Set tw_identificacion_cliente.Ado_datos4.Recordset = rs_aux1
        tw_identificacion_cliente.txt_ci = txt_codigo
        tw_identificacion_cliente.txt_nombre.Visible = True
        tw_identificacion_cliente.txt_nombre.Text = rs_aux1!beneficiario_denominacion
        'Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
        tw_identificacion_cliente.dtc_codigo4.Text = txt_codigo
        tw_identificacion_cliente.dtc_desc4.BoundText = tw_identificacion_cliente.dtc_codigo4.BoundText
        tw_identificacion_cliente.txt_obs3 = txt_codigo.Text + " - " + rs_aux1!beneficiario_denominacion + " - Telef. " + IIf(IsNull(rs_aux1!beneficiario_telefono_fijo), "0", rs_aux1!beneficiario_telefono_Cel)
     End If
  
  End If
  'WWWWWWWWWWWWWWWW APROBAR
   
  'WWWWWWWWWWWWWWWW
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
    
End Sub

Private Sub valida_campos()
  If (txt_codigo.Text = "") Then
    MsgBox "Debe registrar el " + lbl_titulo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If txt_campo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If txt_campo2.Text = "" Then
'    MsgBox "Debe registrar la " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_campo3.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3.Text = "" Then
    MsgBox "Debe registrar la " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub
 
Private Sub graba_persona()
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ro_personal_contratado where beneficiario_codigo = '" & txt_codigo.Text & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount = 0 Then
        rs_aux1.AddNew
        rs_aux1!beneficiario_codigo = txt_codigo.Text
        'rs_aux1!idfuncionario = CORREL
    'Else
        'MsgBox " YA EXISTE EL CODIGO ..."
    End If
        rs_aux1!ARCHIVO_Foto = Trim(LblInicial.Caption) + Ado_datos.Recordset("beneficiario_codigo") + ".JPG"
        rs_aux1!archivo_foto_cargado = "N"
        rs_aux1!archivo_hojavida = Trim(LblInicial.Caption) + Ado_datos.Recordset("beneficiario_codigo") + "_HV.PDF"
        rs_aux1!archivo_hojavida_cargado = "N"
        rs_aux1!archivo_respaldo = Trim(LblInicial.Caption) + Ado_datos.Recordset("beneficiario_codigo") + "_DOC.PDF"
        rs_aux1!archivo_respaldo_cargado = "N"
        rs_aux1!archivo_vacaciones = Trim(LblInicial.Caption) + Ado_datos.Recordset("beneficiario_codigo") + "_VAC.PDF"
        rs_aux1!archivo_vacaciones_cargado = "N"
        rs_aux1!archivo_otros = Trim(LblInicial.Caption) + Ado_datos.Recordset("beneficiario_codigo") + "_OTR.PDF"
        rs_aux1!archivo_otros_cargado = "N"
        rs_aux1!usr_codigo = glusuario 'frmLogin.txtUserName.Text
        rs_aux1!fecha_registro = Date
        'rs_aux1!hora_registro = Format(Time, "hh:mm:ss")
        rs_aux1!estado_codigo = "REG"
        rs_aux1.Update
End Sub

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    If Ado_datos.Recordset.RecordCount > 0 Then Ado_datos.Recordset.MoveLast
    Ado_datos.Recordset.AddNew
    'lblStatus.Caption = "Agregar registro"
'    fraDatos.Enabled = True
'    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
'    dg_datos.Enabled = False
    VAR_SW = "ADD"
    txt_codigo.Enabled = True
    Txt_descripcion.Visible = False
    txt_campo1.Visible = True
    Txt_campo2.Visible = True
    Txt_campo3.Visible = True
    txt_campo1.SetFocus
'    BtnVer.Visible = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         If (Ado_datos.Recordset("tipoben_codigo") = 1) Then
            VAR_AUX = Left(Ado_datos.Recordset("beneficiario_nombres"), 1) + Ado_datos.Recordset("beneficiario_primer_apellido")
            VAR_PWD = Encriptar(Trim(Ado_datos.Recordset("beneficiario_codigo")))
'            db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, IdNivelAcceso, estado_codigo, fecha_registro, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones)" & _
'            "values ('" & Left(Ado_datos.Recordset("beneficiario_nombres"), 1) & "' + '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "', '" & Ado_datos.Recordset("beneficiario_codigo") & "','" & Trim(Ado_datos.Recordset("beneficiario_nombres")) & "', '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "','" & Trim(Ado_datos.Recordset("beneficiario_segundo_apellido")) & "','" & Ado_datos.Recordset("beneficiario_codigo") & "', '1', 'REG', '" & Date & "', '0', '0', '0', '0', '0') "
            
            db.Execute "insert into gc_usuarios(usr_codigo, beneficiario_codigo, usr_nombres, usr_primer_apellido, usr_segundo_apellido, usr_clave, dgral_codigo, da_codigo, unidad_codigo, ocup_codigo, usr_observaciones, idnivelacceso, estado_codigo, fecha_registro)" & _
            "values ('" & VAR_AUX & "', '" & Ado_datos.Recordset("beneficiario_codigo") & "','" & Trim(Ado_datos.Recordset("beneficiario_nombres")) & "', '" & Ado_datos.Recordset("beneficiario_primer_apellido") & "','" & Trim(Ado_datos.Recordset("beneficiario_segundo_apellido")) & "','" & VAR_PWD & "', '1', '0', '0', '0', '-', '1', 'REG', '" & Date & "') "

            RUTA1 = "PERSONAL" + "\" + Trim(Ado_datos.Recordset("beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
            MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\FINIQUITO"
            MkDir RUTA1 + "\MEMOS"
            MkDir RUTA1 + "\RESPALDOS"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
            MkDir RUTA1 + "\EVALUACIONES"
            MkDir RUTA1 + "\LICENCIAS"
            MkDir RUTA1 + "\VACACIONES"
            Call graba_persona
         End If
         If (Ado_datos.Recordset("tipoben_codigo") = 21) Or (Ado_datos.Recordset("tipoben_codigo") = 2) Then
            RUTA1 = "CLIENTES\" + Trim(Ado_datos.Recordset("beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
            MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\RESPALDOS"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
         End If
         If (Ado_datos.Recordset("tipoben_codigo") = 22) Or (Ado_datos.Recordset("tipoben_codigo") = 3) Then
            RUTA1 = "PROVEEDORES\" + Trim(Ado_datos.Recordset("beneficiario_iniciales")) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo"))
            MsgBox RUTA1
            MkDir RUTA1
            MkDir RUTA1 + "\CONTRATOS"
            MkDir RUTA1 + "\RESPALDOS"
            MkDir RUTA1 + "\HOJA_VIDA"
            MkDir RUTA1 + "\OTROS"
         End If
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
        Call ABRIR_TABLAS_AUX
        Select Case Ado_datos.Recordset!tipoben_codigo
          Case Is < 20
            Call OptFilGral1_Click        'TODOS
        End Select
        
        'Call ABRIR_TABLA
        rs_datos.MoveFirst
'        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
'        fraDatos.Enabled = False
'        dg_datos.Enabled = True
        Txt_descripcion.Visible = True
        txt_campo1.Visible = False
        Txt_campo2.Visible = False
        Txt_campo3.Visible = False
        txt_codigo.Enabled = True
    End If
    
      Unload Me
End Sub


Private Sub CmdCancelaDet_Click()
    fraDatos.Enabled = True
    Fra_aux1.Visible = False
    FraGrabarCancelar.Visible = True
End Sub

Private Sub CmdGrabaDet_Click()
  'Validacion
  If Txt_descripcion11.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo8 = "" Or dtc_codigo8 = "0" Then
    MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'INI Graba Calle
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select max(calle_codigo) as Codigo from gc_calles where zona_codigo = " & dtc_codigo8.Text & "    ", db, adOpenStatic
    'If rs_aux2.RecordCount > 0 Then
    If rs_aux2!Codigo > 0 Then
        VAR_COD2 = Round(CDbl(rs_aux2!Codigo) + 1, 0)
    Else
        VAR_COD2 = (Val(dtc_codigo8.Text) * 100) + 1
    End If
    db.Execute "insert into gc_calles(zona_codigo, calle_codigo, calle_denominacion, calle_tipo, correl, estado_codigo, fecha_registro, usr_codigo)" & _
    "values ('" & dtc_codigo8.Text & "', " & VAR_COD2 & ", '" & Txt_descripcion11.Text & "', '" & dtc_codigo11.Text & "', '0', 'APR', '" & Date & "', '" & glusuario & "') "
    
   'FIN Graba Calle
    'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
    db.Execute "Update gc_zonas Set correl = " & VAR_COD2 & " Where zona_codigo= '" & dtc_codigo8.Text & "' "
    'gc_calles
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
    
    dtc_codigo9.Text = VAR_COD2
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    fraDatos.Enabled = True
    Fra_aux1.Visible = False
    FraGrabarCancelar.Visible = True
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_aux4.BoundText
    dtc_codigo4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
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

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    Call pnivel5(dtc_codigo7.BoundText)
    dtc_desc8.Enabled = True
    Call pnivel7(dtc_codigo7.BoundText)
    dtc_desc10.Enabled = True
End Sub
   
Private Sub pnivel5(codigo7 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "' order by zona_denominacion"
   Set dtc_codigo8.RowSource = Nothing
   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo8.ReFill
   dtc_codigo8.BoundText = Empty
   
   Set dtc_desc8.RowSource = Nothing
   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc8.ReFill
   dtc_desc8.BoundText = Empty
End Sub

Private Sub pnivel7(codigo9 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_edificaciones where munic_codigo = '" & codigo9 & "' order by edif_descripcion"
   Set dtc_codigo10.RowSource = Nothing
   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Call pnivel6(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
End Sub

Private Sub pnivel6(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo8 & "' order by calle_denominacion"
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    If glPersNew = "FICHA" Then
        Text2.Enabled = True
        Text1.Enabled = True
        Call ABRIR_TABLA
        txt_codigo.Enabled = False
    Else
        Text2.Enabled = False
        Text1.Enabled = False
        Ado_datos.Recordset.AddNew
        txt_codigo.Enabled = True
    End If
    'txt_codigo.Enabled = True
'    mbDataChanged = False
'    fraDatos.Enabled = False
'    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    Fra_aux1.Visible = False
    'WWWWWWWWWWWWWWWWWWWWW
'      On Error GoTo AddErr
    'If Ado_datos.Recordset.RecordCount > 0 Then Ado_datos.Recordset.MoveLast
    VAR_SW = "ADD"
    VAR_PAIS = "BOL"
    
    'lblStatus.Caption = "Agregar registro"
    fraDatos.Enabled = True
'    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    'Fra_ABM.Enabled = True
    

    Txt_descripcion.Visible = False
    txt_campo1.Visible = True
    Txt_campo2.Visible = True
    Txt_campo3.Visible = True
'    txt_campo1.SetFocus
'    BtnVer.Visible = False
  Exit Sub
AddErr:
  MsgBox Err.Description

    'WWWWWWWWWWWWWWWWWWWWW
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
   If glPersNew = "FICHA" Then
        Set rs_datos = New ADODB.Recordset
        If rs_datos.State = 1 Then rs_datos.Close
        queryinicial = "select * from gc_beneficiario WHERE beneficiario_codigo = '" & rw_ficha_rrhh.Ado_datos.Recordset!beneficiario_codigo & "'"
        'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
        rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
        Text1.Text = IIf(IsNull(rw_ficha_rrhh.Ado_datos.Recordset!codigo_corto), 0, rw_ficha_rrhh.Ado_datos.Recordset!codigo_corto)
        Text2.Text = IIf(IsNull(rw_ficha_rrhh.Ado_datos.Recordset!codigo_interno), 0, rw_ficha_rrhh.Ado_datos.Recordset!codigo_interno)
        'rs_datos.Sort = "beneficiario_denominacion"
        If rs_datos.RecordCount > 0 Then
        
        End If
        
        Set Ado_datos.Recordset = rs_datos
       
        'Set dg_datos.DataSource = Ado_datos.Recordset
        Ado_datos.Recordset.MoveFirst
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If glPersNew = "NEWC" Then
'        Set rs_aux1 = New ADODB.Recordset
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & glBenef & "' ", db, adOpenStatic
'        Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
'        mw_solicitud.txt_ci = txt_codigo
'        mw_solicitud.txt_nombre.Visible = True
'        mw_solicitud.txt_nombre.Text = rs_aux1!beneficiario_denominacion
'        'Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
'        mw_solicitud.dtc_codigo4.Text = txt_codigo
'        mw_solicitud.dtc_desc4.BoundText = mw_solicitud.dtc_codigo4.BoundText
'        mw_solicitud.txt_obs = txt_codigo.Text + " - " + rs_aux1!beneficiario_denominacion + " - Telef. " + IIf(IsNull(rs_aux1!beneficiario_telefono_fijo), "0", rs_aux1!beneficiario_telefono_Cel)
'     End If
'     If glPersNew = "NEWF" Then
'        Set rs_aux1 = New ADODB.Recordset
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & glBenef & "' ", db, adOpenStatic
'        Set tw_identificacion_cliente.Ado_datos4.Recordset = rs_aux1
'        tw_identificacion_cliente.txt_ci = txt_codigo
'        tw_identificacion_cliente.txt_nombre.Visible = True
'        tw_identificacion_cliente.txt_nombre.Text = rs_aux1!beneficiario_denominacion
'        'Set mw_solicitud.Ado_datos4.Recordset = rs_aux1
'        tw_identificacion_cliente.dtc_codigo4.Text = txt_codigo
'        tw_identificacion_cliente.dtc_desc4.BoundText = tw_identificacion_cliente.dtc_codigo4.BoundText
'        tw_identificacion_cliente.txt_obs = txt_codigo.Text + " - " + rs_aux1!beneficiario_denominacion + " - Telef. " + IIf(IsNull(rs_aux1!beneficiario_telefono_fijo), "0", rs_aux1!beneficiario_telefono_Cel)
'     End If
'  If glPersNew = "P" Then
'     FrmVentas.DtcNIT = Ado_datos.Recordset("codigo_beneficiario")
'     FrmVentas.DtcdesNIT = Ado_datos.Recordset("denominacion_Beneficiario")
'  End If
'
'  glPersNew = "N"
   
   If (rs_datos.State = adStateClosed) Then rs_datos.Close
   'Set rs_datos = Nothing
End Sub

Private Sub ABRIR_TABLAS_AUX()
  'carga    fc_tipo_beneficiario
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    If glPersNew = "FICHA" Then
        rs_datos1.Open "SELECT * FROM gc_tipo_beneficiario ORDER BY tipoben_descripcion ", db, adOpenStatic
    Else
        rs_datos1.Open "SELECT * FROM gc_tipo_beneficiario WHERE tipoben_codigo = 2 or tipoben_codigo = 4 or tipoben_codigo = 6 ORDER BY tipoben_descripcion ", db, adOpenStatic
    End If
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
      'gc_tipo_documento_id     'Tipo Doc. de Id.
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from gc_tipo_documento_id", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'gc_Departamento    'Expedido en...
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_departamento order by depto_sigla", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    'gc_pais
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_pais where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'gc_Departamento  '<>
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from gc_departamento order by depto_descripcion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    'gc_provincia
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_provincia ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'gc_municipio
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_municipio ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'gc_zonas
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_zonas ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'gc_calles
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_calles ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    'gc_edificaciones
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    'gc_calle_tipo
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "Select * from gc_calle_tipo order by calle_tipo_nombre", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub OptFilGral1_Click()
   'TODOS
    Set rs_datos = New ADODB.Recordset
   If rs_datos.State = 1 Then rs_datos.Close
   queryinicial = "select * from gc_beneficiario WHERE  tipoben_codigo < 20 "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_datos.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rs_datos
End Sub

Private Sub txt_campo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo1_LostFocus()
    Txt_descripcion.Text = txt_campo1.Text + " " + Txt_campo2.Text + " " + Txt_campo3.Text
End Sub

Private Sub txt_campo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_campo3_LostFocus()
    Txt_descripcion.Text = txt_campo1.Text + " " + Txt_campo2.Text + " " + Txt_campo3.Text
End Sub

Private Function ExisteBenef(CodBenef As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo_resp = '" & CodBenef & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteBenef = rs!Cuantos > 0
End Function

Private Function ExisteBenef2(CodBenef As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE beneficiario_codigo = '" & CodBenef & "'"
    rs2.Open GlSqlAux, db, adOpenStatic
    ExisteBenef2 = rs2!Cuantos > 0
End Function


Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_aux2.BoundText = dtc_codigo2.BoundText
'    dtc_campo2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    
    Call pnivel2(VAR_PAIS)
    dtc_desc5.Enabled = True
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
    Call pnivel2(dtc_codigo4.BoundText)
    dtc_desc5.Enabled = True
End Sub
   
Private Sub pnivel2(codigo4 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo4 & "'"
   Set dtc_codigo5.RowSource = Nothing
   Set dtc_codigo5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo5.ReFill
   dtc_codigo5.BoundText = Empty
   
   Set dtc_desc5.RowSource = Nothing
   Set dtc_desc5.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc5.ReFill
   dtc_desc5.BoundText = Empty

End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    Call pnivel3(dtc_codigo5.BoundText)
    dtc_desc6.Enabled = True
End Sub
   
Private Sub pnivel3(codigo5 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo5 & "'"
   Set dtc_codigo6.RowSource = Nothing
   Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo6.ReFill
   dtc_codigo6.BoundText = Empty
   
   Set dtc_desc6.RowSource = Nothing
   Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc6.ReFill
   dtc_desc6.BoundText = Empty
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
    Call pnivel4(dtc_codigo6.BoundText)
    dtc_desc7.Enabled = True
End Sub
   
Private Sub pnivel4(codigo6 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo6 & "'"
   Set dtc_codigo7.RowSource = Nothing
   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo7.ReFill
   dtc_codigo7.BoundText = Empty
   
   Set dtc_desc7.RowSource = Nothing
   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc7.ReFill
   dtc_desc7.BoundText = Empty
End Sub


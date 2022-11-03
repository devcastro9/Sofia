VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gw_p_gc_edificaciones_aux 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Proyectos de Edificación"
   ClientHeight    =   7065
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "gw_p_gc_edificaciones_aux.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   120
      Picture         =   "gw_p_gc_edificaciones_aux.frx":0A02
      ScaleHeight     =   1200
      ScaleWidth      =   8295
      TabIndex        =   55
      Top             =   3720
      Width           =   8350
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H80000018&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   7320
         Picture         =   "gw_p_gc_edificaciones_aux.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   885
      End
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H80000018&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   6465
         MaskColor       =   &H00000000&
         Picture         =   "gw_p_gc_edificaciones_aux.frx":6CC3E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancelar"
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox Txt_descripcion11 
         DataField       =   "calle_denominacion"
         Height          =   645
         Left            =   2160
         TabIndex        =   12
         Text            =   "-"
         Top             =   360
         Width           =   4215
      End
      Begin VB.ComboBox dtc_codigo11 
         Height          =   315
         ItemData        =   "gw_p_gc_edificaciones_aux.frx":6CE48
         Left            =   120
         List            =   "gw_p_gc_edificaciones_aux.frx":6CE5B
         TabIndex        =   11
         Text            =   "CALLE"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Vía de Acceso"
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
         TabIndex        =   57
         Top             =   120
         Width           =   2070
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
         Left            =   2400
         TabIndex        =   56
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "gw_p_gc_edificaciones_aux.frx":6CE7C
      ScaleHeight     =   915
      ScaleWidth      =   8295
      TabIndex        =   29
      Top             =   4920
      Width           =   8350
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   4200
         MaskColor       =   &H00000000&
         Picture         =   "gw_p_gc_edificaciones_aux.frx":D8EAE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   2760
         Picture         =   "gw_p_gc_edificaciones_aux.frx":D979A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   45
      TabIndex        =   26
      Top             =   0
      Width           =   8535
      Begin VB.TextBox txt_campo6 
         DataField       =   "texto_borrar"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "gw_p_gc_edificaciones_aux.frx":D9F70
         Top             =   5520
         Visible         =   0   'False
         Width           =   6780
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ubicación del Proyecto"
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
         Height          =   2895
         Left            =   40
         TabIndex        =   38
         Top             =   1500
         Width           =   8445
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
            Left            =   7440
            Picture         =   "gw_p_gc_edificaciones_aux.frx":D9F72
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Nueva Via (Calle, Av, etc)"
            Top             =   1680
            Width           =   900
         End
         Begin VB.TextBox txt_campo2 
            DataField       =   "edif_referencia"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "gw_p_gc_edificaciones_aux.frx":DAA4A
            Top             =   2520
            Width           =   6135
         End
         Begin VB.TextBox txt_campo1 
            BackColor       =   &H00FFFFFF&
            DataField       =   "edif_nro"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "-"
            Top             =   2520
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAA4C
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5160
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo dtc_campo2 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAA65
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3360
            TabIndex        =   46
            Top             =   960
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_sigla"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAA7E
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2400
            TabIndex        =   45
            Top             =   360
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "correl_edif"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAA97
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2520
            TabIndex        =   44
            Top             =   960
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
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAAB0
            DataField       =   "pais_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6360
            TabIndex        =   39
            Top             =   195
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo9 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAAC9
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7320
            TabIndex        =   40
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
         Begin MSDataListLib.DataCombo dtc_desc9 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAAE2
            DataField       =   "prov_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4200
            TabIndex        =   3
            Top             =   600
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "prov_descripcion"
            BoundColumn     =   "prov_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAAFB
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3240
            TabIndex        =   41
            Top             =   315
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
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB14
            DataField       =   "depto_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "depto_descripcion"
            BoundColumn     =   "depto_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB2D
            DataField       =   "munic_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   1275
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB46
            DataField       =   "zona_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4200
            TabIndex        =   5
            Top             =   1275
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB5F
            DataField       =   "zona_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7320
            TabIndex        =   50
            Top             =   960
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB78
            DataField       =   "calle_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4080
            TabIndex        =   52
            Top             =   1440
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAB91
            DataField       =   "calle_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2865
            TabIndex        =   6
            Top             =   1800
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación Referencial Cercana"
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
            Index           =   8
            Left            =   2160
            TabIndex        =   54
            Top             =   2280
            Width           =   2805
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. del Edificio"
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
            Index           =   7
            Left            =   120
            TabIndex        =   53
            Top             =   2280
            Width           =   1410
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
            TabIndex        =   51
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
            Left            =   4200
            TabIndex        =   49
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label lbl_titulo3 
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
            Left            =   120
            TabIndex        =   47
            Top             =   1020
            Width           =   855
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
            Left            =   4200
            TabIndex        =   43
            Top             =   340
            Width           =   840
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
            TabIndex        =   42
            Top             =   340
            Width           =   1290
         End
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DABAA
         DataField       =   "codigo6"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   34
         Top             =   5280
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DABC3
         DataField       =   "codigo5"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   33
         Top             =   4800
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DABDC
         DataField       =   "edif_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         Appearance      =   0  'Flat
         DataField       =   "edif_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1100
         Width           =   8055
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DABF5
         DataField       =   "edif_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3945
         TabIndex        =   0
         Top             =   330
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAC0E
         DataField       =   "codigo5"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2475
         TabIndex        =   17
         Top             =   4485
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "gw_p_gc_edificaciones_aux.frx":DAC27
         DataField       =   "codigo6"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2115
         TabIndex        =   18
         Top             =   5460
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
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
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   5640
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lbl_titulo1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Edificio"
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
         Left            =   2520
         TabIndex        =   37
         Top             =   340
         Width           =   1500
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación del Proyecto de Edificación"
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
         TabIndex        =   36
         Top             =   830
         Width           =   3810
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   950
         TabIndex        =   35
         Top             =   330
         Width           =   1360
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Propietario/ Responsable"
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
         Height          =   360
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   4485
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Empresa o Institución"
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
         Left            =   120
         TabIndex        =   30
         Top             =   5595
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   240
         TabIndex        =   28
         Top             =   340
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   240
         Index           =   2
         Left            =   6840
         TabIndex        =   27
         Top             =   675
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   8655
      TabIndex        =   20
      Top             =   7065
      Width           =   8655
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   25
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   240
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
   Begin Crystal.CrystalReport CR01 
      Left            =   2160
      Top             =   6480
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2400
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4560
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
      Left            =   240
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   2400
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   4560
      Top             =   6960
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
      Left            =   6720
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   6720
      Top             =   6960
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   6720
      Top             =   7200
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
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   120
      Top             =   7080
      Visible         =   0   'False
      Width           =   5985
      _ExtentX        =   10557
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
      Caption         =   " <-- Inicio                   Viviendas - Edificaciones                   Fin -->"
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
Attribute VB_Name = "gw_p_gc_edificaciones_aux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod, VAR_PAIS As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String

Dim VAR_COD2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Dim RUTA1, RUTA2 As String
         'RUTA1 = "BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset("dtc_campo2")) + "\" + Trim(Ado_datos.Recordset("edif_codigo"))
         RUTA1 = "BIENES\EDIFICIOS\" + Trim(dtc_campo2) + "\" + Trim(txt_codigo)
         MsgBox "Se esta creando la carpeta: " + RUTA1
         MkDir RUTA1
        '        MkDir RUTA1 + "\CONTRATOS"
         rs_datos!estado_codigo = "APR"
         'rs_datos!fecha_registro = Date
         rs_datos!fecha_aprueba = Date
         'rs_datos!usr_codigo = glusuario
         rs_datos!usr_codigo_aprueba = glusuario
         rs_datos.UpdateBatch 'adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAux1_Click()
    'Validacion 1
    If dtc_codigo3 = "" Or dtc_codigo3 = "0" Then
        MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
        VAR_VAL = "ERR"
        Exit Sub
    End If
    FraGrabarCancelar.Visible = False
    Fra_aux1.Visible = True
    Fra_ABM.Enabled = False
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        'Call ABRIR_TABLA
        rs_datos.MoveFirst
        mbDataChanged = False
'        Fra_ABM.Enabled = False
'        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
'        dg_datos.Enabled = True
        txt_codigo.Enabled = True
    End If
   Unload Me
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(Dtc_aux2) + 1))
        'var_cod = RTrim(RTrim(left(dtc_codigo2.Text,1)) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        Set rstbeneaux = New ADODB.Recordset
        SQL_FOR = "select * from gc_edificaciones where edif_codigo = '" & var_cod & "'  "
        rstbeneaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rstbeneaux.RecordCount > 0 Then
            MsgBox " CODIGO DUPLICADO, Vuelva a intentar..."
            Exit Sub
        End If
        txt_codigo.Caption = var_cod
        rs_datos!edif_codigo = var_cod
        rs_datos!edif_codigo_corto = LTrim(Str(Val(Dtc_aux2) + 1))
        rs_datos!estado_codigo = "REG"
        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
        rs_datos!archivo_foto_cargado = "N"
        'rs_datos!ges_gestion = Year(Date)
        'rs_datos!correl_da = 0
        db.Execute "Update gc_municipio Set correl_edif = CAST('" & Dtc_aux2.Text & "' AS INT) + 1 Where munic_codigo= '" & dtc_codigo2.Text & "' "
        db.Execute "Update gc_departamento Set correl_edif = CAST('" & Dtc_aux2.Text & "' AS INT) + 1 Where depto_codigo= '" & Left(var_cod, 1) & "' "
     End If
     rs_datos!edif_tipo = dtc_codigo1.Text
     rs_datos!edif_descripcion = RTrim(Txt_descripcion.Text)
     
     rs_datos!pais_codigo = VAR_PAIS     'dtc_codigo7
     rs_datos!depto_codigo = dtc_codigo8.Text
     rs_datos!prov_codigo = dtc_codigo9.Text
     rs_datos!munic_codigo = IIf(dtc_codigo2.Text = "", "NN", dtc_codigo2.Text)
     
     rs_datos!zona_codigo = IIf(dtc_codigo3.Text = "", "0", dtc_codigo3.Text)
     rs_datos!calle_codigo = IIf(dtc_codigo4.Text = "", "0", dtc_codigo4.Text)
     
     rs_datos!edif_nro = txt_campo1
     rs_datos!edif_referencia = RTrim(Txt_campo2)
     
     'rs_datos!texto_borrar = ""
     rs_datos!loc_eje_X = "0"
     rs_datos!loc_eje_Y = "0"
     rs_datos!latitud = "0"
     rs_datos!longitud = "0"  'IIf(txt_campo3.Text = "", "0", txt_campo3.Text)
     rs_datos!altitud_snm = "0"  'IIf(txt_campo4.Text = "", "0", txt_campo4.Text)
     If rs_datos!beneficiario_codigo = "0" Or IsNull(rs_datos!beneficiario_codigo) Then
        rs_datos!beneficiario_codigo = IIf(glBenef = "0" Or glBenef = " ", "0", glBenef)
     'Else
        'rs_datos!beneficiario_codigo =
     End If
     'rs_datos!beneficiario_codigo = IIf(txt_ci = "0" Or txt_ci = " ", "0", txt_ci)
     rs_datos!beneficiario_codigo_empresa = "0"
     
     rs_datos!CORREL = "0"
     
     If rs_datos!ARCHIVO_Foto = ".JPG" Or rs_datos!ARCHIVO_Foto = "" Then
        rs_datos!ARCHIVO_Foto = txt_codigo.Caption + ".JPG"
     End If
'     sino = MsgBox("Desea APROBAR los datos de este Registro? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
         Dim RUTA1, RUTA2 As String
         'RUTA1 = "BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset("dtc_campo2")) + "\" + Trim(Ado_datos.Recordset("edif_codigo"))
         RUTA1 = "BIENES\EDIFICIOS\" + Trim(dtc_campo2) + "\" + Trim(txt_codigo)
         MsgBox "Se esta creando la carpeta: " + RUTA1
         'RmDir RUTA1
         MkDir RUTA1
        '        MkDir RUTA1 + "\CONTRATOS"
         rs_datos!estado_codigo = "APR"
         'rs_datos!fecha_registro = Date
         rs_datos!fecha_aprueba = Date
         'rs_datos!usr_codigo = glusuario
         rs_datos!usr_codigo_aprueba = glusuario
'         rs_datos.UpdateBatch 'adAffectAll
'      End If
     rs_datos!Fecha_Registro = Date
     rs_datos!usr_codigo = glusuario
     rs_datos.UpdateBatch adAffectAll
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
     Fra_ABM.Enabled = False
'      fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
'      dg_datos.Enabled = True
     'txt_codigo.Enabled = True
     If glPersNew = "NEWE" Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
        Set mw_solicitud.Ado_datos3.Recordset = rs_aux1
        mw_solicitud.dtc_codigo3.Text = txt_codigo.Caption
        'mw_solicitud.dtc_desc3.Text = Txt_descripcion.Text
        'mw_solicitud.dtc_aux3.Text = "0"
        mw_solicitud.dtc_desc3.BoundText = mw_solicitud.dtc_codigo3.BoundText
        mw_solicitud.dtc_aux3.BoundText = mw_solicitud.dtc_codigo3.BoundText
     End If
  End If
  'WWWWWWWWWWWWWWWWW
    Unload Me
  'WWWWWWWWWWWWWWWWW
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar el " + lbl_titulo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar la " + lbl_titulo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub CmdCancelaDet_Click()
    FraGrabarCancelar.Visible = True
    Fra_aux1.Visible = False
    Fra_ABM.Enabled = True
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
  If dtc_codigo3 = "" Or dtc_codigo3 = "0" Then
    MsgBox "Debe registrar: " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'INI Graba Calle
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select max(calle_codigo) as Codigo from gc_calles where zona_codigo = " & dtc_codigo3.Text & "    ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        If IsNull(rs_aux2!Codigo) Then
            VAR_COD2 = (Val(dtc_codigo3.Text) * 100) + 1
        Else
            VAR_COD2 = Round(CDbl(rs_aux2!Codigo) + 1, 0)
        End If
    Else
        VAR_COD2 = 1
    End If
    db.Execute "insert into gc_calles(zona_codigo, calle_codigo, calle_denominacion, calle_tipo, correl, estado_codigo, fecha_registro, usr_codigo)" & _
    "values ('" & dtc_codigo3.Text & "', " & VAR_COD2 & ", '" & Txt_descripcion11.Text & "', '" & dtc_codigo11.Text & "', '0', 'APR', '" & Date & "', '" & glusuario & "') "
    
   'FIN Graba Calle
    'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
    db.Execute "Update gc_zonas Set correl = " & VAR_COD2 & " Where zona_codigo= '" & dtc_codigo3.Text & "' "
    'gc_calles
    Call pnivel3(dtc_codigo3.BoundText)
    dtc_desc4.Enabled = True
    
    dtc_codigo4.Text = VAR_COD2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    FraGrabarCancelar.Visible = True
    Fra_aux1.Visible = False
    Fra_ABM.Enabled = True
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_desc8.BoundText = Dtc_aux2.BoundText
    dtc_codigo8.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_campo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_campo2.BoundText
'    dtc_aux2.BoundText = dtc_campo2.BoundText
    dtc_codigo2.BoundText = dtc_campo2.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    dtc_aux2.BoundText = dtc_codigo2.BoundText
    dtc_campo2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
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

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    Dtc_aux2.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    'Call pnivel7(dtc_codigo7.BoundText)
    
    Call pnivel7(VAR_PAIS)
    dtc_desc8.Enabled = True
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    dtc_aux2.BoundText = dtc_desc2.BoundText
    dtc_campo2.BoundText = dtc_desc2.BoundText
    Call pnivel2(dtc_codigo2.BoundText)
    dtc_desc3.Enabled = True
End Sub
   
Private Sub pnivel2(codigo2 As String)
   'Dim strConsultaF As String
     
   'strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo2 & "'"
   'strConsultaF = "select zona_codigo as codigo, zona_denominacion as descripcion, estado_codigo, fecha_registro, usr_codigo, correl as correl, pais_codigo As codigo1, depto_codigo As codigo2, prov_codigo As codigo3, munic_codigo As codigo4, comun_codigo As codigo5, hora_registro From gc_zonas where estado_codigo='APR' AND munic_codigo = '" & codigo2 & "' order by descripcion "
   'strConsultaF = "select zona_codigo as codigo, zona_denominacion as descripcion, estado_codigo, munic_codigo As codigo4 From gc_zonas where estado_codigo='APR' AND munic_codigo = '" & codigo2 & "' order by descripcion "
      
   Set dtc_codigo3.RowSource = Nothing
   'Set dtc_codigo3.RowSource = db.Execute(strConsultaF, "codigo2", adCmdText)
   Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo3.ReFill
   dtc_codigo3.BoundText = Empty
   
   Set dtc_desc3.RowSource = Nothing
   'Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc3.ReFill
   dtc_desc3.BoundText = Empty

End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    Call pnivel3(dtc_codigo3.BoundText)
    dtc_desc4.Enabled = True
End Sub
   
Private Sub pnivel3(codigo3 As String)
   'Dim strConsultaF As String
   
   'strConsultaF = "select * from gc_calles where zona_codigo = '" & codigo3 & "'"
   'strConsultaF = "select calle_codigo as codigo, calle_denominacion as descripcion, estado_codigo, fecha_registro, usr_codigo, correl as correl, zona_codigo As codigo1, calle_tipo As codigo2, hora_registro From gc_calles where estado_codigo='APR' AND zona_codigo = '" & codigo3 & "' order by descripcion "
   'strConsultaF = "select calle_codigo as codigo, calle_denominacion as descripcion, estado_codigo, zona_codigo As codigo1 From gc_calles where estado_codigo='APR' AND zona_codigo = '" & codigo3 & "' order by descripcion "

   Set dtc_codigo4.RowSource = Nothing
   'Set dtc_codigo4.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo4.ReFill
   dtc_codigo4.BoundText = Empty
   
   Set dtc_desc4.RowSource = Nothing
   'Set dtc_desc4.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc4.ReFill
   dtc_desc4.BoundText = Empty

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
    Call pnivel7(dtc_codigo7.BoundText)
    dtc_desc8.Enabled = True
End Sub
   
Private Sub pnivel7(codigo7 As String)
   Dim strConsultaF As String
     
   strConsultaF = "select * from gc_departamento where pais_codigo = '" & codigo7 & "'"
   Set dtc_codigo8.RowSource = Nothing
   Set dtc_codigo8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_codigo8.ReFill
   dtc_codigo8.BoundText = Empty
   
   Set dtc_desc8.RowSource = Nothing
   Set dtc_desc8.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc3.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_zonas '" & codigo2 & "' ")
   dtc_desc8.ReFill
   dtc_desc8.BoundText = Empty

End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    Dtc_aux2.BoundText = dtc_desc8.BoundText
    Call pnivel8(dtc_codigo8.BoundText)
    dtc_desc9.Enabled = True
End Sub

Private Sub pnivel8(codigo8 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo8 & "'"
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
    Call pnivel9(dtc_codigo9.BoundText)
    dtc_desc2.Enabled = True
End Sub
  
Private Sub pnivel9(codigo9 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo9 & "'"
   Set dtc_codigo2.RowSource = Nothing
   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo2.ReFill
   dtc_codigo2.BoundText = Empty
   
   Set dtc_desc2.RowSource = Nothing
   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc2.ReFill
   dtc_desc2.BoundText = Empty
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = True
'    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    Fra_aux1.Visible = False
      
'      On Error GoTo AddErr
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
'    Fra_ABM.Enabled = True
'    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
'    dg_datos.Enabled = False
    VAR_SW = "ADD"
    VAR_PAIS = "BOL"
    
    txt_codigo.Enabled = False
    dtc_desc8.Enabled = False
    dtc_desc9.Enabled = False
    dtc_desc2.Enabled = False
    dtc_desc3.Enabled = False
    dtc_desc4.Enabled = False
    
'    Txt_descripcion.SetFocus
'    BtnVer.Visible = False
  Exit Sub
AddErr:
  MsgBox Err.Description

End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    'txt_codigo.Enabled = True
    mbDataChanged = False
'    Fra_ABM.Enabled = False
'    dg_datos.Enabled = True
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
    VAR_SW = "ADD"
    VAR_PAIS = "BOL"
    FraGrabarCancelar.Visible = True
    Fra_aux1.Visible = False
    rs_datos.AddNew
    'Txt_descripcion.SetFocus
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_edificacion_tipo order by edif_tipo_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_gc_edificacion_tipo", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_municipio order by munic_descripcion", db, adOpenStatic
    'rs_datos2.Open "gp_listar_gc_municipio", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_zonas order by zona_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_gc_zonas", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from gc_calles order by calle_denominacion", db, adOpenStatic
    rs_datos4.Open "gp_listar_gc_calles", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "Select * from gc_beneficiario where (tipoben_codigo < 20 and tipoben_codigo <> 1) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos5.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    'rs_datos6.Open "Select * from gc_beneficiario where (tipoben_codigo > 19) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos6.Open "gp_listar_gc_beneficiario_empresas", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_pais where estado_codigo = 'APR' and pais_continente = 'AMERICA' ", db, adOpenKeyset
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'gc_Departamento  '<>
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_departamento order by depto_descripcion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'gc_provincia
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_provincia ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  'queryinicial = "select  *, edif_nro as campo1, edif_referencia as campo2, edif_tipo as codigo1, zona_codigo as codigo3, calle_codigo as codigo4, beneficiario_codigo as codigo5, beneficiario_codigo_empresa as codigo6, latitud As campo3, longitud As campo4, altitud_snm As campo5, loc_eje_X As campo6, loc_eje_Y As campo7 From gc_edificaciones "
  queryinicial = "select * From gc_edificaciones"
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  'Set Ado_datos.Recordset = rs_datos.DataSource
'  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'Esto mostrará la posición de registro actual para este Recordset
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
'    If Ado_datos.Recordset.AbsolutePosition >= 0 Then
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    End If
''    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
''        'BtnVer.Visible = True
''        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
''        Image2 = Img_Foto
''    Else
''        'BtnVer.Visible = False
''        'chkEstado.Value = vbUnchecked
''    End If
'  End If
'End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub


Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(codigoe As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_negociacion_cabecera WHERE edif_codigo = '" & codigoe & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub txt_campo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

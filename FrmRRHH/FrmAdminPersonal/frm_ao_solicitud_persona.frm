VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_ao_solicitud_persona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Contratación Personal - Detalle de Solicitud"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_solicitud_persona.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_solicitud_persona.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   8355
      TabIndex        =   25
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   480
         Picture         =   "frm_ao_solicitud_persona.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1440
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_solicitud_persona.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE LA SOLICITUD"
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
         Left            =   3255
         TabIndex        =   28
         Top             =   240
         Width           =   4245
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
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   1125
      Width           =   8415
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante existente en la Base de Datos"
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
         Left            =   3960
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante NUEVO"
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
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6840
         MaxLength       =   80
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   7560
         MaxLength       =   80
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtFecha 
         DataField       =   "benef_fecha"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   6480
         TabIndex        =   0
         Top             =   2925
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90832897
         CurrentDate     =   42370
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_ao_solicitud_persona.frx":D6864
         DataField       =   "ocup_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "ocup_codigo"
         BoundColumn     =   "ocup_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_solicitud_persona.frx":D687E
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   1120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_ao_solicitud_persona.frx":D6898
         DataField       =   "ocup_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   2295
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "ocup_descripcion"
         BoundColumn     =   "ocup_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_ao_solicitud_persona.frx":D68B2
         DataField       =   "nivel_educ_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   4320
         TabIndex        =   36
         Top             =   2295
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "nivel_educ_descripcion"
         BoundColumn     =   "nivel_educ_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_ao_solicitud_persona.frx":D68CC
         DataField       =   "nivel_educ_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   7440
         TabIndex        =   39
         Top             =   2040
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
         Bindings        =   "frm_ao_solicitud_persona.frx":D68E6
         DataField       =   "munic_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   3360
         TabIndex        =   40
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "munic_codigo"
         BoundColumn     =   "munic_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_ao_solicitud_persona.frx":D6900
         DataField       =   "munic_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   240
         TabIndex        =   41
         Top             =   3000
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "munic_descripcion"
         BoundColumn     =   "munic_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_solicitud_persona.frx":D691A
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   7320
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "puesto_codigo"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8400
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblbien 
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
         Index           =   10
         Left            =   240
         TabIndex        =   38
         Top             =   2760
         Width           =   1890
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nivel Educacional Requerido"
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
         Left            =   4320
         TabIndex        =   37
         Top             =   2025
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Correl.Postula"
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
         Left            =   6960
         TabIndex        =   34
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label txtBenef 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "benef_id"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
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
         Left            =   6960
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Solicitud"
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
         TabIndex        =   32
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Ejecutora (Solicitante)"
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
         Left            =   1665
         TabIndex        =   31
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
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
         TabIndex        =   30
         Top             =   495
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
         Left            =   1680
         TabIndex        =   29
         Top             =   495
         Width           =   5055
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Perfil Profesional"
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
         TabIndex        =   24
         Top             =   2025
         Width           =   1515
      End
      Begin VB.Label lblbien 
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
         Index           =   8
         Left            =   240
         TabIndex        =   6
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Fecha Postulación"
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
         Index           =   7
         Left            =   4560
         TabIndex        =   5
         Top             =   2940
         Visible         =   0   'False
         Width           =   1770
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Presupuesto requerido"
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   600
      TabIndex        =   8
      Top             =   1320
      Width           =   5535
      Begin MSMask.MaskEdBox mskMonto_pendiente 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   8421376
         ForeColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_limite 
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   8421376
         ForeColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_ext 
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   12648447
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto_nal 
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   12648447
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskMonto 
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label labPorcExt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0%"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pendiente de pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Límite del pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contraparte nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fuente externa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label labTipoMoneda 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "labTipoMoneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label labPorcTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label labPorcNal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0%"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   120
      Top             =   4680
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
      Left            =   2280
      Top             =   4680
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
      Left            =   4440
      Top             =   4680
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
      Left            =   120
      Top             =   5040
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
      Left            =   2280
      Top             =   5040
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Elije la Persona"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   46
      Top             =   2280
      Visible         =   0   'False
      Width           =   8320
      Begin VB.CommandButton BtnOk 
         BackColor       =   &H00C0C000&
         Caption         =   "Aceptar"
         Height          =   675
         Left            =   3000
         Picture         =   "frm_ao_solicitud_persona.frx":D6934
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1320
         Width           =   765
      End
      Begin VB.CommandButton BtnNo 
         BackColor       =   &H00C0C000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   4320
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_solicitud_persona.frx":D7336
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Cancelar"
         Top             =   1320
         Width           =   765
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "frm_ao_solicitud_persona.frx":D78C0
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   840
         TabIndex        =   49
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "frm_ao_solicitud_persona.frx":D78DA
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "frm_ao_solicitud_persona.frx":D78F4
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   2400
         TabIndex        =   51
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_primer_apellido"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "frm_ao_solicitud_persona.frx":D790E
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   1560
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_nombres"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux4 
         Bindings        =   "frm_ao_solicitud_persona.frx":D7928
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   3120
         TabIndex        =   53
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_telefono_Cel"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux2 
         Bindings        =   "frm_ao_solicitud_persona.frx":D7942
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   5400
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_segundo_apellido"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux5 
         Bindings        =   "frm_ao_solicitud_persona.frx":D795C
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   315
         Left            =   5880
         TabIndex        =   55
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   -2147483637
         ListField       =   "beneficiario_domicilio_legal"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Apellidos y Nombres"
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
         Left            =   840
         TabIndex        =   56
         Top             =   480
         Width           =   1890
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   2535
      Left            =   120
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   8320
      Begin VB.TextBox txtPat 
         DataField       =   "benef_primer_apellido"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   4320
         MaxLength       =   15
         TabIndex        =   63
         Top             =   495
         Width           =   3855
      End
      Begin VB.TextBox txtCI 
         DataField       =   "benef_codigo"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   62
         Top             =   495
         Width           =   2295
      End
      Begin VB.TextBox txtNom 
         DataField       =   "benef_nombres"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   61
         Top             =   1100
         Width           =   3855
      End
      Begin VB.TextBox txtTelefono 
         DataField       =   "benef_telefonos_ref"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   60
         Top             =   1545
         Width           =   2895
      End
      Begin VB.TextBox txtMat 
         DataField       =   "benef_segundo_apellido"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   59
         Top             =   1100
         Width           =   3855
      End
      Begin VB.TextBox txtDireccion 
         DataField       =   "benef_direccion_domicilio"
         DataSource      =   "frm_ao_solicitud_rrhh.ado_detalle2"
         Height          =   405
         Left            =   2160
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   2000
         Width           =   6015
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   69
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblbien 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         Caption         =   "Nro. Documento de Identidad "
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
         Height          =   435
         Index           =   4
         Left            =   240
         TabIndex        =   68
         Top             =   285
         Width           =   1515
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   67
         Top             =   855
         Width           =   1620
      End
      Begin VB.Label lblbien 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4320
         TabIndex        =   66
         Top             =   855
         Width           =   840
      End
      Begin VB.Label lblbien 
         BackColor       =   &H00404040&
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
         Left            =   240
         TabIndex        =   65
         Top             =   1530
         Width           =   2010
      End
      Begin VB.Label lblbien 
         BackColor       =   &H00404040&
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
         TabIndex        =   64
         Top             =   2040
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frm_ao_solicitud_persona"
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

Dim nomb2 As String

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
Para_Aceptado = "N"
Me.Hide
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas
'nomb2 = txtPat + " " + txtMat + " " + txtNom
If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("ges_gestion") = glGestion
      frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("unidad_codigo") = Txt_campo1.Text
      frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("solicitud_codigo") = txt_codigo
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
      'frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_id").Value = txtBenef.Text
   End If
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_primer_apellido") = "-"     'txtPat.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_segundo_apellido").Value = "-"     'txtMat.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_nombres").Value = "-"     'txtNom.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_direccion_domicilio").Value = "-"     'txtDireccion.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_telefonos_ref").Value = "0"     'txtTelefono.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_codigo").Value = "0"    'txtCI.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("puesto_codigo").Value = dtc_codigo1.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("ocup_codigo").Value = dtc_codigo2.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("munic_codigo").Value = dtc_codigo4.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("nivel_educ_codigo").Value = dtc_codigo3.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("observaciones") = dtc_desc1.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("benef_fecha") = TxtFecha.Value
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("estado_codigo") = "REG"
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("usr_codigo") = glusuario 'frmLogin.txtUserName.Text
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("fecha_registro") = Date
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    frm_ao_solicitud_rrhh.Ado_detalle2.Recordset.Update
    
    frm_ao_solicitud_rrhh.ABRIR_TABLA_DET
    
   Para_Aceptado = "S"
   'frm_ao_solicitud_rrhh.ado_detalle2.Refresh '.Recordset.Requery
   Unload Me
End If
End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
Valida = True
'If Val(Me.mskMonto) > Val(Me.mskMonto_pendiente) Then
'    ValidaMontos = False
'    MsgBox "El monto indicado sobrepasa el monto pendiente de pago", vbInformation
'    Me.mskMonto.SelStart = 0
'    Me.mskMonto.SelLength = Len(Me.mskMonto)
'    Me.mskMonto.SetFocus
'End If
    If dtc_codigo2 = "" Then
        Valida = False
    End If
    If dtc_codigo3 = "" Then
        Valida = False
    End If
End Function

Private Sub BtnNo_Click()
    Frame2.Visible = False
End Sub

Private Sub BtnOk_Click()
    txtCI.Text = dtc_codigo5.Text
    txtPat.Text = Trim(dtc_aux1.Text)
    txtMat.Text = Trim(dtc_aux2.Text)
    txtNom.Text = Trim(dtc_aux3.Text)
    txtTelefono.Text = Trim(dtc_aux4.Text)
    txtDireccion.Text = Trim(dtc_aux5.Text)
    Frame2.Visible = False
    Frame4.Visible = True
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux1.BoundText
    dtc_desc5.BoundText = dtc_aux1.BoundText
    dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
    dtc_aux4.BoundText = dtc_aux1.BoundText
    dtc_aux5.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux2.BoundText
    dtc_desc5.BoundText = dtc_aux2.BoundText
    dtc_aux1.BoundText = dtc_aux2.BoundText
    dtc_aux3.BoundText = dtc_aux2.BoundText
    dtc_aux4.BoundText = dtc_aux2.BoundText
    dtc_aux5.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux3.BoundText
    dtc_desc5.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
    dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux4.BoundText = dtc_aux3.BoundText
    dtc_aux5.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux4.BoundText
    dtc_desc5.BoundText = dtc_aux4.BoundText
    dtc_aux1.BoundText = dtc_aux4.BoundText
    dtc_aux2.BoundText = dtc_aux4.BoundText
    dtc_aux3.BoundText = dtc_aux4.BoundText
    dtc_aux5.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_aux5.BoundText
    dtc_desc5.BoundText = dtc_aux5.BoundText
    dtc_aux1.BoundText = dtc_aux5.BoundText
    dtc_aux2.BoundText = dtc_aux5.BoundText
    dtc_aux3.BoundText = dtc_aux5.BoundText
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
    dtc_aux1.BoundText = dtc_codigo5.BoundText
    dtc_aux2.BoundText = dtc_codigo5.BoundText
    dtc_aux3.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
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
    dtc_aux1.BoundText = dtc_desc5.BoundText
    dtc_aux2.BoundText = dtc_desc5.BoundText
    dtc_aux3.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Load()
'If glProceso = "CONSULTORIA" Then
'    Me.Caption = "Consultoría - Captura de datos personales"
'Else
'    Me.Caption = "Recursos Humanos - Captura de datos personales"
'End If
'Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA

    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
    rs_clasif1.Open "SELECT * FROM rc_puestos ORDER BY puesto_descripcion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    
    Set rs_clasif2 = New ADODB.Recordset
    If rs_clasif2.State = 1 Then rs_clasif2.Close
    rs_clasif2.Open "SELECT * FROM gc_ocupacion_profesion ORDER BY ocup_descripcion ", db, adOpenStatic
    Set Ado_clasif2.Recordset = rs_clasif2
    
    Set rs_clasif3 = New ADODB.Recordset
    If rs_clasif3.State = 1 Then rs_clasif3.Close
    rs_clasif3.Open "SELECT * FROM rc_nivel_educacional ORDER BY nivel_educ_descripcion ", db, adOpenStatic
    Set Ado_clasif3.Recordset = rs_clasif3
    
    Set rs_clasif4 = New ADODB.Recordset
    If rs_clasif4.State = 1 Then rs_clasif4.Close
    rs_clasif4.Open "SELECT * FROM gc_municipio where region_codigo = 'SI' ORDER BY munic_descripcion ", db, adOpenStatic
    Set Ado_clasif4.Recordset = rs_clasif4
    
    Set rs_clasif5 = New ADODB.Recordset
    If rs_clasif5.State = 1 Then rs_clasif5.Close
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    
	Call SeguridadSet(Me)
End Sub

Private Sub Option1_Click()
    Frame4.Visible = True
    Frame2.Visible = False
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame4.Visible = False
End Sub

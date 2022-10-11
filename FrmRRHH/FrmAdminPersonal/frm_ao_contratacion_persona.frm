VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_contratacion_persona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Contratación Personal - Registro de Proponentes"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_contratacion_persona.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_contratacion_persona.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   9075
      TabIndex        =   3
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   360
         Picture         =   "frm_ao_contratacion_persona.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1440
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_contratacion_persona.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE PROPONENTES"
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
         Left            =   3420
         TabIndex        =   6
         Top             =   240
         Width           =   4635
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
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   1125
      Width           =   9135
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   8600
         TabIndex        =   68
         Top             =   1090
         Width           =   290
      End
      Begin VB.Frame Frame5 
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
         ForeColor       =   &H00FFFF80&
         Height          =   2295
         Left            =   0
         TabIndex        =   52
         Top             =   4320
         Visible         =   0   'False
         Width           =   8925
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "frm_ao_contratacion_persona.frx":D6864
            DataField       =   "ocup_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   3600
            TabIndex        =   59
            Top             =   360
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
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frm_ao_contratacion_persona.frx":D687E
            DataField       =   "ocup_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Top             =   495
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "ocup_descripcion"
            BoundColumn     =   "ocup_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "frm_ao_contratacion_persona.frx":D6898
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   4560
            TabIndex        =   61
            Top             =   495
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "nivel_educ_descripcion"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "frm_ao_contratacion_persona.frx":D68B2
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   7920
            TabIndex        =   62
            Top             =   360
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D68CC
            DataField       =   "munic_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   3240
            TabIndex        =   63
            Top             =   960
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D68E6
            DataField       =   "munic_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2160
            TabIndex        =   64
            Top             =   1020
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "munic_descripcion"
            BoundColumn     =   "munic_codigo"
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtFecha 
            DataField       =   "cotiza_fecha"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   6360
            TabIndex        =   65
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   4210752
            CheckBox        =   -1  'True
            Format          =   92471297
            CurrentDate     =   42005
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha2 
            DataField       =   "cotiza_fecha_limite_postulacion"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   3360
            TabIndex        =   66
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   92471297
            CurrentDate     =   42005
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker txtFecha3 
            DataField       =   "cotiza_fecha_programada_contrato"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   360
            TabIndex        =   67
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   92471297
            CurrentDate     =   42005
            MinDate         =   2
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha.Inicio.Convocatoria"
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
            Left            =   120
            TabIndex        =   58
            Top             =   1530
            Width           =   2325
         End
         Begin VB.Label lbl_campo2 
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
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nivel Educacional (Mayor Importancia)"
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
            Left            =   4560
            TabIndex        =   56
            Top             =   240
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
            Left            =   120
            TabIndex        =   55
            Top             =   1050
            Width           =   1890
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha Límite Postulación"
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
            Left            =   3240
            TabIndex        =   54
            Top             =   1530
            Width           =   2235
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Fecha Presentacion Propuesta"
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
            Left            =   5880
            TabIndex        =   53
            Top             =   1530
            Width           =   2775
         End
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H000040C0&
         Caption         =   "Elije puesto convocado"
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
         TabIndex        =   51
         Top             =   1080
         Width           =   4215
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
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   8685
         Begin VB.CommandButton BtnNo 
            BackColor       =   &H00C0C000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4320
            MaskColor       =   &H00000000&
            Picture         =   "frm_ao_contratacion_persona.frx":D6900
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Cancelar"
            Top             =   1320
            Width           =   765
         End
         Begin VB.CommandButton BtnOk 
            BackColor       =   &H00C0C000&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   3000
            Picture         =   "frm_ao_contratacion_persona.frx":D6E8A
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1320
            Width           =   765
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frm_ao_contratacion_persona.frx":D788C
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   840
            TabIndex        =   17
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D78A6
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   120
            TabIndex        =   18
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D78C0
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2400
            TabIndex        =   19
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D78DA
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   120
            TabIndex        =   20
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D78F4
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   3120
            TabIndex        =   21
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D790E
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   5400
            TabIndex        =   22
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
            Bindings        =   "frm_ao_contratacion_persona.frx":D7928
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   5880
            TabIndex        =   23
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
            TabIndex        =   24
            Top             =   480
            Width           =   1890
         End
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5235
         MaxLength       =   80
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Caption         =   "Elije el Medio de Comunicación"
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
         Height          =   2055
         Left            =   240
         TabIndex        =   41
         Top             =   2040
         Visible         =   0   'False
         Width           =   8685
         Begin VB.CommandButton BtnOk2 
            BackColor       =   &H00C0C000&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   3000
            Picture         =   "frm_ao_contratacion_persona.frx":D7942
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1200
            Width           =   765
         End
         Begin VB.CommandButton BtnNo2 
            BackColor       =   &H00C0C000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4320
            MaskColor       =   &H00000000&
            Picture         =   "frm_ao_contratacion_persona.frx":D8344
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Cancelar"
            Top             =   1200
            Width           =   765
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frm_ao_contratacion_persona.frx":D88CE
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2520
            TabIndex        =   44
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "nivel_educ_descripcion"
            BoundColumn     =   "nivel_educ_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_ao_contratacion_persona.frx":D88E8
            DataField       =   "nivel_educ_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   315
            Left            =   2880
            TabIndex        =   46
            Top             =   360
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "Nombre del Medio"
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
            Left            =   720
            TabIndex        =   45
            Top             =   600
            Width           =   1680
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante Existente en la Base de Datos"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000040C0&
         Caption         =   "Postulante Nuevo"
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
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
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
         Height          =   2415
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Visible         =   0   'False
         Width           =   8685
         Begin VB.TextBox txtDireccion 
            DataField       =   "benef_direccion_domicilio"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   405
            Left            =   3000
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1800
            Width           =   5175
         End
         Begin VB.TextBox txtMat 
            DataField       =   "benef_segundo_apellido"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   34
            Top             =   1100
            Width           =   3855
         End
         Begin VB.TextBox txtTelefono 
            DataField       =   "benef_telefonos_ref"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   33
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtNom 
            DataField       =   "benef_nombres"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   285
            Left            =   4320
            MaxLength       =   30
            TabIndex        =   30
            Top             =   1100
            Width           =   3855
         End
         Begin VB.TextBox txtCI 
            DataField       =   "benef_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   27
            Top             =   495
            Width           =   2655
         End
         Begin VB.TextBox txtPat 
            DataField       =   "benef_primer_apellido"
            DataSource      =   "frm_ao_contratacion.ado_detalle2"
            Height          =   285
            Left            =   4320
            MaxLength       =   15
            TabIndex        =   26
            Top             =   495
            Width           =   3855
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
            Left            =   3000
            TabIndex        =   37
            Top             =   1530
            Width           =   2115
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
            TabIndex        =   35
            Top             =   1530
            Width           =   2010
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
            TabIndex        =   32
            Top             =   855
            Width           =   840
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
            TabIndex        =   31
            Top             =   855
            Width           =   1620
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
            TabIndex        =   29
            Top             =   240
            Width           =   2715
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
            TabIndex        =   28
            Top             =   240
            Width           =   1380
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         MaxLength       =   80
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtSW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         MaxLength       =   80
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_contratacion_persona.frx":D8902
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   1080
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_contratacion_persona.frx":D891C
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
         Height          =   315
         Left            =   7680
         TabIndex        =   50
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label lbl_convoca 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
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
         Left            =   7800
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cód. RRHH"
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
         Left            =   6480
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
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
         TabIndex        =   40
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Convocatoria"
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
         Left            =   7760
         TabIndex        =   12
         Top             =   220
         Width           =   1200
      End
      Begin VB.Label txtBenef 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "rrhh_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
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
         Left            =   6480
         TabIndex        =   11
         Top             =   495
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9120
         Y1              =   960
         Y2              =   960
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
         TabIndex        =   10
         Top             =   220
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
         TabIndex        =   9
         Top             =   220
         Width           =   1560
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle2"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1140
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
         TabIndex        =   7
         Top             =   480
         Width           =   6015
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   360
      Top             =   7680
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
      Top             =   7680
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
      Top             =   7680
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
      Top             =   8040
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
      Top             =   8040
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
      Top             =   8040
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
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   6840
      Top             =   7800
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
End
Attribute VB_Name = "frm_ao_contratacion_persona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Para_Aceptado As String
Dim rs_datos As New ADODB.Recordset

Dim rs_clasif1 As New ADODB.Recordset
Dim rs_clasif2 As New ADODB.Recordset
Dim rs_clasif3 As New ADODB.Recordset
Dim rs_clasif4 As New ADODB.Recordset
Dim rs_clasif5 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim nomb2 As String
Dim puesto2 As String

Dim VAR_TIME As Integer

Private Sub BtnCancelar_Click()
'cancela la edicion de datos
    Para_Aceptado = "N"
'    txtSW = "0"
    Unload Me
End Sub

Private Sub BtnGrabar_Click()
'acepta las modificaciones realizadas

If Valida Then
    Dim SQLS As String
    SQLS = ""
   'If txtSW = "ADD" Then
   If Option1.Value = True Then
   
  Set rs_aux3 = New ADODB.Recordset
   If rs_aux3.State = 1 Then rs_aux3.Close
   queryinicial = "select * from gc_beneficiario WHERE beneficiario_codigo ='" & Trim(txtCI.Text) & "'"
   rs_aux3.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   If rs_aux3.RecordCount > 0 Then
   sino = MsgBox("Esta persona ya EXISTE en la BASE DE DATOS", vbInformation, "SOFIA")
   Exit Sub
   End If
   ''
 Ado_datos.Recordset.AddNew
    Ado_datos.Recordset!beneficiario_codigo = Trim(txtCI.Text)
     Ado_datos.Recordset!depto_sigla = "LPZ"
     Ado_datos.Recordset!beneficiario_iniciales = Left(Trim(txtPat.Text), 1) & Left(Trim(txtMat.Text), 1) & Left(Trim(txtNom.Text), 1)
     Ado_datos.Recordset!tipodoc_codigo = "C.I"
     Ado_datos.Recordset!tipoben_codigo = "0"
     Ado_datos.Recordset!beneficiario_nit = txtCI.Text
     Ado_datos.Recordset!beneficiario_primer_apellido = Trim(txtPat.Text)
     Ado_datos.Recordset!beneficiario_segundo_apellido = Trim(txtMat.Text)
     Ado_datos.Recordset!beneficiario_nombres = Trim(txtNom.Text)
     Ado_datos.Recordset!beneficiario_denominacion = Trim(txtPat.Text) & " " & Trim(txtMat.Text) & " " & Trim(txtNom.Text)
    
     Ado_datos.Recordset!beneficiario_telefono_fijo = "0"
     Ado_datos.Recordset!beneficiario_telefono_Of = "0"
     Ado_datos.Recordset!beneficiario_telefono_Cel = txtTelefono.Text
     Ado_datos.Recordset!beneficiario_email = "-"
     Ado_datos.Recordset!beneficiario_email_of = "-"
     Ado_datos.Recordset!beneficiario_domicilio_legal = txtDireccion.Text
     Ado_datos.Recordset!pais_codigo = "BOL"
     Ado_datos.Recordset!depto_codigo = Left(Trim(dtc_codigo4.Text), 1)
     Ado_datos.Recordset!prov_codigo = Left(Trim(dtc_codigo4.Text), 1) & 1
     Ado_datos.Recordset!munic_codigo = Trim(dtc_codigo4.Text)
     Ado_datos.Recordset!zona_codigo = "0"
     Ado_datos.Recordset!calle_codigo = "0"
     Ado_datos.Recordset!edif_codigo = "0"
     
     Ado_datos.Recordset!beneficiario_edif_nro = "0"
     Ado_datos.Recordset!beneficiario_edif_piso_nro = "0"
     Ado_datos.Recordset!beneficiario_edif_depto_nro = "0"
    
     Ado_datos.Recordset!estado_codigo = "APR"
     Ado_datos.Recordset!fecha_registro = Date
     Ado_datos.Recordset!usr_codigo = glusuario
     Ado_datos.Recordset.Update
End If
   
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      frm_ao_contratacion.Ado_detalle2.Recordset("ges_gestion") = glGestion
      frm_ao_contratacion.Ado_detalle2.Recordset("unidad_codigo") = Txt_campo1.Text
      frm_ao_contratacion.Ado_detalle2.Recordset("solicitud_codigo") = txt_codigo
      frm_ao_contratacion.Ado_detalle2.Recordset("rrhh_codigo").Value = frm_ao_contratacion.Ado_datos.Recordset("rrhh_codigo")
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
   End If
    frm_ao_contratacion.Ado_detalle2.Recordset("puesto_codigo").Value = dtc_codigo1.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("benef_primer_apellido") = txtPat.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("benef_segundo_apellido").Value = txtMat.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("benef_nombres").Value = txtNom.Text
    nomb2 = Trim(txtPat) + " " + Trim(txtMat) + " " + Trim(txtNom)
    frm_ao_contratacion.Ado_detalle2.Recordset("beneficiario_denominacion").Value = Trim(nomb2)
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & txtCI.Text & "'   ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        frm_ao_contratacion.Ado_detalle2.Recordset("benef_id").Value = Ado_clasif5.Recordset!beneficiario_id
    Else
        frm_ao_contratacion.Ado_detalle2.Recordset("benef_id").Value = 0
    End If
    frm_ao_contratacion.Ado_detalle2.Recordset("beneficiario_codigo").Value = txtCI.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("ocup_codigo").Value = dtc_codigo2.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("munic_codigo").Value = dtc_codigo4.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("nivel_educ_codigo").Value = dtc_codigo3.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("observaciones") = dtc_desc1.Text
    
    frm_ao_contratacion.Ado_detalle2.Recordset("benef_direccion_domicilio").Value = txtDireccion.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("benef_telefonos_ref").Value = txtTelefono.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("cotiza_fecha") = TxtFecha.Value
    frm_ao_contratacion.Ado_detalle2.Recordset("cotiza_fecha_limite_postulacion").Value = TxtFecha2.Value
    frm_ao_contratacion.Ado_detalle2.Recordset("cotiza_fecha_programada_contrato").Value = TxtFecha3.Value
    
    frm_ao_contratacion.Ado_detalle2.Recordset("usr_codigo") = glusuario 'frmLogin.txtUserName.Text
    frm_ao_contratacion.Ado_detalle2.Recordset("fecha_registro") = Date
    frm_ao_contratacion.Ado_detalle2.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    
    sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        Select Case frm_ao_contratacion.Ado_datos.Recordset("modalidad_codigo")
            Case "INVD"    'INVITACION DIRECTA
                frm_ao_contratacion.Ado_detalle2.Recordset("estado_codigo") = "APR"
                Call GRABA_CALIFICA
            Case "CPEX"    'CONVOCATORIA PUBLICA EXTERNA
                frm_ao_contratacion.Ado_detalle2.Recordset("estado_codigo") = "APR"
                Call GRABA_CALIFICA
            Case "CPIN"    'CONVOCATORIA PUBLICA INTERNA
                frm_ao_contratacion.Ado_detalle2.Recordset("estado_codigo") = "APR"
                Call GRABA_CALIFICA
        End Select
    Else
        frm_ao_contratacion.Ado_detalle2.Recordset("estado_codigo") = "REG"
    End If

    frm_ao_contratacion.Ado_detalle2.Recordset.Update
   'db.Execute "update ro_rrhh_apertura_sobres set cotiza_codigo = " & txtBenef.Text & "  "'
   db.Execute "Update ro_rrhh_apertura_sobres Set cotiza_codigo = " & frm_ao_contratacion.Ado_detalle2.Recordset!cotiza_codigo & " Where rrhh_codigo = " & frm_ao_contratacion.Ado_detalle2.Recordset!rrhh_codigo & "   "
   Para_Aceptado = "S"
   'frm_ao_solicitud_rrhh.ado_detalle2.Refresh '.Recordset.Requery
'   txtSW = "0"
   frm_ao_contratacion.ABRIR_TABLA_DET
   Unload Me
End If
End Sub

Private Sub GRABA_CALIFICA()
    db.Execute "Insert INTO ro_rrhh_apertura_sobres (ges_gestion, rrhh_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, observaciones, puesto_codigo, ocup_codigo, nivel_educ_codigo, munic_codigo, estado_codigo, usr_codigo, fecha_registro, modalidad_codigo, cotiza_codigo) Values ('" & glGestion & "', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "',  '" & txtCI.Text & "', '" & Txt_campo1.Text & "', '" & txt_codigo.Caption & "', '" & nomb2 & "', '" & dtc_codigo1.Text & "', " & dtc_codigo2.Text & ", " & dtc_codigo3.Text & ", '" & dtc_codigo4.Text & "', 'REG', '" & glusuario & "',  '" & Date & "', '" & frm_ao_contratacion.Ado_datos.Recordset!modalidad_codigo & "', " & Val(lbl_convoca.Caption) & ")"
    '
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
  If (dtc_codigo4.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
  If txtPat = "" Then
        Valida = False
    End If
    If txtNom = "" Then
        Valida = False
    End If
End Function

Private Sub BtnNo_Click()
    Frame2.Visible = False
    Frame3.Visible = False
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
    Frame3.Visible = False
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

Private Sub dtc_desc1_LostFocus()
    If txtSW = "IDIR" Then
        Option1.Visible = False
        Option2.Visible = False
        Frame3.Visible = True
    Else
        Option1.Visible = True
        Option2.Visible = True
        Frame3.Visible = False
    End If
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
    Set rs_datos = New ADODB.Recordset
   If rs_datos.State = 1 Then rs_datos.Close
   queryinicial = "select * from gc_beneficiario WHERE  tipoben_codigo < 20 "
   'where usr_usuario= '" & GlUsuario & "' or usr_usuario= 'ADMIN'
   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic, adCmdText
   rs_datos.Sort = "beneficiario_denominacion"
   Set Ado_datos.Recordset = rs_datos


'If glProceso = "CONSULTORIA" Then
'    Me.Caption = "Consultoría - Captura de datos personales"
'Else
'    Me.Caption = "Recursos Humanos - Captura de datos personales"
'End If
'Para_Aceptado = "N"
'LOS DATOS PERSONALES SE CARGAN EN EL FORMULARIO QUE LO LLAMA
    'txtSW = "0"
    parametro = Aux
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
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
    'rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE estado_codigo_contrato = 'REG' AND tipoben_codigo < '20' ORDER BY beneficiario_denominacion ", DB, adOpenStatic
    rs_clasif5.Open "SELECT * FROM gc_beneficiario WHERE tipoben_codigo < '20' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5

End Sub

Private Sub Option1_Click()
    Frame4.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
'    txtSW = "1"
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame4.Visible = False
    Frame3.Visible = False
'    txtSW = "2"
End Sub

Private Sub Option3_Click()
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    'puesto2 = dtc_codigo1.Text
    dtc_desc1.Visible = True
    Option1.Visible = True
    Option2.Visible = True
    Frame5.Visible = True
    Option3.Visible = False
End Sub

Private Sub txtFecha2_LostFocus()
    TxtFecha.Value = TxtFecha2.Value
    'Me.Print Format(DateDiff("y", Fecha_Inicial, Fecha_Final), Formato) & " dias"
    VAR_TIME = DateDiff("y", TxtFecha3, TxtFecha2)
    If Val(VAR_TIME) < 0 Then
        MsgBox "La Fecha Límite Postulación NO puede ser MENOR a la Fecha Inicio Convocatoria, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        TxtFecha2.SetFocus
    End If
End Sub

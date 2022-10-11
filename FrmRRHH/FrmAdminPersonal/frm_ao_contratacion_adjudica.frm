VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_contratacion_adjudica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Personal - Contratación Personal - Personal Seleccionado"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_ao_contratacion_adjudica.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_contratacion_adjudica.frx":6A41E
      ScaleHeight     =   915
      ScaleWidth      =   9075
      TabIndex        =   4
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   480
         Picture         =   "frm_ao_contratacion_adjudica.frx":D6450
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   1560
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_contratacion_adjudica.frx":D665A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_bitacora 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERSONAL SELECCIONADO"
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
         Left            =   3330
         TabIndex        =   7
         Top             =   240
         Width           =   4335
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
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txt_monto3 
         DataField       =   "beneficiario_otro_mensual_bs"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   285
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   55
         Top             =   4935
         Width           =   1575
      End
      Begin VB.TextBox txt_tiempo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "beneficiario_tiempo_meses"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   54
         Top             =   4215
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Datos de la Persona a Contratar"
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
         Height          =   2655
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   8685
         Begin VB.TextBox Text6 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8260
            TabIndex        =   53
            Top             =   2060
            Width           =   260
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8260
            TabIndex        =   52
            Top             =   1340
            Width           =   260
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   2500
            TabIndex        =   51
            Top             =   1340
            Width           =   260
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8260
            TabIndex        =   50
            Top             =   620
            Width           =   260
         End
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D6864
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   5760
            TabIndex        =   24
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_segundo_apellido"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux1 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D687E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   2400
            TabIndex        =   21
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_primer_apellido"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.CommandButton BtnNo 
            BackColor       =   &H00C0C000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4440
            MaskColor       =   &H00000000&
            Picture         =   "frm_ao_contratacion_adjudica.frx":D6898
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Cancelar"
            Top             =   1200
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnOk 
            BackColor       =   &H00C0C000&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   3120
            Picture         =   "frm_ao_contratacion_adjudica.frx":D6E22
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1200
            Visible         =   0   'False
            Width           =   765
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D7824
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   2280
            TabIndex        =   19
            Top             =   1680
            Visible         =   0   'False
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D783E
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D7858
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_nombres"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux4 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D7872
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   5520
            TabIndex        =   23
            Top             =   1320
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
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "frm_ao_contratacion_adjudica.frx":D788C
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   315
            Left            =   240
            TabIndex        =   25
            Top             =   2040
            Width           =   8295
            _ExtentX        =   14631
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
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
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
            Left            =   240
            TabIndex        =   49
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
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
            Left            =   5760
            TabIndex        =   48
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblbien 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
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
            Left            =   2400
            TabIndex        =   47
            Top             =   360
            Width           =   1380
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
            Left            =   5520
            TabIndex        =   46
            Top             =   1080
            Width           =   2010
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
            TabIndex        =   45
            Top             =   1800
            Width           =   2115
         End
         Begin VB.Label lblbien 
            BackColor       =   &H00000000&
            Caption         =   "Nro. Doc.de Identidad "
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
            TabIndex        =   44
            Top             =   360
            Width           =   2115
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
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   1890
         End
      End
      Begin VB.TextBox txt_monto2 
         DataField       =   "beneficiario_haber_mensual_bs"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   285
         Left            =   525
         MaxLength       =   20
         TabIndex        =   43
         Top             =   4935
         Width           =   1575
      End
      Begin VB.TextBox txt_monto1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DataField       =   "beneficiario_monto_adjudica_bs"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   42
         Top             =   4935
         Width           =   1575
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5235
         MaxLength       =   80
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Elija el Postulante a Contratar"
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
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   8655
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
         TabIndex        =   16
         Top             =   1080
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
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   8685
         Begin VB.TextBox txtDireccion 
            DataField       =   "benef_direccion_domicilio"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   405
            Left            =   3000
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   1800
            Width           =   5175
         End
         Begin VB.TextBox txtMat 
            DataField       =   "benef_segundo_apellido"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   32
            Top             =   1100
            Width           =   3855
         End
         Begin VB.TextBox txtTelefono 
            DataField       =   "benef_telefonos_ref"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtNom 
            DataField       =   "benef_nombres"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   285
            Left            =   4320
            MaxLength       =   30
            TabIndex        =   30
            Top             =   1100
            Width           =   3855
         End
         Begin VB.TextBox txtCI 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   285
            Left            =   240
            MaxLength       =   15
            TabIndex        =   29
            Top             =   495
            Width           =   2655
         End
         Begin VB.TextBox txtPat 
            DataField       =   "beneficiario_primer_apellido"
            DataSource      =   "frm_ao_contratacion.ado_detalle3"
            Height          =   285
            Left            =   4320
            MaxLength       =   15
            TabIndex        =   28
            Top             =   495
            Width           =   3855
         End
      End
      Begin VB.TextBox txt_campo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         MaxLength       =   80
         TabIndex        =   14
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
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtFecha 
         DataField       =   "beneficiario_fecha_inicio"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   4215
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94371841
         CurrentDate     =   41640
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_contratacion_adjudica.frx":D78A6
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   1080
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "puesto_descripcion"
         BoundColumn     =   "puesto_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_contratacion_adjudica.frx":D78C0
         DataField       =   "puesto_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   315
         Left            =   7320
         TabIndex        =   15
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
      Begin MSComCtl2.DTPicker txtFecha2 
         DataField       =   "beneficiario_fecha_fin"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   315
         Left            =   3720
         TabIndex        =   38
         Top             =   4215
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94371841
         CurrentDate     =   41640
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFecha3 
         DataField       =   "beneficiario_fecha_contrato"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
         Height          =   315
         Left            =   2280
         TabIndex        =   39
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94371841
         CurrentDate     =   41640
         MinDate         =   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sueldo Mensual                                       Refrigerio/Otro                                        Total Contrato"
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
         Left            =   600
         TabIndex        =   57
         Top             =   4680
         Width           =   7590
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Inicio                                            Fecha de Fin                                        Tiempo (Meses)"
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
         Left            =   600
         TabIndex        =   56
         Top             =   3960
         Width           =   7800
      End
      Begin VB.Label lbl_convoca 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
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
         TabIndex        =   41
         Top             =   495
         Width           =   1095
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
         Left            =   6480
         TabIndex        =   40
         Top             =   240
         Width           =   990
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
         TabIndex        =   36
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Contrato"
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
         Left            =   7755
         TabIndex        =   13
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label txtBenef 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "rrhh_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
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
         TabIndex        =   12
         Top             =   495
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Tramite"
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_contratacion.ado_detalle3"
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
         TabIndex        =   9
         Top             =   495
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
         TabIndex        =   8
         Top             =   495
         Width           =   4695
      End
   End
   Begin MSAdodcLib.Adodc Ado_clasif1 
      Height          =   330
      Left            =   600
      Top             =   6600
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
      Top             =   6600
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
      Top             =   6600
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
      Top             =   6960
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
      Top             =   6960
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
      Top             =   6960
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
Attribute VB_Name = "frm_ao_contratacion_adjudica"
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
Dim rs_aux3 As New ADODB.Recordset

Dim nomb2 As String
Dim VAR_OCUP As String

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
   If swnuevo = 1 Then
      'DB.Execute "Insert INTO ro_Beneficiario_Dependiente (beneficiario_codigo, cod_dependiente, Cod_asegurado, Fecha_asegurado, fecha_nacimiento, primer_apellido, segundo_apellido, nombres, cod_pariente, nomb_pariente, estado_codigo, beneficiario_denominacion, ocupacion_pariente) Values ('" & txtBenef.Text & "', '" & txtCI.Text & "', '" & TxtItem.Text & "', '" & DTPFec_Seguro.Value & "', '" & txtNac.Value & "', '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', " & dtc_codigo1.Text & ", '" & dtc_desc1.Text & "', '" & txtEstado.Text & "', '" & nomb2 & "', '" & TxtOcupacion & "')"
      ''" & txtBenef.Caption & "',
       'DB.Execute "Insert INTO ao_solicitud_persona (ges_gestion, unidad_codigo, solicitud_codigo, benef_primer_apellido, benef_segundo_apellido, benef_nombres, benef_direccion_domicilio, benef_telefonos_ref, benef_codigo, puesto_codigo, ocup_codigo, munic_codigo, nivel_educ_codigo, observaciones, benef_fecha, estado_codigo, fecha_registro, usr_codigo) Values ('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
       '('" & glGestion & "', '" & txt_campo1.Text & "', " & txt_codigo.Caption & ", '" & txtPat.Text & "', '" & txtMat.Text & "', '" & txtNom.Text & "', '" & txtDireccion.Text & "', " & txtTelefono.Text & ", '" & txtCI.Text & "', " & dtc_codigo1.Text & ", " & dtc_codigo2.Text & ", '" & dtc_codigo4.Text & "', '" & dtc_codigo3.Text & "', '" & dtc_desc1.Text & "', '" & txtFecha.Value & "', 'REG', '" & Date & "', '" & GlUsuario & "')"
      frm_ao_contratacion.Ado_detalle3.Recordset("ges_gestion") = glGestion
      frm_ao_contratacion.Ado_detalle3.Recordset("unidad_codigo") = Txt_campo1.Text
      frm_ao_contratacion.Ado_detalle3.Recordset("solicitud_codigo") = txt_codigo
      frm_ao_contratacion.Ado_detalle3.Recordset("rrhh_codigo").Value = frm_ao_contratacion.Ado_datos.Recordset("rrhh_codigo")
      'frm_ao_contratacion.Ado_detalle3.Recordset("adjudica_codigo") = txt_codigo
   Else
      'DB.Execute "update ro_Beneficiario_Dependiente set beneficiario_codigo='" & txtBenef.Text & "', cod_dependiente='" & txtCI.Text & "', Cod_asegurado='" & TxtItem.Text & "', primer_apellido='" & txtPat.Text & "', segundo_apellido='" & txtMat.Text & "', nombres='" & txtNom.Text & "', cod_pariente=" & dtc_codigo1.Text & ", nomb_pariente='" & dtc_desc1.Text & "', estado_codigo='" & txtEstado.Text & "', beneficiario_denominacion='" & nomb2 & "'  "
      ' fecha_registro  hora_registro usr_usuario
   End If
    frm_ao_contratacion.Ado_detalle3.Recordset("puesto_codigo").Value = GlPuesto    'dtc_codigo1.Text
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_haber_mensual_bs") = Txt_monto2.Text
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_otro_mensual_bs") = Txt_monto3.Text
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_tiempo_meses") = txt_tiempo.Text
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_monto_adjudica_bs") = Txt_monto1.Text
    
    frm_ao_contratacion.Ado_detalle3.Recordset("tipo_moneda").Value = "BOB"
    'nomb2 = Trim(txtPat) + " " + Trim(txtMat) + " " + Trim(txtNom)
    frm_ao_contratacion.Ado_detalle3.Recordset("observaciones").Value = "CONTRATADO: " + Trim(dtc_desc5.Text)
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_codigo").Value = dtc_codigo5.Text
'    frm_ao_contratacion.Ado_detalle2.Recordset("ocup_codigo").Value = "10"  'IIf(IsNull(dtc_codigo2.Text), "10", dtc_codigo2.Text)
'    frm_ao_contratacion.Ado_detalle2.Recordset("munic_codigo").Value = "20101"  'dtc_codigo4.Text
'    'frm_ao_contratacion.Ado_detalle2.Recordset("nivel_educ_codigo").Value = dtc_codigo3.Text
    
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_fecha_inicio") = TxtFecha.Value
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_fecha_fin").Value = TxtFecha2.Value
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_fecha_contrato").Value = TxtFecha3.Value
    frm_ao_contratacion.Ado_detalle3.Recordset("beneficiario_fecha_adjudica") = Date
    
    frm_ao_contratacion.Ado_detalle3.Recordset("usr_codigo") = glusuario 'frmLogin.txtUserName.Text
    frm_ao_contratacion.Ado_detalle3.Recordset("fecha_registro") = Date
    frm_ao_contratacion.Ado_detalle3.Recordset("hora_registro") = Format(Time, "HH:mm:ss")
    If frm_ao_contratacion.Ado_detalle3.Recordset("estado_codigo") = "REG" Then
        sino = MsgBox("Desea APROBAR el Registro ? (Ya no podrá modificarlo)", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            Select Case frm_ao_contratacion.Ado_datos.Recordset("modalidad_codigo")
                Case "INVD"    'INVITACION DIRECTA
                    frm_ao_contratacion.Ado_detalle3.Recordset("estado_codigo") = "APR"
                    Call GRABA_FICHA
                Case "CPEX"    'CONVOCATORIA PUBLICA EXTERNA
                    frm_ao_contratacion.Ado_detalle3.Recordset("estado_codigo") = "APR"
                    Call GRABA_FICHA
                Case "CPIN"    'CONVOCATORIA PUBLICA INTERNA
                    frm_ao_contratacion.Ado_detalle3.Recordset("estado_codigo") = "APR"
                    Call GRABA_FICHA
            End Select
            db.Execute "update ro_rrhh_cabecera set estado_codigo = 'APR' where rrhh_codigo = " & txtBenef & " "
        Else
            frm_ao_contratacion.Ado_detalle3.Recordset("estado_codigo") = "REG"
        End If
    Else
        db.Execute "update ro_personal_contratado set beneficiario_haber_mensual = " & Txt_monto2.Text & ", beneficiario_otro_mensual = " & Txt_monto3.Text & ""
        db.Execute "update ro_rrhh_cabecera set estado_codigo = 'APR' where rrhh_codigo = " & txtBenef & " "
    End If
    frm_ao_contratacion.Ado_detalle3.Recordset.Update
   Para_Aceptado = "S"
   'frm_ao_solicitud_rrhh.ado_detalle2.Refresh '.Recordset.Requery
'   txtSW = "0"
   frm_ao_contratacion.ABRIR_TABLA_DET
   Unload Me
End If
End Sub

Private Sub GRABA_FICHA()
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "SELECT * FROM ro_rrhh_apertura_sobres where rrhh_codigo = " & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "  ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        VAR_OCUP = rs_aux3!ocup_codigo
    Else
        VAR_OCUP = "0"
    End If
    
'    db.Execute "Insert INTO ro_personal_contratado_new (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
'    db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro) Values ('" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "')"
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
    rs_aux2.Open "SELECT * FROM rc_puestos where puesto_codigo = '" & GlPuesto & "'  ", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        
        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, unidad_codigo, cargo_codigo, fecha_ingreso, fecha_expiracion, ocup_codigo, beneficiario_haber_mensual, estado_codigo, usr_codigo, fecha_registro, beneficiario_otro_mensual) Values (" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & ", '" & dtc_codigo5.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  '" & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_fecha_inicio & "', '" & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_fecha_fin & "', '" & VAR_OCUP & "', " & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_haber_mensual_bs & ", 'REG', '" & glusuario & "',  '" & Date & "',  " & Txt_monto3.Text & ")"
        
        db.Execute "Insert INTO ro_contratos_personas (id_contrato, beneficiario_codigo, codigo_contrato, ges_gestion, unidad_codigo, solicitud_codigo, numero_consultoria, doc_codigo, objeto_contrato, observacion_contrato, fecha_firma, establece_multas, cod_forma_inicio, fecha_inicio, fecha_fin, tiempo_num, tiempo_dmy, tipo_moneda, tc_us, monto_totalUS, monto_totalBS, cargo_codigo, puesto_codigo, pro_codigo, Codigo_Convenio, fte_codigo, org_codigo, porc_orgfin, porc_contra, estado_contrato, estado_confirmado, Estado_liquidacion , id_liquidacion, ARCHIVO, ARCHIVO_NOMB, usr_usuario, fecha_registro, monto_otroBS) " & _
        "values (" & frm_ao_contratacion.Ado_detalle3.Recordset!adjudica_codigo & ", '" & dtc_codigo5.Text & "', '" & rs_aux2!unidad_codigo & "' + '" & Str(frm_ao_contratacion.Ado_detalle3.Recordset!adjudica_codigo) & "', '" & glGestion & "', '" & rs_aux2!unidad_codigo & "', " & frm_ao_contratacion.Ado_detalle3.Recordset!solicitud_codigo & ", " & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & ", '0', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_descripcion & "', '" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_observaciones & "', '" & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_fecha_contrato & "', 'S', '0', '" & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_fecha_inicio & "', " & _
        " '" & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_fecha_fin & "', '0', 'MES', 'BOB', '6.96', '0', " & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_monto_adjudica_bs & ", '" & rs_aux2!cargo_codigo & "', '" & GlPuesto & "', '', '', '', '', '100', '0', 'REG', 'NO', 'REG', '0', '', '', '" & glusuario & "', '" & Date & "', " & Txt_monto3.Text & " )"
        
        'Values ('" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "', '" & rs_aux2!unidad_codigo & "',  '" & rs_aux2!cargo_codigo & "',  'REG', '" & glusuario & "',  '" & Date & "')"
        'ges_gestion, rrhh_codigo, adjudica_codigo, beneficiario_codigo, unidad_codigo, solicitud_codigo, puesto_codigo, beneficiario_monto_adjudica_bs,
        '              beneficiario_monto_adjudica_dol, beneficiario_haber_mensual_bs, beneficiario_haber_mensual_dol, beneficiario_tiempo_meses, tipo_moneda,
         '             beneficiario_fecha_inicio, beneficiario_fecha_fin, beneficiario_fecha_adjudica, beneficiario_fecha_contrato, proceso_codigo, subproceso_codigo, etapa_codigo,
          '            clasif_codigo , doc_codigo, doc_numero, cite_tramite, observaciones, estado_codigo, usr_codigo, fecha_registro, hora_registro

    Else
        db.Execute "Insert INTO ro_personal_contratado (rrhh_codigo, beneficiario_codigo, puesto_codigo, estado_codigo, usr_codigo, fecha_registro, beneficiario_haber_mensual, beneficiario_otro_mensual) Values ('" & frm_ao_contratacion.Ado_datos.Recordset!rrhh_codigo & "', '" & txtCI.Text & "', '" & GlPuesto & "',  'REG', '" & glusuario & "',  '" & Date & "', " & frm_ao_contratacion.Ado_detalle3.Recordset!beneficiario_haber_mensual_bs & ", " & Txt_monto3.Text & ")"
    End If
    'Set Ado_clasif1.Recordset = rs_aux2

End Sub

Function Valida()
'valida que el monto asignado al beneficiario no sobrepase el monto pendiente de asignacion
  Valida = True
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    Valida = False
  End If
'  If (dtc_codigo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    Valida = False
'  End If
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
    Frame2.Visible = False
'    Frame3.Visible = False
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
'    Frame3.Visible = False
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

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux1.BoundText = dtc_codigo5.BoundText
    dtc_aux2.BoundText = dtc_codigo5.BoundText
    dtc_aux3.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
'    Set rs_clasif1 = New ADODB.Recordset
'    If rs_clasif1.State = 1 Then rs_clasif1.Close
'    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
'    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & Txt_campo1 & "' ORDER BY puesto_descripcion ", db, adOpenStatic
'    Set Ado_clasif1.Recordset = rs_clasif1
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
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

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux1.BoundText = dtc_desc5.BoundText
    dtc_aux2.BoundText = dtc_desc5.BoundText
    dtc_aux3.BoundText = dtc_desc5.BoundText
    dtc_aux4.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Activate()
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    'rs_clasif1.Open "SELECT * FROM rc_puestos where puesto_vacante = 'SI' ORDER BY puesto_descripcion  ", DB, adOpenStatic
    'rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & Txt_campo1 & "' ORDER BY puesto_descripcion ", db, adOpenStatic
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
    'rs_clasif5.Open "SELECT * FROM rv_beneficiario_invitacion where puesto_codigo  = '" & GlPuesto & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_clasif5.Open "SELECT * FROM rv_beneficiario_calificacion where puesto_codigo  = '" & GlPuesto & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_clasif5.Recordset = rs_clasif5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux1.BoundText = dtc_codigo5.BoundText
    dtc_aux2.BoundText = dtc_codigo5.BoundText
    dtc_aux3.BoundText = dtc_codigo5.BoundText
    dtc_aux4.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText

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
End Sub

Private Sub Option1_Click()
    Frame4.Visible = True
    Frame2.Visible = False
'    Frame3.Visible = False
'    txtSW = "1"
End Sub

Private Sub Option2_Click()
    Frame2.Visible = True
    Frame4.Visible = False
'    Frame3.Visible = False
'    txtSW = "2"
    Set rs_clasif1 = New ADODB.Recordset
    If rs_clasif1.State = 1 Then rs_clasif1.Close
    rs_clasif1.Open "SELECT * FROM rv_puestos_solicitud where unidad_codigo_sol = '" & parametro & "' and solicitud_codigo = " & txt_codigo.Caption & " ORDER BY puesto_descripcion ", db, adOpenStatic
    Set Ado_clasif1.Recordset = rs_clasif1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    Option2.Visible = False

End Sub

Private Sub txt_monto3_LostFocus()
    If Txt_monto3.Text = "" Then
        Txt_monto3.Text = "0"
    End If
    If txt_tiempo.Text = "" Or txt_tiempo.Text = "0" Then
        txt_tiempo.Text = "1"
    End If
    Txt_monto1.Text = (CDbl(Txt_monto2.Text) + CDbl(Txt_monto3.Text)) * CDbl(txt_tiempo.Text)
End Sub

Private Sub txtFecha_LostFocus()
    TxtFecha3.Value = TxtFecha.Value
End Sub

Private Sub txtFecha2_LostFocus()
    'Me.Print Format(DateDiff("m", Fecha_Inicial, Fecha_Final), Formato) & " meses"
    txt_tiempo = DateDiff("m", TxtFecha, TxtFecha2)
    If Val(txt_tiempo) < 0 Then
        MsgBox "La Fecha de Inicio NO puede ser MAYOR a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        TxtFecha2.SetFocus
    End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ao_solicitud_cotiza_datosE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización Venta - Datos Complementarios Cotiza (Europa)"
   ClientHeight    =   10410
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12030
   ControlBox      =   0   'False
   Icon            =   "frm_ao_solicitud_cotiza_datosE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000011&
      FillColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   11715
      TabIndex        =   55
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000011&
         Height          =   620
         Left            =   1455
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_solicitud_cotiza_datosE.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Cancelar"
         Top             =   60
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000011&
         Height          =   620
         Left            =   120
         Picture         =   "frm_ao_solicitud_cotiza_datosE.frx":12EE
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATOS COMPLEMENTARIOS COTIZACION"
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
         Left            =   3990
         TabIndex        =   58
         Top             =   200
         Width           =   6375
      End
   End
   Begin VB.Frame Fra_datos99 
      BackColor       =   &H00C0C0C0&
      Height          =   8175
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   11775
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   10800
         TabIndex        =   68
         Top             =   6015
         Width           =   270
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   10800
         TabIndex        =   67
         Top             =   6660
         Width           =   270
      End
      Begin VB.PictureBox Fra_datos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   11460
         TabIndex        =   27
         Top             =   1080
         Width           =   11520
         Begin VB.TextBox Txt_campo5 
            DataField       =   "cotiza_nro_montador"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   9840
            TabIndex        =   36
            Text            =   "2"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "modelo_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5445
            TabIndex        =   35
            Top             =   555
            Width           =   2295
         End
         Begin VB.TextBox Txt_campo8 
            DataField       =   "dimension_fosa_m"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   9600
            TabIndex        =   34
            Text            =   "0"
            Top             =   2565
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo9 
            DataField       =   "dimension_fosa_fondo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   285
            Left            =   9600
            TabIndex        =   33
            Text            =   "0"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo10 
            DataField       =   "dimension_fosa_frente"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   285
            Left            =   9600
            TabIndex        =   32
            Text            =   "0"
            Top             =   1755
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo3 
            DataField       =   "cotiza_luz"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   285
            Left            =   2880
            TabIndex        =   31
            Text            =   "0"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo2 
            DataField       =   "cotiza_energia"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   285
            Left            =   2880
            TabIndex        =   30
            Text            =   "0"
            Top             =   1755
            Width           =   1365
         End
         Begin VB.TextBox Txt_campo7 
            DataField       =   "bien_cotiza_num_accesos"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   29
            Text            =   "0"
            Top             =   2565
            Width           =   1365
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   310
            Left            =   9120
            TabIndex        =   28
            Top             =   4680
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1AC4
            DataField       =   "pais_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   360
            TabIndex        =   37
            Top             =   1275
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "pais_descripcion"
            BoundColumn     =   "pais_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1ADD
            DataField       =   "tipo_eqp"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   360
            TabIndex        =   38
            Top             =   525
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "tipo_eqp_descripcion"
            BoundColumn     =   "tipo_eqp"
            Text            =   "ASCENSOR SOCIAL"
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
         Begin MSDataListLib.DataCombo dtc_codigo21 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1AF7
            DataField       =   "bien_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   39
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "36NO"
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
         Begin MSDataListLib.DataCombo dtc_desc21 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B12
            DataField       =   "bien_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   40
            Top             =   2145
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B2C
            DataField       =   "pais_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   2520
            TabIndex        =   41
            Top             =   960
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
         Begin MSDataListLib.DataCombo dtc_desc24 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B45
            DataField       =   "bien_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   42
            Top             =   1785
            Visible         =   0   'False
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "modelo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B5F
            DataField       =   "tipo_eqp"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   3600
            TabIndex        =   43
            Top             =   240
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "tipo_eqp"
            BoundColumn     =   "tipo_eqp"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B78
            DataField       =   "marca_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   5760
            TabIndex        =   44
            Top             =   960
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "marca_codigo"
            BoundColumn     =   "marca_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1B92
            DataField       =   "marca_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   3435
            TabIndex        =   45
            Top             =   1275
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "marca_descripcion"
            BoundColumn     =   "marca_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1BAB
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   9765
            TabIndex        =   46
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1BC5
            DataField       =   "beneficiario_codigo"
            DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
            Height          =   315
            Left            =   6960
            TabIndex        =   47
            Top             =   1275
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Equipo"
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
            Left            =   360
            TabIndex        =   54
            Top             =   225
            Width           =   1755
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"frm_ao_solicitud_cotiza_datosE.frx":1BDE
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
            Left            =   480
            TabIndex        =   53
            Top             =   1770
            Width           =   8970
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"frm_ao_solicitud_cotiza_datosE.frx":1C71
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
            Left            =   480
            TabIndex        =   52
            Top             =   2175
            Width           =   8970
         End
         Begin VB.Label lbl_campo7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frm_ao_solicitud_cotiza_datosE.frx":1D03
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
            Left            =   480
            TabIndex        =   51
            Top             =   2580
            Width           =   8970
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "País de Origen del Equipo                 Marca del Equipo                                            Proveedor del Equipo"
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
            Left            =   360
            TabIndex        =   50
            Top             =   1005
            Width           =   8670
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Montadores"
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
            Left            =   9360
            TabIndex        =   49
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label Label11 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo del Equipo"
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
            Left            =   5520
            TabIndex        =   48
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.TextBox Txt_campo14 
         DataField       =   "modelo_otras_caracteristicas"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   525
         Left            =   480
         TabIndex        =   26
         Text            =   "0"
         Top             =   7200
         Width           =   9765
      End
      Begin VB.TextBox Txt_campo12 
         Alignment       =   2  'Center
         DataField       =   "dimension_cabina_lado"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Text            =   "0"
         Top             =   4575
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo11 
         Alignment       =   2  'Center
         DataField       =   "dimension_cabina_frente"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   4575
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo13 
         Alignment       =   2  'Center
         DataField       =   "dimension_cabina_alto"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   8160
         TabIndex        =   11
         Text            =   "0"
         Top             =   4575
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_desc61 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1D97
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datosE"
         Height          =   315
         Left            =   5520
         TabIndex        =   9
         Top             =   5295
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_descripcion"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo61 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1DB1
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datosE"
         Height          =   315
         Left            =   9120
         TabIndex        =   10
         Top             =   4935
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_codigo"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1DCC
         DataField       =   "modelo_motor"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   4920
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "modelo_motor"
         BoundColumn     =   "modelo_motor"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1DE6
         DataField       =   "modelo_motor"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   5295
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "motor_descripcion"
         BoundColumn     =   "modelo_motor"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1DFF
         DataField       =   "boton_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   4080
         TabIndex        =   22
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "boton_codigo"
         BoundColumn     =   "boton_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1E19
         DataField       =   "boton_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   480
         TabIndex        =   23
         Top             =   5955
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "boton_descripcion_cabina"
         BoundColumn     =   "boton_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1E32
         DataField       =   "senal_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   4080
         TabIndex        =   24
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "senal_codigo"
         BoundColumn     =   "senal_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1E4C
         DataField       =   "senal_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datos0"
         Height          =   315
         Left            =   480
         TabIndex        =   25
         Top             =   6600
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "senal_descripcion_cabina"
         BoundColumn     =   "senal_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux5 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1E65
         DataField       =   "boton_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   69
         Top             =   6000
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "boton_descripcion_pasillo"
         BoundColumn     =   "boton_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux6 
         Bindings        =   "frm_ao_solicitud_cotiza_datosE.frx":1E7E
         DataField       =   "senal_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   70
         Top             =   6645
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "senal_descripcion_pasillo"
         BoundColumn     =   "senal_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Edificio"
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
         Left            =   9840
         TabIndex        =   66
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Txt_campo2A 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO-"
         DataField       =   "edif_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         Height          =   300
         Left            =   9840
         TabIndex        =   65
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Trámite "
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
         Left            =   360
         TabIndex        =   64
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   2160
         TabIndex        =   63
         Top             =   240
         Width           =   2160
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         Height          =   300
         Left            =   360
         TabIndex        =   62
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label Txt_Correl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "cotiza_codigo"
         DataSource      =   "aw_p_ao_solicitud_cotiza_venta.Ado_datos"
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
         Height          =   300
         Left            =   7920
         TabIndex        =   61
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cotización"
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
         Left            =   8040
         TabIndex        =   60
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Txt_descripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2160
         TabIndex        =   59
         Top             =   510
         Width           =   5295
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Otras Características"
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
         Left            =   480
         TabIndex        =   19
         Top             =   6960
         Width           =   1860
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Señalización de Cabina                                                                 Señalización de Pasillo"
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
         Left            =   480
         TabIndex        =   18
         Top             =   6345
         Width           =   7155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Botonera de Cabina                                                                         Botonera de Pasillo"
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
         Left            =   480
         TabIndex        =   17
         Top             =   5700
         Width           =   6855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Motor"
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
         Left            =   480
         TabIndex        =   16
         Top             =   5040
         Width           =   825
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuarto de Control"
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
         Left            =   5760
         TabIndex        =   15
         Top             =   5040
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensión Cabina Frente (mm)                  Dimensión Cabina Lado (mm)                    Dimensión Cabina Alto (mm)"
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
         Left            =   480
         TabIndex        =   14
         Top             =   4320
         Width           =   9660
      End
      Begin VB.Label txt_conti 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "pais_continente"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datosE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5760
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_solicitud_cotiza_venta.Ado_datosE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
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
      ScaleWidth      =   12030
      TabIndex        =   0
      Top             =   10410
      Width           =   12030
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   1920
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   -360
      Top             =   9480
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   1920
      Top             =   9480
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
      Caption         =   "Ado_datos21"
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
      Left            =   4200
      Top             =   9480
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
      Left            =   6480
      Top             =   9480
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
   Begin MSAdodcLib.Adodc Ado_datos61 
      Height          =   330
      Left            =   8760
      Top             =   9480
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
      Caption         =   "Ado_datos61"
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
      Left            =   0
      Top             =   9840
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
End
Attribute VB_Name = "frm_ao_solicitud_cotiza_datosE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
'BUSCADOR

'OTROS
Dim var_cod As String
Dim VAR_VAL As String

Dim VAR_1A, VAR_2A As Double
Dim VAR_3B, VAR_4B, VAR_5B, VAR_6B, VAR_7B As Double
Dim VAR_8C, VAR_9C, VAR_10C, VAR_11C, VAR_12C As Double
Dim VAR_13D, VAR_14D As Double
Dim totbs2, totdl2, totbs3, totdl3 As Double
Dim VAR_SUBD, VAR_SUBB As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        mw_solicitud_cotiza_venta.Ado_datos0.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
'DE ASIA
  
'EUROPA
 On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & txt_conti & "' and cotiza_codigo = " & Txt_Correl.Caption & "    ", db, adOpenKeyset, adLockOptimistic
    'rs_datos10.Open "select * from ao_solicitud_cotiza_modelo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & txt_conti & "' and cotiza_codigo = " & Txt_Correl.Caption & "    ", db, adOpenKeyset, adLockOptimistic
    If rs_datos10.RecordCount > 0 Then
         
        'SOLO EL REGISTRO ACTIVO
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!pais_codigo = dtc_codigo7.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!bien_codigo = IIf(IsNull(dtc_codigo21.Text) Or dtc_codigo21.Text = "", "NA1", dtc_codigo21.Text)
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_fecha = Date     'DTPfecha1.Value
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!modelo_codigo = Txt_campo4.Text     '      'MODELO PROVISIONAL AUTOMATICO JQA 02-2015
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!modelo_codigo_h = "S/M"  'Txt_campo5.Text    'dtc_codigo41.Text
         'mw_solicitud_cotiza_venta.Ado_datos0.Recordset!modelo_codigo_x = "S/M"   'Txt_campo6.Text    'dtc_codigo51.Text
    
        '2
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_energia = Txt_campo2.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_luz = Txt_campo3.Text

         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!bien_cotiza_num_accesos = Txt_campo7.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_fosa_m = Txt_campo8.Text        '
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_fosa_fondo = Txt_campo9.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_fosa_frente = Txt_campo10.Text  '
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cuadro_ctrl_codigo = IIf((dtc_codigo61.Text = ""), 1, dtc_codigo61.Text)
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_nro_montador = IIf((Txt_campo5.Text = ""), "2", Txt_campo5.Text)
         'mw_solicitud_cotiza_venta.Ado_datos0.RECORDSET!Foto = Date
         'mw_solicitud_cotiza_venta.Ado_datos0.RECORDSET!ARCHIVO_Foto = var_cod + ".JPG"
         'mw_solicitud_cotiza_venta.Ado_datos0.RECORDSET!archivo_foto_cargado = "N"
         
         '3 hora_registro
         If parametro = "DNMOD" Then
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!proceso_codigo = "TEC"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!subproceso_codigo = "TEC-05"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!etapa_codigo = "TEC-05-01"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!clasif_codigo = "TEC"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!doc_codigo = "R-313"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!doc_numero = "0"  'txt_campo1.Text
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!poa_codigo = "3.2.7"
         Else
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!proceso_codigo = "COM"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!subproceso_codigo = "COM-01"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!etapa_codigo = "COM-01-04"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!clasif_codigo = "COM"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!doc_codigo = "R-224"
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!doc_numero = "0"  'txt_campo1.Text
            mw_solicitud_cotiza_venta.Ado_datos0.Recordset!poa_codigo = "3.1.1"
         End If
         
         'WWWWWWWWWWWWWW JQA 02-2015
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!marca_codigo = dtc_codigo3.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_cabina_frente = Txt_campo11.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_cabina_lado = Txt_campo12.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!dimension_cabina_alto = Txt_campo13.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!modelo_motor = dtc_codigo4.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!boton_codigo = dtc_codigo5.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!senal_codigo = dtc_codigo6.Text
             mw_solicitud_cotiza_venta.Ado_datos0.Recordset!modelo_otras_caracteristicas = Txt_campo14.Text
        'WWWWWWWWWWWWWW JQA 08-2015
         
         '4 - WWWWWWWWWWWWWW JQA 02-2015
            If VAR_COD2 < 10 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-00000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 9 And VAR_COD2 < 100 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-0000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 99 And VAR_COD2 < 1000 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-000" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 999 And VAR_COD2 < 10000 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-00" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 9999 And VAR_COD2 < 100000 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-0" + Trim(txt_codigo)
            End If
            If VAR_COD2 > 99999 Then
               mw_solicitud_cotiza_venta.Ado_datos0.Recordset!unidad_codigo_ant = parametro + "-" + Trim(txt_codigo)
            End If
  
         Txt_campo5.Text = mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_nro_montador
         '5 - mw_solicitud_cotiza_venta.Ado_datos0.RECORDSET!unidad_codigo_ant = var_cod2     'txt_codigo1.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!tipo_eqp = dtc_codigo2.Text
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!Fecha_Registro = Date     'no cambia
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
         mw_solicitud_cotiza_venta.Ado_datos0.Recordset.Update    'Batch 'adAffectAll
         
         db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_nro_montador = " & Txt_campo5.Text & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & txt_conti & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
         db.Execute "Update ao_solicitud_cotiza_venta Set pais_codigo = '" & dtc_codigo7.Text & "'  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & txt_conti & "' and cotiza_codigo = " & Txt_Correl.Caption & "  "
         
         'ACTUALIZA Proveedor
         
         db.Execute "update ao_solicitud_cotiza_modelo set beneficiario_codigo = '212391920010' where pais_codigo = '" & dtc_codigo7.Text & "' AND unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and cotiza_codigo = " & Txt_Correl.Caption & "  "
         MsgBox "Se guardó con éxito, la Cotización Nro.: " + Str(mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_codigo)

         'MsgBox "Se guardó con éxito, la Cotización Nro.: " + Str(mw_solicitud_cotiza_venta.Ado_datos0.Recordset!cotiza_codigo)
    End If
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Unload Me
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
    If (Txt_campo4 = "") Then
    MsgBox "Debe registrar el Modelo del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo7.Text = "") Then
    MsgBox "Debe registrar el Pais Origen del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo5.Text = "") Then
    MsgBox "Debe registrar cantidad de Montadores (Instaladores / Ajustadores) ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo2.Text = "") Then
    MsgBox "Debe registrar el Tipo de Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo3.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo9.Text = "" Then
    MsgBox "Debe registrar: Dimención Fosa Fondo (mm) ...", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo10.Text = "" Then
    MsgBox "Debe registrar: Dimención Fosa Frente (mm) ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo7.Text = "" Then
    MsgBox "Debe registrar:Número de Accesos ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo8.Text = "" Then
    MsgBox "Debe registrar: Espacio Libre Bajo Dintel (mm) ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo61 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If

'  If (dtc_codigo11 = "") Then
'    MsgBox "Debe registrar Parámetros de Cálculo, Consulte con el Administrador ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fob_me1 = "") Or (txt_fob_me1 = "0") Then
'    MsgBox "Debe registrar el Precio FOB del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo21_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    dtc_desc24.BoundText = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc21_Click(Area As Integer)
    dtc_codigo21.BoundText = dtc_desc21.BoundText
    dtc_desc24.BoundText = dtc_desc21.BoundText
End Sub

Private Sub dtc_desc24_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_desc24.BoundText
    dtc_codigo21.BoundText = dtc_desc24.BoundText
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_codigo61.BoundText = dtc_desc61.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLA
    mbDataChanged = False
'    If swnuevo = 2 Then
'        dtc_desc2.BoundText = dtc_codigo2.BoundText
'        dtc_desc3.BoundText = dtc_codigo3.BoundText
'    End If
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from ac_costos_comercializacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 0 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'AMERICA' order by pais_descripcion", db, adOpenStatic
    End If
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 1 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'ASIA' order by pais_descripcion", db, adOpenStatic
    End If
    If mw_solicitud_cotiza_venta.SSTab1.Tab = 2 Then
        rs_datos7.Open "Select * from gc_pais where pais_continente = 'EUROPA' order by pais_descripcion", db, adOpenStatic
    End If
'    rs_datos7.Open "Select * from gc_pais where pais_continente = '" & txt_conti & "' order by pais_descripcion", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    
    'ac_bienes_marcas
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from ac_bienes_marcas WHERE  (pais_codigo = 'ESP') ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'Tipo de Equipo
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from ac_bienes_equipo_tipos ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'Cuarto de Control
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cuadro_ctrl ", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

'Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub Fra_datos_Click()

End Sub

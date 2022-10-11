VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tw_solicitud_bienes7 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cotización - Detalle de Servicios"
   ClientHeight    =   7815
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   55
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   360
         Picture         =   "tw_solicitud_bienes7.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   120
         Width           =   1340
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1680
         MaskColor       =   &H00000000&
         Picture         =   "tw_solicitud_bienes7.frx":07D6
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE SERVICIOS"
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
         Left            =   5055
         TabIndex        =   56
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   10695
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
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
         Height          =   1335
         Left            =   120
         TabIndex        =   43
         Top             =   5160
         Width           =   10455
         Begin VB.TextBox Txt_campo10 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_venta_base"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   285
            Left            =   240
            TabIndex        =   50
            Text            =   "1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo16 
            Alignment       =   2  'Center
            DataField       =   "bien_cantidad"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   285
            Left            =   5400
            TabIndex        =   49
            Text            =   "1"
            Top             =   840
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_solicitud_bienes7.frx":10C2
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   1800
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unimed_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_solicitud_bienes7.frx":10DB
            DataField       =   "unimed_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   2520
            TabIndex        =   71
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "unimed_descripcion"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REG"
            DataField       =   "estado_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
            Left            =   9120
            TabIndex        =   52
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Txt_campo11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "bien_total_venta"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
            Left            =   7200
            TabIndex        =   51
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estado del  Registro"
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
            Index           =   2
            Left            =   9120
            TabIndex        =   48
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl_campo10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Precio Unitario Bs."
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
            TabIndex        =   47
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lbl_desc2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Unidad de Medida"
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
            Left            =   2520
            TabIndex        =   46
            Top             =   360
            Width           =   2370
         End
         Begin VB.Label lbl_campo16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad"
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
            Left            =   5280
            TabIndex        =   45
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lbl_campo11 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Bs."
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
            Left            =   7200
            TabIndex        =   44
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox Txt_campo14 
         DataField       =   "unimed_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   3480
         TabIndex        =   25
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt_campo15 
         DataField       =   "fosa_dimension_frente"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   5565
         TabIndex        =   24
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bienes NUEVOS"
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
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bien existente en la Base de Datos"
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
         Height          =   255
         Left            =   5280
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.TextBox Txt_campo17 
         DataField       =   "venta_o_compra"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   5640
         TabIndex        =   23
         Text            =   "V"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   4335
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   10455
         Begin VB.TextBox TxtObservacion 
            DataField       =   "observacion"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   525
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
            Top             =   3660
            Width           =   9975
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2865
            TabIndex        =   42
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   6345
            TabIndex        =   34
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   9945
            TabIndex        =   33
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9825
            TabIndex        =   30
            Top             =   1320
            Width           =   365
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "tw_solicitud_bienes7.frx":10F4
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   7840
            TabIndex        =   2
            Top             =   520
            Width           =   2380
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo4 
            Bindings        =   "tw_solicitud_bienes7.frx":110D
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   240
            TabIndex        =   32
            Top             =   1305
            Width           =   9955
            _ExtentX        =   17568
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "bien_descripcion_anterior"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            Bindings        =   "tw_solicitud_bienes7.frx":1126
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   3360
            TabIndex        =   35
            Top             =   2100
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "marca_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo3 
            Bindings        =   "tw_solicitud_bienes7.frx":113F
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   6840
            TabIndex        =   36
            Top             =   2100
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "modelo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo18 
            Bindings        =   "tw_solicitud_bienes7.frx":1158
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   240
            TabIndex        =   39
            Top             =   2100
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "pais_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "tw_solicitud_bienes7.frx":1171
            DataField       =   "fosa_dimension_frente"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   240
            TabIndex        =   68
            Top             =   2880
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "solicitud_tipo_descripcion"
            BoundColumn     =   "solicitud_tipo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "tw_solicitud_bienes7.frx":118A
            DataField       =   "fosa_dimension_frente"
            DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   6720
            TabIndex        =   69
            Top             =   2880
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            ListField       =   "solicitud_tipo"
            BoundColumn     =   "solicitud_tipo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "tw_solicitud_bienes7.frx":11A3
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle7"
            Height          =   315
            Left            =   240
            TabIndex        =   72
            Top             =   520
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Servicio - Para el Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   240
            TabIndex        =   75
            Top             =   3360
            Width           =   3645
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "El Servicio será utilizado para:"
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
            TabIndex        =   70
            Top             =   2595
            Width           =   2700
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Industria/Pais Origen"
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
            Top             =   1815
            Width           =   1860
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marca"
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
            Left            =   3360
            TabIndex        =   38
            Top             =   1815
            Width           =   570
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Modelo"
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
            Left            =   6840
            TabIndex        =   37
            Top             =   1815
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Caracteristicas"
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
            TabIndex        =   31
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label lbl_descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripcion"
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
            TabIndex        =   29
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl_codigo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   7920
            TabIndex        =   28
            Top             =   240
            Width           =   660
         End
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "tw_solicitud_bienes7.frx":11BC
         DataField       =   "bien_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
         Height          =   315
         Left            =   8400
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "par_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.CommandButton BtnVer2 
         BackColor       =   &H00808000&
         Caption         =   "Plano Corte Transversal"
         Height          =   360
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3840
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Cargar Plano Planta Tipo"
         Height          =   360
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSDataListLib.DataCombo dtc_aux2 
         Bindings        =   "tw_solicitud_bienes7.frx":11D5
         DataField       =   "bien_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
         Height          =   315
         Left            =   7200
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "subgrupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.PictureBox Img_Foto 
         Height          =   2055
         Left            =   5880
         ScaleHeight     =   1995
         ScaleWidth      =   1995
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Image Image1 
            DataField       =   "foto"
            DataSource      =   "aw_p_ao_solicitud.ado_detalle1"
            Height          =   1995
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.PictureBox Img_Foto2 
         Height          =   2055
         Left            =   8280
         ScaleHeight     =   1995
         ScaleWidth      =   1995
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
         Begin VB.Image Image2 
            DataField       =   "foto_bien"
            DataSource      =   "Ado_datos"
            Height          =   2002
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2002
         End
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "tw_solicitud_bienes7.frx":11EE
         DataField       =   "bien_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
         Height          =   315
         Left            =   5400
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483637
         ForeColor       =   0
         ListField       =   "grupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.Label lbl_det 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "par_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
         Left            =   7080
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl_edif 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
         Left            =   9120
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Frente1"
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
         TabIndex        =   27
         Top             =   3720
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "x  Fondo1"
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
         Left            =   3090
         TabIndex        =   26
         Top             =   3720
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_descripcion 
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
         Left            =   3720
         TabIndex        =   16
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle7"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
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
         Height          =   480
         Index           =   8
         Left            =   2800
         TabIndex        =   13
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbl_codigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Trámite"
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
         Left            =   320
         TabIndex        =   12
         Top             =   360
         Width           =   750
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
      ScaleWidth      =   10980
      TabIndex        =   5
      Top             =   7815
      Width           =   10980
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   7800
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2520
      Top             =   7800
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Left            =   4920
      Top             =   7800
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   120
      TabIndex        =   57
      Top             =   5040
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox Txt_campo5 
         DataField       =   "bien_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   240
         TabIndex        =   62
         Text            =   "0"
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox Txt_campo6 
         DataField       =   "bien_descripcion"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   2880
         TabIndex        =   61
         Text            =   "0"
         Top             =   525
         Width           =   7335
      End
      Begin VB.TextBox Txt_campo7 
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   1680
         TabIndex        =   60
         Text            =   "0"
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox Txt_campo8 
         DataField       =   "marca_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   960
         TabIndex        =   59
         Text            =   "0"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox Txt_campo9 
         DataField       =   "modelo_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle7"
         Height          =   285
         Left            =   6360
         TabIndex        =   58
         Text            =   "0"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Marca"
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
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Modelo"
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
         Left            =   5520
         TabIndex        =   66
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Caracteristicas"
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
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lbl_descripcion2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Descripcion"
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
         Left            =   2880
         TabIndex        =   64
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lbl_codigo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Left            =   240
         TabIndex        =   63
         Top             =   255
         Width           =   660
      End
   End
End
Attribute VB_Name = "tw_solicitud_bienes7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos1 As New ADODB.Recordset
Attribute rs_datos1.VB_VarHelpID = -1
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset

Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
'BUSCADOR
Dim var_val2 As String
Dim var_cod5 As String
Dim VAR_VAL As String
Dim var_ctm, var_itm As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        tw_identificacion_cliente.Ado_detalle2.Recordset.CancelUpdate
        
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
  
     Set rs_aux1 = New ADODB.Recordset
     SQL_FOR = "select * from ao_solicitud_bienes where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " and bien_codigo = '" & dtc_codigo1.Text & "'  "
     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
     If rs_aux1.RecordCount > 0 Then
        If swnuevo = 1 Then
            MsgBox "El bien ya existe, verifique los datos y vuelva a intentar ..."
            'var_cod = 0
            Exit Sub
        End If
     End If
     If lbl_det.Caption = "30000" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle3.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle3.Recordset("unidad_codigo").Value = Txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle3.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle3.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle3.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle3.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle3.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle3.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "20000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "24000", dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "24300", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
            tw_identificacion_cliente.Ado_detalle3.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle3.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle3.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle3.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle3.Recordset!bien_codigo_padre = GlExtension
            tw_identificacion_cliente.Ado_detalle3.Recordset.UpdateBatch adAffectAll
     End If
     If lbl_det.Caption = "39800" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle5.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle5.Recordset("unidad_codigo").Value = Txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle5.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle5.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle5.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle5.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle5.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle5.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle5.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle5.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle5.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle5.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "30000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle5.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "39000", dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle5.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "39810", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle5.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
            tw_identificacion_cliente.Ado_detalle5.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle5.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle5.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            tw_identificacion_cliente.Ado_detalle5.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle5.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle5.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle5.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle5.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle5.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle5.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle5.Recordset!bien_codigo_padre = GlExtension
            tw_identificacion_cliente.Ado_detalle5.Recordset.UpdateBatch adAffectAll
    End If
    If lbl_det.Caption = "34800" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle6.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle6.Recordset("unidad_codigo").Value = Txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle6.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle6.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle6.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle6.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle6.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle6.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "99900", dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
            tw_identificacion_cliente.Ado_detalle6.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle6.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle6.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle6.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle6.Recordset!bien_codigo_padre = GlExtension
            tw_identificacion_cliente.Ado_detalle6.Recordset.UpdateBatch adAffectAll
    End If
    If lbl_det.Caption = "24300" Then
         If swnuevo = 1 Then
            'tw_identificacion_cliente.Ado_detalle7.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle7.Recordset("ges_gestion").Value = tw_identificacion_cliente.Ado_datos.Recordset("ges_gestion")
            tw_identificacion_cliente.Ado_detalle7.Recordset("unidad_codigo").Value = Txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle7.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle7.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle7.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle7.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle7.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle7.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "20000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "24000", dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "24300", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
            tw_identificacion_cliente.Ado_detalle7.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle7.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle7.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle7.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle7.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle7.Recordset!bien_codigo_padre = GlExtension
            tw_identificacion_cliente.Ado_detalle7.Recordset.UpdateBatch adAffectAll
    End If
    If lbl_det.Caption = "21000" Then
         If swnuevo = 1 Then
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("ges_gestion").Value = glGestion
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("unidad_codigo").Value = Txt_campo1.Caption
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("estado_codigo").Value = "REG"
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("venta_o_compra").Value = "V"
    '        frm_ao_solicitud_servicio_basico.Ado_detalle6.Recordset("archivo_foto_cargado").Value = "N"
    '        frm_ao_solicitud_servicio_basico.Ado_detalle6.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
                frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "20000", dtc_aux1.Text)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "24000", dtc_aux2.Text)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "24300", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_precio_compra").Value = 0
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("bien_total_compra").Value = 0
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 1, dtc_codigo5.Text)
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset("usr_codigo").Value = glusuario
            frm_ao_solicitud_servicio_basico.Ado_detalle2.Recordset.UpdateBatch adAffectAll
    End If
    'End If
'     Set rs_aux1 = New ADODB.Recordset
'     SQL_FOR = "select * from ao_solicitud_edificacion where unidad_codigo = '" & mw_solicitud.Ado_datos.Recordset("unidad_codigo") & "' and solicitud_codigo = " & mw_solicitud.Ado_datos.Recordset("solicitud_codigo") & " and edif_codigo = '" & dtc_codigo1.Text & "'  "
'     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'     If rs_aux1.RecordCount > 0 Then
'        MsgBox "El código ya existe, consulte con el administrador del Sistema..."
'        var_cod = 0
'        Exit Sub
'     Else
'        mw_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
'     End If
     
     
'     var_cod = mw_solicitud.Ado_detalle1.Recordset.RecordCount
'     db.Execute "Update ao_solicitud Set correl_edificacion = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "  "
    If lbl_det = "43340" Then
        
      If swnuevo = 1 Then
     'Graba en Cotiza    1
        Set rs_aux4 = New ADODB.Recordset
        SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    ", db, adOpenStatic
            If Not rs_aux5.EOF Then
                 var_cod5 = IIf(IsNull(rs_aux5!Codigo), 1, rs_aux5!Codigo + 1)
            End If
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = Txt_campo1.Caption
            rs_aux4!solicitud_codigo = txt_codigo.Caption
            rs_aux4!trafico_codigo = "0"  'Ado_datos.Recordset!trafico_codigo
            rs_aux4!cotiza_codigo = var_cod5
            'Call correl_bien
            'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
            rs_aux4!bien_codigo = dtc_codigo1.Text       '"MAN-002"       'VAR_COD3  '"36NO-" + Trim(Str(VAR_COD2))
            rs_aux4!modelo_codigo = Txt_campo3.Text     'Ado_datos.Recordset!modelo_codigo
            rs_aux4!modelo_codigo_h = "0"        'Ado_datos.Recordset!modelo_codigo_h1
            rs_aux4!modelo_codigo_x = "0"       'Ado_datos.Recordset!modelo_codigo_x1
            rs_aux4!cotiza_fecha = Date
            rs_aux4!cotiza_cantidad = IIf(Txt_campo16 = "", 1, Txt_campo16)
            rs_aux4!cotiza_tdc_bol = GlTipoCambioOficial
            rs_aux4!cotiza_precio_fob_bs = IIf(Txt_campo10 = "", 0, Txt_campo10)
            rs_aux4!cotiza_precio_fob_dol = CDbl(Txt_campo10) * GlTipoCambioOficial
            rs_aux4!cotiza_precio_total_bs = IIf(Txt_campo11 = "", 0, Txt_campo11)
            rs_aux4!cotiza_precio_total_dol = CDbl(Txt_campo11) * GlTipoCambioOficial
            rs_aux4!costo_monto = IIf(Txt_campo11 = "", 0, Txt_campo11)
            '
            Select Case Txt_campo1.Caption
               Case "DNMAN"
                   rs_aux4!proceso_codigo = "TEC"
                   rs_aux4!subproceso_codigo = "TEC-02"
                   rs_aux4!etapa_codigo = "TEC-02-01"
                   rs_aux4!poa_codigo = "3.2.3"
                   rs_aux4!clasif_codigo = "TEC"
                   rs_aux4!doc_codigo = "R-355"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
               Case "DNREP"
                   rs_aux4!proceso_codigo = "TEC"
                   rs_aux4!subproceso_codigo = "TEC-03"
                   rs_aux4!etapa_codigo = "TEC-03-01"
                   rs_aux4!poa_codigo = "3.2.4"
                   rs_aux4!clasif_codigo = "TEC"
                   rs_aux4!doc_codigo = "R-362"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
               Case "UGADM"
                   rs_aux4!proceso_codigo = "FIN"
                   rs_aux4!subproceso_codigo = "FIN-03"
                   rs_aux4!etapa_codigo = "FIN-03-03"
                   rs_aux4!poa_codigo = "4.2.3"
                   rs_aux4!clasif_codigo = "ADM"
                   rs_aux4!doc_codigo = "R-146"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014

            End Select

            If lbl_det.Caption = "24300" Then
                rs_aux4!edif_codigo = tw_identificacion_cliente.dtc_codigo3.Text
                
            End If
            If lbl_det.Caption = "21000" Then
                rs_aux4!edif_codigo = frm_ao_solicitud_servicio_basico.dtc_codigo3.Text
                rs_aux4!proceso_codigo = "TEC"
                rs_aux4!subproceso_codigo = "TEC-2"
                rs_aux4!etapa_codigo = "TEC-2-1"
                rs_aux4!poa_codigo = "3.2.3"
                rs_aux4!clasif_codigo = "TEC"
                rs_aux4!doc_codigo = "R-362"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
            End If
            
            rs_aux4!doc_numero = "0"
            
            
            rs_aux4!estado_codigo = "REG"
            rs_aux4!Fecha_Registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update
            db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod5 & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
      End If
    End If
'     Frame1.Visible = False
'     Frame2.Visible = False

     Unload Me

'     Call ABRIR_TABLA
'     rs_datos.MoveLast
'     mbDataChanged = False
'
'      Fra_ABM.Enabled = False
'      fraOpciones.Visible = True
'      FraGrabarCancelar.Visible = False
'      dg_datos.Enabled = True
'      txt_codigo.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar el " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'     Txt_campo12.Caption = var_itm
'     Txt_campo13.Caption = var_ctm
  
'        If Txt_campo2.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo3.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo4.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo5.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo6.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo7.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo8.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo9.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo10.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo11.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
End Sub

Private Sub BtnVer_Click()
'  On Error GoTo QError
'  If mw_solicitud.Ado_detalle1.Recordset("estado_codigo") = "REG" Then
'    Dim ARCH_FOTO As String
'    Dim SW0 As String
'    If mw_solicitud.Ado_detalle1.Recordset!archivo_foto_cargado = "N" Then
'      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_aux3.Text) & "\" & Trim(dtc_codigo1.Text) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FED2"
''      If GlServidor = "SRVPRO" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'      SW0 = 1
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(dtc_aux3.Text) & "\" & Trim(dtc_codigo1.Text) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FED2"
''          If GlServidor = "SRVPRO" Then
''            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
''          Else
'            e = NombreCarpeta
''          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'          SW0 = 1
'      Else
'        SW0 = 0
'      End If
'    End If
'    If SW0 = 1 Then
'    '    If GlServidor = "SRVPRO" Then
'    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    '    Else
'            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(dtc_aux3.Text) + "\" + Trim(dtc_codigo1.Text) + "\" + Trim(dtc_codigo1.Text) + "-A.JPG"
'    '    End If
'        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
'        CodBien = mw_solicitud.Ado_detalle1.Recordset!edif_codigo
'        'If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'        If Guardar_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = " & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & " and edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'            MsgBox "Se cargo la Imagen Correctamente !!"
'        Else
'            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'        End If
'    Else
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    End If
'  Else
'    MsgBox "Debe Aprobar el registro, para crear la carpeta correspondiente..."
'  End If
'QError:
'    ' Manejo de errores
'    If Err.Number > 0 Then
'        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    '    db.RollbackTrans
'        Screen.MousePointer = vbDefault
'    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux1.BoundText
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
    Txt_campo2.BoundText = dtc_aux1.BoundText
    Txt_campo3.BoundText = dtc_aux1.BoundText
    Txt_campo4.BoundText = dtc_aux1.BoundText
    Txt_campo18.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux2.BoundText
    dtc_desc1.BoundText = dtc_aux2.BoundText
    dtc_aux1.BoundText = dtc_aux2.BoundText
    dtc_aux3.BoundText = dtc_aux2.BoundText
    Txt_campo2.BoundText = dtc_aux2.BoundText
    Txt_campo3.BoundText = dtc_aux2.BoundText
    Txt_campo4.BoundText = dtc_aux2.BoundText
    Txt_campo18.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux3.BoundText
    dtc_desc1.BoundText = dtc_aux3.BoundText
    dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
    Txt_campo2.BoundText = dtc_aux3.BoundText
    Txt_campo3.BoundText = dtc_aux3.BoundText
    Txt_campo4.BoundText = dtc_aux3.BoundText
    Txt_campo18.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Change()
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    dtc_codigo2.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    dtc_codigo2.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo1_LostFocus()
    If Len(TxtObservacion.Text) = 0 Then
        TxtObservacion.Text = dtc_desc1.Text + " - " + Txt_campo4
    End If
End Sub

Private Sub dtc_codigo2_Change()
    dtc_desc2.BoundText = dtc_codigo2.Text
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.Text
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Change()
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
    Txt_campo2.BoundText = dtc_desc1.BoundText
    Txt_campo3.BoundText = dtc_desc1.BoundText
    Txt_campo4.BoundText = dtc_desc1.BoundText
    Txt_campo18.BoundText = dtc_desc1.BoundText
    dtc_codigo2.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
    Txt_campo2.BoundText = dtc_desc1.BoundText
    Txt_campo3.BoundText = dtc_desc1.BoundText
    Txt_campo4.BoundText = dtc_desc1.BoundText
    Txt_campo18.BoundText = dtc_desc1.BoundText
    dtc_codigo2.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    If Len(TxtObservacion.Text) = 0 Then
        TxtObservacion.Text = dtc_desc1.Text + " - " + Txt_campo4
    End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
    var_val2 = "2"
    Frame1.Visible = True
    Select Case lbl_det
        Case "43340"
            Label1.Caption = "DETALLE DE BIENES (Equipos)"
            Option1.Caption = "Equipo NUEVO"
            Option2.Caption = "Equipo existente en la Base de Datos"
        Case "30000"
            Label1.Caption = "DETALLE DE BIENES (Insumos)"
            Option1.Caption = "Insumos NUEVOS"
            Option2.Caption = "Insumos existentes en la Base de Datos"
        Case "34800"
            Label1.Caption = "DETALLE DE BIENES (Herramientas)"
            Option1.Caption = "Herramientas NUEVAS"
            Option2.Caption = "Herramientas existentes en la Base de Datos"
        Case "39800"
            Label1.Caption = "DETALLE DE BIENES (Repuestos)"
            Option1.Caption = "Repuestos NUEVOS"
            Option2.Caption = "Repuestos existentes en la Base de Datos"
        Case "24300"
            Label1.Caption = "SERVICIOS TECNICOS EXTERNOS"
            Option1.Caption = "Servicios NUEVOS"
            Option2.Caption = "Servicios existentes en la Base de Datos"
        Case "21000"
            Label1.Caption = "SERVICIOS BASICOS"
            Option1.Caption = "Servicios NUEVOS"
            Option2.Caption = "Servicios existentes en la Base de Datos"
     End Select
'    If lbl_det = "43340" Then
'        Label1.Caption = "DETALLE DE BIENES (Equipos)"
'        Option1.Caption = "Equipo NUEVO"
'        Option2.Caption = "Equipo existente en la Base de Datos"
'    Else
'        Label1.Caption = "DETALLE DE BIENES (Insumos)"
'        Option1.Caption = "Insumos NUEVOS"
'        Option2.Caption = "Insumos existentes en la Base de Datos"
'    End If
End Sub

Private Sub Form_Load()
    'Call ABRIR_TABLA
    mbDataChanged = False
'    Frame1.Visible = True
'    Frame2.Visible = False
    var_val2 = "2"
''    If swnuevo = 2 Then
''        dtc_desc2.BoundText = dtc_codigo2.BoundText
''        dtc_desc3.BoundText = dtc_codigo3.BoundText
''    End If
'    If mw_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado") = "S" Then
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' and edif_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto")
'        Image1 = Img_Foto
'    End If
'    If mw_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado") = "S" Then
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' edif_codigo = '" & mw_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto1")
'        Image2 = Img_Foto
'    End If
''    mw_solicitud.Ado_detalle1.Recordset("ges_gestion").Value = Year(Date)
''        mw_solicitud.Ado_detalle1.Recordset("unidad_codigo").Value = txt_campo1.Caption
''        mw_solicitud.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
''        mw_solicitud.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
''        mw_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado").Value = "N"
''        mw_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado").Value = "N"
''        mw_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
End Sub

Private Sub ABRIR_TABLA()
    'ac_bienes
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    If lbl_det = "43340" Then
        'lbl_det = "par_codigo" + " = " + "'43340'"
        rs_datos1.Open "select * from ac_bienes where (par_codigo = '43340' OR par_codigo = '99990') AND edif_codigo = '" & tw_identificacion_cliente.dtc_codigo3 & "' ", db, adOpenStatic   'order by descripcion
    End If
    'If Txt_campo1.Caption = "DNMAN" Then
    If lbl_det = "30000" Then
            rs_datos1.Open "select * from ac_bienes where par_codigo = '34110' OR par_codigo = '33100' OR par_codigo= '22210' or par_codigo= '24300' ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    'If Txt_campo1.Caption = "DNREP" Then
    If lbl_det = "39800" Then
        rs_datos1.Open "select * from ac_bienes where (par_codigo= '39810' or par_codigo= '39820') ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion 'and estado_codigo = 'APR'
    End If
    If lbl_det = "24300" Then
        '    rs_datos1.Open "select * from ac_bienes where par_codigo= '24300'  ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
        rs_datos1.Open "select * from ac_bienes where par_codigo= '24300'  ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    If lbl_det = "40000" Then
        '    rs_datos1.Open "select * from ac_bienes where grupo_codigo = '30000' OR par_codigo= '22210' OR (grupo_codigo = '40000' and par_codigo <> '43340') ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
        rs_datos1.Open "select * from ac_bienes where (grupo_codigo = '40000' and par_codigo <> '43340') ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    If lbl_det = "21000" Then
            'rs_datos1.Open "select * from ac_bienes where subgrupo_codigo= '21000' or subgrupo_codigo= '22000'  ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
            rs_datos1.Open "select * from ac_bienes where subgrupo_codigo= '21000' ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    'End If
        'lbl_det = " par_codigo = '" & 34110 & "' AND par_codigo = '" & 33100 & "' "
    
    'rs_datos1.Open "select * from ac_bienes where par_codigo = '" & 43340 & "' AND edif_codigo = '" & lbl_edif.Caption & "' ", db, adOpenStatic   'order by descripcion
    Set Ado_datos1.Recordset = rs_datos1
    If swnuevo = 2 Then
        dtc_codigo1.Text = Txt_campo5.Text
    End If
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    
    'ac_bienes_unidad_medida
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If lbl_det = "43340" Or lbl_det = "24300" Or lbl_det = "21000" Then
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo = 'T' order by unimed_descripcion ", db, adOpenStatic
    Else
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' order by unimed_descripcion ", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_datos2
    If swnuevo = 2 Then
        dtc_codigo2.Text = Txt_campo14.Text
    End If
    dtc_desc2.BoundText = dtc_codigo2.Text
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    If lbl_det = "24300" Then
        rs_datos3.Open "Select * from gc_tipo_solicitud WHERE solicitud_tipo = 7 OR solicitud_tipo = 10 OR solicitud_tipo = 4 OR solicitud_tipo = 5 OR solicitud_tipo = 8 order by solicitud_tipo_descripcion ", db, adOpenStatic
    Else
        rs_datos3.Open "Select * from gc_tipo_solicitud  order by solicitud_tipo_descripcion ", db, adOpenStatic
    End If
    Set Ado_datos3.Recordset = rs_datos3
    If swnuevo = 2 Then
        dtc_codigo5.Text = Txt_campo15.Text
    End If
    dtc_desc5.BoundText = dtc_codigo5.BoundText
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

Private Sub Option1_Click()
'    Frame2.Visible = True
'    Frame1.Visible = False
    var_val2 = "1"
End Sub

Private Sub Option2_Click()
'    Frame1.Visible = True
'    Frame2.Visible = False
    var_val2 = "2"
End Sub

Private Sub Txt_campo10_Change()
If Txt_campo16.Text = "" Then
        Txt_campo16.Text = "1"
        
    End If
    If Txt_campo10.Text <> "" Then
    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
    End If
End Sub

Private Sub Txt_campo10_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Or (KeyAscii = 8) Then     '(KeyAscii = 8) Or
    'MsgBox "ERROR ..."
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub Txt_campo16_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Or (KeyAscii = 8) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub Txt_campo16_Change()
  If Txt_campo10.Text = "" Then
        Txt_campo10.Text = "1"
    End If
    If Txt_campo16.Text <> "" Then
    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
    End If
End Sub

Private Sub Txt_campo18_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo18.BoundText
    dtc_desc1.BoundText = Txt_campo18.BoundText
    dtc_aux2.BoundText = Txt_campo18.BoundText
    dtc_aux3.BoundText = Txt_campo18.BoundText
    Txt_campo2.BoundText = Txt_campo18.BoundText
    dtc_aux1.BoundText = Txt_campo18.BoundText
    Txt_campo4.BoundText = Txt_campo18.BoundText
    Txt_campo3.BoundText = Txt_campo18.BoundText
End Sub

'Private Sub Txt_campo2_Click()
'    Call dtc_desc1_LostFocus
'End Sub

'Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

'Private Sub Txt_campo4_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub Txt_campo2_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo2.BoundText
    dtc_desc1.BoundText = Txt_campo2.BoundText
    dtc_aux2.BoundText = Txt_campo2.BoundText
    dtc_aux3.BoundText = Txt_campo2.BoundText
    dtc_aux1.BoundText = Txt_campo2.BoundText
    Txt_campo3.BoundText = Txt_campo2.BoundText
    Txt_campo4.BoundText = Txt_campo2.BoundText
    Txt_campo18.BoundText = Txt_campo2.BoundText
End Sub

Private Sub Txt_campo3_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo3.BoundText
    dtc_desc1.BoundText = Txt_campo3.BoundText
    dtc_aux2.BoundText = Txt_campo3.BoundText
    dtc_aux3.BoundText = Txt_campo3.BoundText
    Txt_campo2.BoundText = Txt_campo3.BoundText
    dtc_aux1.BoundText = Txt_campo3.BoundText
    Txt_campo4.BoundText = Txt_campo3.BoundText
    Txt_campo18.BoundText = Txt_campo3.BoundText
End Sub

Private Sub Txt_campo4_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo4.BoundText
    dtc_desc1.BoundText = Txt_campo4.BoundText
    dtc_aux2.BoundText = Txt_campo4.BoundText
    dtc_aux3.BoundText = Txt_campo4.BoundText
    Txt_campo2.BoundText = Txt_campo4.BoundText
    dtc_aux1.BoundText = Txt_campo4.BoundText
    Txt_campo3.BoundText = Txt_campo4.BoundText
    Txt_campo18.BoundText = Txt_campo4.BoundText
End Sub

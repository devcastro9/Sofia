VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_solicitud_bienes2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cotizaci�n - Detalle de Servicios"
   ClientHeight    =   7005
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_solicitud_bienes2.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   68
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3000
         MaskColor       =   &H00000000&
         Picture         =   "frm_solicitud_bienes2.frx":6C032
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_solicitud_bienes2.frx":6C23C
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   120
         Width           =   765
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
         Left            =   5655
         TabIndex        =   71
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   5655
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   10695
      Begin VB.Frame Frame3 
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
         Height          =   1455
         Left            =   120
         TabIndex        =   55
         Top             =   4080
         Width           =   10455
         Begin VB.TextBox Txt_campo10 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_venta_base"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   240
            TabIndex        =   62
            Text            =   "1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo16 
            Alignment       =   2  'Center
            DataField       =   "bien_cantidad"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   5400
            TabIndex        =   61
            Text            =   "1"
            Top             =   840
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "frm_solicitud_bienes2.frx":6C446
            DataField       =   "unimed_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   2520
            TabIndex        =   65
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unimed_descripcion"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "frm_solicitud_bienes2.frx":6C45F
            DataField       =   "unimed_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   1800
            TabIndex        =   66
            Top             =   840
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483637
            ForeColor       =   0
            ListField       =   "unimed_codigo"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "REG"
            DataField       =   "estado_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
            Height          =   300
            Left            =   9120
            TabIndex        =   64
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Txt_campo11 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "bien_total_venta"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
            TabIndex        =   63
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFC0&
            Height          =   480
            Index           =   2
            Left            =   9120
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbl_campo10 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lbl_desc2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   2520
            TabIndex        =   58
            Top             =   360
            Width           =   2370
         End
         Begin VB.Label lbl_campo16 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   5280
            TabIndex        =   57
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lbl_campo11 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
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
            ForeColor       =   &H00FFFF80&
            Height          =   360
            Left            =   7200
            TabIndex        =   56
            Top             =   360
            Width           =   1455
         End
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
         Height          =   2775
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Visible         =   0   'False
         Width           =   10455
         Begin VB.TextBox Txt_campo9 
            DataField       =   "modelo_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   6360
            TabIndex        =   8
            Text            =   "0"
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox Txt_campo8 
            DataField       =   "marca_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Text            =   "0"
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox Txt_campo7 
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   240
            TabIndex        =   6
            Text            =   "0"
            Top             =   1440
            Width           =   9975
         End
         Begin VB.TextBox Txt_campo6 
            DataField       =   "bien_descripcion"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   2880
            TabIndex        =   5
            Text            =   "0"
            Top             =   640
            Width           =   7335
         End
         Begin VB.TextBox Txt_campo5 
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Text            =   "0"
            Top             =   640
            Width           =   2415
         End
         Begin VB.Label lbl_codigo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "C�digo"
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
            TabIndex        =   50
            Top             =   375
            Width           =   660
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
            TabIndex        =   49
            Top             =   360
            Width           =   1080
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
            TabIndex        =   48
            Top             =   1140
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nacional o Extranjero"
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
            Left            =   6360
            TabIndex        =   47
            Top             =   1875
            Width           =   1920
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Tipo de Servicio"
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
            TabIndex        =   46
            Top             =   1875
            Width           =   1485
         End
      End
      Begin VB.TextBox Txt_campo14 
         DataField       =   "fosa_dimension_frente"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   285
         Left            =   3480
         TabIndex        =   31
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt_campo15 
         DataField       =   "fosa_dimension_fondo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   285
         Left            =   5565
         TabIndex        =   30
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000040C0&
         Caption         =   "Servicios NUEVOS"
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
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   4935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H000040C0&
         Caption         =   "Servicio existente en la Base de Datos"
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
         Left            =   5280
         TabIndex        =   1
         Top             =   1080
         Width           =   5055
      End
      Begin VB.TextBox Txt_campo17 
         DataField       =   "venta_o_compra"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   285
         Left            =   5640
         TabIndex        =   29
         Text            =   "V"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame1 
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
         Height          =   2775
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   10455
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2865
            TabIndex        =   54
            Top             =   2115
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   6345
            TabIndex        =   40
            Top             =   2115
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   9945
            TabIndex        =   39
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9840
            TabIndex        =   36
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "frm_solicitud_bienes2.frx":6C478
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   7840
            TabIndex        =   3
            Top             =   520
            Width           =   2380
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "frm_solicitud_bienes2.frx":6C491
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   525
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Txt_campo4 
            Bindings        =   "frm_solicitud_bienes2.frx":6C4AA
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Top             =   1300
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "bien_descripcion_anterior"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            Bindings        =   "frm_solicitud_bienes2.frx":6C4C3
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   3360
            TabIndex        =   41
            Top             =   2100
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "marca_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo3 
            Bindings        =   "frm_solicitud_bienes2.frx":6C4DC
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   6840
            TabIndex        =   42
            Top             =   2100
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "modelo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo18 
            Bindings        =   "frm_solicitud_bienes2.frx":6C4F5
            DataField       =   "bien_codigo"
            DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   51
            Top             =   2100
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ListField       =   "pais_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pais Origen"
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
            TabIndex        =   52
            Top             =   1815
            Width           =   1050
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Tipo de servicio"
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
            TabIndex        =   44
            Top             =   1815
            Width           =   1455
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Otros"
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
            Left            =   6840
            TabIndex        =   43
            Top             =   1815
            Width           =   480
         End
         Begin VB.Label Label3 
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
            TabIndex        =   37
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label lbl_descripcion 
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
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl_codigo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "C�digo"
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
            Left            =   7920
            TabIndex        =   34
            Top             =   240
            Width           =   660
         End
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "frm_solicitud_bienes2.frx":6C50E
         DataField       =   "bien_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   315
         Left            =   8400
         TabIndex        =   27
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSDataListLib.DataCombo dtc_aux2 
         Bindings        =   "frm_solicitud_bienes2.frx":6C527
         DataField       =   "bien_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   315
         Left            =   7200
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         Bindings        =   "frm_solicitud_bienes2.frx":6C540
         DataField       =   "bien_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
         Height          =   315
         Left            =   5400
         TabIndex        =   23
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
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "par_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
         Left            =   7080
         TabIndex        =   67
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lbl_edif 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "edif_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
         Left            =   9120
         TabIndex        =   53
         Top             =   120
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_descripcion 
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
         ForeColor       =   &H00FFFF80&
         Height          =   300
         Left            =   3720
         TabIndex        =   22
         Top             =   480
         Width           =   6615
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "frm_ao_compra_servicio.ado_detalle2"
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
         Height          =   300
         Left            =   1080
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
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
         ForeColor       =   &H00FFFFC0&
         Height          =   480
         Index           =   8
         Left            =   2800
         TabIndex        =   19
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lbl_codigo 
         BackColor       =   &H00000000&
         Caption         =   "Codigo Tr�mite"
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
         Height          =   480
         Left            =   320
         TabIndex        =   18
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
      ScaleWidth      =   10935
      TabIndex        =   11
      Top             =   7005
      Width           =   10935
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   16
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   6480
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
      Top             =   6480
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
End
Attribute VB_Name = "frm_solicitud_bienes2"
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
   sino = MsgBox("Est� Seguro de CANCELAR la operaci�n ? ", vbYesNo + vbQuestion, "Atenci�n")
   If sino = vbYes Then
        mw_solicitud.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     If lbl_det.Caption = "30000" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle3.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle3.Recordset("unidad_codigo").Value = txt_campo1.Caption
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
            tw_identificacion_cliente.Ado_detalle3.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "99900", dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
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
            
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_frente").Value = 0    'Txt_campo14.Text
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle3.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle3.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle3.Recordset.UpdateBatch adAffectAll
     End If
     If lbl_det.Caption = "TRANS" Or lbl_det.Caption = "ADUAN" Or lbl_det.Caption = "DESCA" Then
'   compra_codigo, compra_codigo_det, , , '                      , , , compra_concepto, tipo_descuento,
'                      almacen_codigo , usr_usuario, fecha_registro, hora_registro

         If swnuevo = 1 Then
            fw_compras_gral.Ado_detalle2.Recordset("ges_gestion").Value = glGestion
            'OBS
            fw_compras_gral.Ado_detalle2.Recordset("unidad_codigo").Value = txt_campo1.Caption
            fw_compras_gral.Ado_detalle2.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            fw_compras_gral.Ado_detalle2.Recordset("estado_codigo").Value = "REG"
            fw_compras_gral.Ado_detalle2.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                fw_compras_gral.Ado_detalle2.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                fw_compras_gral.Ado_detalle2.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
'                fw_compras_gral.Ado_detalle2.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
'                fw_compras_gral.Ado_detalle2.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            End If
            fw_compras_gral.Ado_detalle2.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            fw_compras_gral.Ado_detalle2.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "99900", dtc_aux2.Text)
            fw_compras_gral.Ado_detalle2.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            fw_compras_gral.Ado_detalle2.Recordset("compra_precio_unitario_bs").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
'            fw_compras_gral.Ado_detalle2.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            fw_compras_gral.Ado_detalle2.Recordset("compra_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            fw_compras_gral.Ado_detalle2.Recordset("compra_precio_total_bs").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
            fw_compras_gral.Ado_detalle2.Recordset("compra_precio_unitario_dol").Value = 0
            fw_compras_gral.Ado_detalle2.Recordset("compra_precio_total_dol").Value = 0
            
            fw_compras_gral.Ado_detalle2.Recordset("compra_descuento_bs").Value = IIf(Txt_campo10 = "", 0, CDbl(Txt_campo10) * 0.87)
            fw_compras_gral.Ado_detalle2.Recordset("compra_descuento_dol").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            fw_compras_gral.Ado_detalle2.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            fw_compras_gral.Ado_detalle2.Recordset("usr_codigo").Value = glusuario
            fw_compras_gral.Ado_detalle2.Recordset.UpdateBatch adAffectAll
    End If
    If lbl_det.Caption = "39800" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle6.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle6.Recordset("unidad_codigo").Value = txt_campo1.Caption
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
            
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_frente").Value = 0    'Txt_campo14.Text
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle6.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle6.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle6.Recordset.UpdateBatch adAffectAll
    End If
'     Set rs_aux1 = New ADODB.Recordset
'     SQL_FOR = "select * from ao_solicitud_edificacion where unidad_codigo = '" & mw_solicitud.Ado_datos.Recordset("unidad_codigo") & "' and solicitud_codigo = " & mw_solicitud.Ado_datos.Recordset("solicitud_codigo") & " and edif_codigo = '" & dtc_codigo1.Text & "'  "
'     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'     If rs_aux1.RecordCount > 0 Then
'        MsgBox "El c�digo ya existe, consulte con el administrador del Sistema..."
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
        SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    ", db, adOpenStatic
            If Not rs_aux5.EOF Then
                 var_cod5 = IIf(IsNull(rs_aux5!Codigo), 1, rs_aux5!Codigo + 1)
            End If
            rs_aux4.AddNew
            rs_aux4!ges_gestion = Year(Date)
            rs_aux4!unidad_codigo = txt_campo1.Caption
            rs_aux4!solicitud_codigo = txt_codigo.Caption
            rs_aux4!edif_codigo = tw_identificacion_cliente.dtc_codigo3.Text
            rs_aux4!trafico_codigo = "0"  'Ado_datos.Recordset!trafico_codigo
            rs_aux4!cotiza_codigo = var_cod5
            'Call correl_bien
            'VAR_COD3 = "36NO-" + Trim(Str(VAR_COD2))
            rs_aux4!bien_codigo = "MAN-002"       'VAR_COD3  '"36NO-" + Trim(Str(VAR_COD2))
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
            rs_aux4!proceso_codigo = "TEC"
            rs_aux4!subproceso_codigo = "TEC-2"
            rs_aux4!etapa_codigo = "TEC-2-1"
            rs_aux4!poa_codigo = "3.2.3"
            rs_aux4!clasif_codigo = "TEC"
            rs_aux4!doc_codigo = "R-362"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
            rs_aux4!doc_numero = "0"
            
            
            rs_aux4!estado_codigo = "REG"
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update
            db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod5 & " Where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
      End If
    End If
     Frame1.Visible = False
     Frame2.Visible = False

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
    MsgBox "Debe registrar el " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'     Txt_campo12.Caption = var_itm
'     Txt_campo13.Caption = var_ctm
  
'        If Txt_campo2.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo3.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo4.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo5.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo6.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo7.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo8.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo9.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo10.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'          VAR_VAL = "ERR"
'          Exit Sub
'        End If
'        If Txt_campo11.Text = "" Then
'          MsgBox "Debe registrar : " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
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
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atenci�n")
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
'        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atenci�n"
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

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
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
End Sub

Private Sub dtc_desc1_LostFocus()
'    Select Case dtc_aux2.Text
'        Case "DPTO"
'            lbl_campo8.Caption = "Depto.de 1 Dorm."
'            lbl_campo7.Caption = "N�Habit.Servicio"
'        Case "OFIG"
'            lbl_campo3.Caption = "�rea Pasillos"
'        Case "OFIU"
'            lbl_campo3.Caption = "�rea Pasillos"
'        Case "COMR"
'            lbl_campo3.Caption = "�rea Pasillos"
'        Case "EDUC"
'            lbl_campo2.Caption = "�rea Aulas"
'            lbl_campo3.Caption = "�rea Admin."
'        Case "HOTL"
'            lbl_campo8.Caption = "N�Dormitorios"
'        Case "REST"
'            lbl_campo3.Caption = "�rea Comedor"
'        Case "HOSP"
'            lbl_campo8.Caption = "N� de Camas"
'        Case "HOSs"
'            lbl_campo8.Caption = "N� de Camas"
'        Case "GARJ"
'            lbl_campo8.Caption = "N�de Parqueos"
'     End Select
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
    var_val2 = "2"
    Select Case lbl_det
        Case "PROVI"
            Label1.Caption = "DETALLE DE BIENES (Equipos)"
            Option1.Caption = "Equipo NUEVO"
            Option2.Caption = "Equipo existente en la Base de Datos"
        Case "TRANS"
            Label1.Caption = "DETALLE DE SERVICIOS (Transporte)"
            Option1.Caption = "Servicios NUEVOS"
            Option2.Caption = "Servicios existentes en la Base de Datos"
        Case "ADUAN"
            Label1.Caption = "DETALLE DE SERVICIOS (Aduana)"
            Option1.Caption = "Servicios NUEVOS"
            Option2.Caption = "Servicios existentes en la Base de Datos"
        Case "DESCA"
            Label1.Caption = "DETALLE DE SERVICIOS (Descarguio)"
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
    Frame1.Visible = True
    Frame2.Visible = False
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
    Select Case lbl_det
        Case "PROVI"    'PROVISION DE EQUIPOS
            rs_datos1.Open "select * from ac_bienes where par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "TRANS"    'TRANSPORTE
            rs_datos1.Open "select * from ac_bienes where par_codigo = '22300' and  bien_codigo_anterior = 'TRANSP' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "ADUAN"    'DESADUANIZACION
            rs_datos1.Open "select * from ac_bienes where par_codigo = '22300' and  bien_codigo_anterior = 'ADUANA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "DESCA"    'DESCARGUIO Y OTROS
            rs_datos1.Open "select * from ac_bienes where par_codigo = '22300' and  bien_codigo_anterior = 'DESCAR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    End Select
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_aux2.BoundText = dtc_codigo1.BoundText
    
    'ac_bienes_unidad_medida
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
'    If lbl_det = "43340" Then
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo = 'T' ", db, adOpenStatic   'order by descripcion
'    Else
'        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' ", db, adOpenStatic   'order by descripcion
'    End If
    
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
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
    Frame2.Visible = True
    Frame1.Visible = False
    var_val2 = "1"
End Sub

Private Sub Option2_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    var_val2 = "2"
End Sub

Private Sub Txt_campo10_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub Txt_campo16_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub Txt_campo16_Change()
    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
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

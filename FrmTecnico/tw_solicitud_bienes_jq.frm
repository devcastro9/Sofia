VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tw_solicitud_bienes_jq 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cotización - Detalle de Bienes"
   ClientHeight    =   6735
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   10635
      TabIndex        =   20
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1680
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   120
         Width           =   1340
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE BIENES (Equipos)"
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
         Left            =   5145
         TabIndex        =   21
         Top             =   240
         Width           =   4785
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   5415
      Left            =   120
      TabIndex        =   10
      Top             =   1200
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
         Height          =   1455
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   10455
         Begin VB.TextBox Txt_campo19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "bien_cantidad_por_empaque"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7080
            TabIndex        =   53
            Text            =   "1"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox Txt_campo10 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_venta_base"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   285
            Left            =   240
            TabIndex        =   47
            Text            =   "1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "bien_cantidad"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5400
            TabIndex        =   46
            Text            =   "1"
            Top             =   840
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            DataField       =   "unimed_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   2520
            TabIndex        =   50
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
            DataField       =   "unimed_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   1800
            TabIndex        =   51
            Top             =   840
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unimed_codigo"
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
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
            TabIndex        =   49
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Txt_campo11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "bien_total_venta"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
            Left            =   8760
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nro. de Horas por cada Visita"
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
            Left            =   6960
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lbl_campo10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Costo Unitario Bs. por Periodo del Servicio Bs."
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
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label lbl_desc2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Periodicidad del Servicio (Unidad de Medida)"
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
            Left            =   2520
            TabIndex        =   43
            Top             =   240
            Width           =   2370
         End
         Begin VB.Label lbl_campo16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nro.de Periodos (por Gestión)"
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
            Left            =   5280
            TabIndex        =   42
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lbl_campo11 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Estimado por Gestión Bs."
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
            Left            =   8760
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox Txt_campo14 
         DataField       =   "unimed_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   3480
         TabIndex        =   24
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt_campo15 
         DataField       =   "fosa_dimension_frente"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   5565
         TabIndex        =   23
         Text            =   "0"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt_campo17 
         DataField       =   "venta_o_compra"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   5640
         TabIndex        =   22
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
         Height          =   2775
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   10455
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2865
            Locked          =   -1  'True
            TabIndex        =   39
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
            Locked          =   -1  'True
            TabIndex        =   31
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
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1320
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   7800
            TabIndex        =   3
            Top             =   520
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   520
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Txt_campo4 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   1300
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_descripcion_anterior"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   3360
            TabIndex        =   32
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
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   6840
            TabIndex        =   33
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
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   315
            Left            =   240
            TabIndex        =   36
            Top             =   2100
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "pais_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
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
            TabIndex        =   37
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   1815
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ubicacacion Fisica / Caracteristicas"
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
            TabIndex        =   28
            Top             =   1020
            Width           =   3210
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
            TabIndex        =   26
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
            Left            =   7800
            TabIndex        =   25
            Top             =   240
            Width           =   660
         End
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         DataField       =   "bien_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
         Height          =   315
         Left            =   8400
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "par_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux2 
         DataField       =   "bien_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
         Height          =   315
         Left            =   7200
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "subgrupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         DataField       =   "bien_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "grupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipo existente en la Base de Datos"
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
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipo NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lbl_det 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "par_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
         TabIndex        =   52
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
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
         TabIndex        =   13
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
         TabIndex        =   15
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
         DataSource      =   "tw_identificacion_cliente.ado_detalle2"
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   11
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
      TabIndex        =   4
      Top             =   6735
      Width           =   10935
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   9
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
      Left            =   240
      TabIndex        =   54
      Top             =   3840
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox Txt_campo5 
         DataField       =   "bien_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   240
         TabIndex        =   59
         Text            =   "0"
         Top             =   640
         Width           =   2415
      End
      Begin VB.TextBox Txt_campo6 
         DataField       =   "bien_descripcion"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   2880
         TabIndex        =   58
         Text            =   "0"
         Top             =   640
         Width           =   7335
      End
      Begin VB.TextBox Txt_campo7 
         Height          =   285
         Left            =   240
         TabIndex        =   57
         Text            =   "0"
         Top             =   1440
         Width           =   9975
      End
      Begin VB.TextBox Txt_campo8 
         DataField       =   "marca_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   240
         TabIndex        =   56
         Text            =   "0"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox Txt_campo9 
         DataField       =   "modelo_codigo"
         DataSource      =   "frm_to_identificacion_cliente.ado_detalle2"
         Height          =   285
         Left            =   6360
         TabIndex        =   55
         Text            =   "0"
         Top             =   2160
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
         TabIndex        =   64
         Top             =   1875
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
         Left            =   6360
         TabIndex        =   63
         Top             =   1875
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Ubicacacion Fisica / Caracteristicas"
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
         TabIndex        =   62
         Top             =   1140
         Width           =   3210
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
         TabIndex        =   61
         Top             =   360
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
         TabIndex        =   60
         Top             =   375
         Width           =   660
      End
   End
End
Attribute VB_Name = "tw_solicitud_bienes_jq"
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
Dim rs_aux6 As New ADODB.Recordset
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
        aw_p_ao_solicitud.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     If swnuevo = 1 Then
       ' tw_identificacion_cliente.Ado_detalle2.Recordset("ges_gestion").Value = glGestion
        tw_identificacion_cliente.Ado_detalle2.Recordset("ges_gestion").Value = tw_identificacion_cliente.Ado_datos.Recordset("ges_gestion")
        tw_identificacion_cliente.Ado_detalle2.Recordset("unidad_codigo").Value = txt_campo1.Caption
        tw_identificacion_cliente.Ado_detalle2.Recordset("solicitud_codigo").Value = txt_codigo.Caption
        tw_identificacion_cliente.Ado_detalle2.Recordset("estado_codigo").Value = "REG"
        tw_identificacion_cliente.Ado_detalle2.Recordset("venta_o_compra").Value = "V"
'        tw_identificacion_cliente.Ado_detalle2.Recordset("archivo_foto_cargado").Value = "N"
'        tw_identificacion_cliente.Ado_detalle2.Recordset("archivo_plano_cargado").Value = "N"
     End If
        If var_val2 = "1" Then
            tw_identificacion_cliente.Ado_detalle2.Recordset("bien_codigo").Value = IIf(Txt_campo5.Text = "", "NA1", Txt_campo5.Text)
            tw_identificacion_cliente.Ado_detalle2.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
            tw_identificacion_cliente.Ado_detalle2.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
        Else
            tw_identificacion_cliente.Ado_detalle2.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
            tw_identificacion_cliente.Ado_detalle2.Recordset("marca_codigo").Value = IIf(Txt_campo2.Text = "", "S/M", Txt_campo2.Text)
            tw_identificacion_cliente.Ado_detalle2.Recordset("modelo_codigo").Value = IIf(Txt_campo3.Text = "", "S/M", Txt_campo3.Text)
            'OJO FALTA GRABAR EN ac_bienes
        End If
        tw_identificacion_cliente.Ado_detalle2.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
        tw_identificacion_cliente.Ado_detalle2.Recordset("subgrupo_codigo").Value = IIf(Dtc_aux2.Text = "", "99900", Dtc_aux2.Text)
        tw_identificacion_cliente.Ado_detalle2.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
        If Txt_campo16.Text <> "" Then
            Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
        Else
            Txt_campo16.Text = 1
            Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
        End If
        tw_identificacion_cliente.Ado_detalle2.Recordset("bien_precio_venta_base").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
        tw_identificacion_cliente.Ado_detalle2.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
        tw_identificacion_cliente.Ado_detalle2.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
        tw_identificacion_cliente.Ado_detalle2.Recordset("bien_total_venta").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
        tw_identificacion_cliente.Ado_detalle2.Recordset("bien_cantidad_por_empaque").Value = IIf(Txt_campo19 = "", 2, Txt_campo19)
        'tw_identificacion_cliente.Ado_detalle3.Recordset("bien_total_compra").Value = 0    '
        Select Case txt_campo1.Caption
           Case "DNMAN"
               tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_frente").Value = "10"
           Case "DNREP"
               tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_frente").Value = "7"
           Case "DNINS"
               tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_frente").Value = "4"
           Case "DNAJS"
               tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_frente").Value = "5"
           Case "DNMOD"
               tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_frente").Value = "9"
        End Select
        tw_identificacion_cliente.Ado_detalle2.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
        'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
        'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
        tw_identificacion_cliente.Ado_detalle2.Recordset("fecha_registro").Value = Date
        'aw_p_ao_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
        tw_identificacion_cliente.Ado_detalle2.Recordset("usr_codigo").Value = glusuario
        tw_identificacion_cliente.Ado_detalle2.Recordset.UpdateBatch adAffectAll
     
'     Set rs_aux1 = New ADODB.Recordset
'     SQL_FOR = "select * from ao_solicitud_edificacion where unidad_codigo = '" & aw_p_ao_solicitud.Ado_datos.Recordset("unidad_codigo") & "' and solicitud_codigo = " & aw_p_ao_solicitud.Ado_datos.Recordset("solicitud_codigo") & " and edif_codigo = '" & dtc_codigo1.Text & "'  "
'     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'     If rs_aux1.RecordCount > 0 Then
'        MsgBox "El código ya existe, consulte con el administrador del Sistema..."
'        var_cod = 0
'        Exit Sub
'     Else
'        aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
'     End If
     
     
'     var_cod = aw_p_ao_solicitud.Ado_detalle1.Recordset.RecordCount
'     db.Execute "Update ao_solicitud Set correl_edificacion = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "  "
    If lbl_det = "43340" Then
     'Graba en Cotiza    1
        Set rs_aux4 = New ADODB.Recordset
        If rs_aux4.State = 1 Then rs_aux4.Close
        SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " AND bien_codigo = '" & dtc_codigo1.Text & "'    "
        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If swnuevo = 1 Then
            'Call cotiza_codigo
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
            rs_aux4!pais_continente = "NN"
            'Call correl_bien
            rs_aux4!bien_codigo = IIf(dtc_codigo1.Text = "", Txt_campo5.Text, dtc_codigo1.Text) '"MAN-002"       'CODIGO Servicio de Mantenimeitno
            rs_aux4!proceso_codigo = "TEC"
            rs_aux4!subproceso_codigo = "TEC-02"
            rs_aux4!etapa_codigo = "TEC-02-01"
            rs_aux4!poa_codigo = "3.2.3"
            rs_aux4!clasif_codigo = "TEC"
            rs_aux4!doc_codigo = "R-362"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
            rs_aux4!doc_numero = "0"
            rs_aux4!estado_codigo = "APR"
        
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
            rs_aux4!fecha_registro = Date
            rs_aux4!usr_codigo = glusuario
            rs_aux4.Update
        Else
            db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_cantidad = " & CDbl(Txt_campo16) & ", cotiza_precio_fob_bs = " & CDbl(Txt_campo10.Text) & ", cotiza_precio_total_bs = " & CDbl(Txt_campo11.Caption) & ", costo_monto = " & CDbl(Txt_campo11.Caption) & " Where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " and bien_codigo = '" & dtc_codigo1.Text & "'    "
        End If
        
        If swnuevo = 1 Then
            db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod5 & " Where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
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
    MsgBox "Debe registrar el " + lbl_codigo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo5.Text = "" Then
    MsgBox "Debe registrar el " + lbl_codigo2.Caption, vbCritical + vbExclamation, "Validación de datos"
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
'  If aw_p_ao_solicitud.Ado_detalle1.Recordset("estado_codigo") = "REG" Then
'    Dim ARCH_FOTO As String
'    Dim SW0 As String
'    If aw_p_ao_solicitud.Ado_detalle1.Recordset!archivo_foto_cargado = "N" Then
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
'        CodBien = aw_p_ao_solicitud.Ado_detalle1.Recordset!edif_codigo
'        'If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'        If Guardar_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = " & aw_p_ao_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & " and edif_codigo = '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
'            MsgBox "Se cargo la Imagen Correctamente !!"
'        Else
'            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'        End If
'    Else
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto")
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
    Dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
    Txt_campo2.BoundText = dtc_aux1.BoundText
    Txt_campo3.BoundText = dtc_aux1.BoundText
    Txt_campo4.BoundText = dtc_aux1.BoundText
    Txt_campo18.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo1.BoundText = Dtc_aux2.BoundText
    dtc_desc1.BoundText = Dtc_aux2.BoundText
    dtc_aux1.BoundText = Dtc_aux2.BoundText
    dtc_aux3.BoundText = Dtc_aux2.BoundText
    Txt_campo2.BoundText = Dtc_aux2.BoundText
    Txt_campo3.BoundText = Dtc_aux2.BoundText
    Txt_campo4.BoundText = Dtc_aux2.BoundText
    Txt_campo18.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux3.BoundText
    dtc_desc1.BoundText = dtc_aux3.BoundText
    Dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
    Txt_campo2.BoundText = dtc_aux3.BoundText
    Txt_campo3.BoundText = dtc_aux3.BoundText
    Txt_campo4.BoundText = dtc_aux3.BoundText
    Txt_campo18.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo1_LostFocus()
    'ao_solicitud_calculo_trafico
    'FALTA ESCALERAS Y MINICARGAS !!
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "  ", db, adOpenStatic      'order by descripcion
    If rs_aux6.RecordCount > 0 Then
        If rs_aux6!trafico_num_paradas < 9 Then
            Txt_campo19.Text = "2"
        Else
            Txt_campo19.Text = "4"
        End If
    End If
    'Set Ado_datos2.Recordset = rs_aux6
    'dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Dtc_aux2.BoundText = dtc_desc1.BoundText
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
'            lbl_campo7.Caption = "NºHabit.Servicio"
'        Case "OFIG"
'            lbl_campo3.Caption = "Área Pasillos"
'        Case "OFIU"
'            lbl_campo3.Caption = "Área Pasillos"
'        Case "COMR"
'            lbl_campo3.Caption = "Área Pasillos"
'        Case "EDUC"
'            lbl_campo2.Caption = "Área Aulas"
'            lbl_campo3.Caption = "Área Admin."
'        Case "HOTL"
'            lbl_campo8.Caption = "NºDormitorios"
'        Case "REST"
'            lbl_campo3.Caption = "Área Comedor"
'        Case "HOSP"
'            lbl_campo8.Caption = "Nº de Camas"
'        Case "HOSs"
'            lbl_campo8.Caption = "Nº de Camas"
'        Case "GARJ"
'            lbl_campo8.Caption = "Nºde Parqueos"
'     End Select
End Sub

Private Sub dtc_desc2_Change()
'    Select Case dtc_desc2.Text
'       Case "MENSUAL"
'           Txt_campo16.Text = "12"
'       Case "BIMESTRAL"
'           Txt_campo16.Text = "6"
'       Case "TRIMESTRAL"
'           Txt_campo16.Text = "4"
'       Case "CUATRIMESTRAL"
'           Txt_campo16.Text = "3"
'       Case "QMES"
'           Txt_campo16.Text = "2.5"
'        Case "SMES"
'           Txt_campo16.Text = "2"
'        Case "ANUAL"
'           Txt_campo16.Text = "1"
'    End Select
'    Txt_campo19.Text = "2"
'    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc2_LostFocus()
    Select Case dtc_codigo2.Text
       Case "MES"
           Txt_campo16.Text = "12"
       Case "BMES"
           Txt_campo16.Text = "6"
       Case "TMES"
           Txt_campo16.Text = "4"
       Case "CMES"
           Txt_campo16.Text = "3"
       Case "QMES"
           Txt_campo16.Text = "2.5"
        Case "SMES"
           Txt_campo16.Text = "2"
        Case "ANUAL"
           Txt_campo16.Text = "1"
    End Select
    Txt_campo19.Text = "2"
    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
End Sub

Private Sub Form_Activate()
    Call ABRIR_TABLA
    mbDataChanged = False
    var_val2 = "2"
    If lbl_det = "43340" Then
        Label1.Caption = "DETALLE DE BIENES (Equipos)"
        Option1.Caption = "Equipo NUEVO"
        Option2.Caption = "Equipo existente en la Base de Datos"
        If Txt_campo16.Text = "0" Or Txt_campo16.Text = "" Then
            Txt_campo16.Text = "12"
        End If
    Else
        Label1.Caption = "DETALLE DE BIENES (Insumos)"
        Option1.Caption = "Insumos NUEVOS"
        Option2.Caption = "Insumos existentes en la Base de Datos"
    End If
End Sub

Private Sub Form_Load()
    'Call ABRIR_TABLA
    mbDataChanged = False
    Frame1.Visible = True
'    Frame2.Visible = False
    var_val2 = "2"
''    If swnuevo = 2 Then
''        dtc_desc2.BoundText = dtc_codigo2.BoundText
''        dtc_desc3.BoundText = dtc_codigo3.BoundText
''    End If
'    If aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado") = "S" Then
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' and edif_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto")
'        Image1 = Img_Foto
'    End If
'    If aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado") = "S" Then
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_solicitud_edificacion Where unidad_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("unidad_codigo") & "' and solicitud_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("solicitud_codigo") & "' edif_codigo = '" & aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo") & "' ", "Foto1")
'        Image2 = Img_Foto
'    End If
''    aw_p_ao_solicitud.Ado_detalle1.Recordset("ges_gestion").Value = Year(Date)
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("unidad_codigo").Value = txt_campo1.Caption
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_foto_cargado").Value = "N"
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_plano_cargado").Value = "N"
''        aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
    'ac_bienes
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    If lbl_det = "43340" Then
        'lbl_det = "par_codigo" + " = " + "'43340'"
        rs_datos1.Open "select * from ac_bienes where (par_codigo = '43340' AND edif_codigo = '" & tw_identificacion_cliente.dtc_codigo3 & "') OR (par_codigo = '99990')  ", db, adOpenStatic   'order by descripcion
    Else
        rs_datos1.Open "select * from ac_bienes where par_codigo = '34110' OR par_codigo = '33100' OR par_codigo= '22210' OR par_codigo = '99990' ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
        'lbl_det = " par_codigo = '" & 34110 & "' AND par_codigo = '" & 33100 & "' "
    End If
    'rs_datos1.Open "select * from ac_bienes where par_codigo = '" & 43340 & "' AND edif_codigo = '" & lbl_edif.Caption & "' ", db, adOpenStatic   'order by descripcion
    Set Ado_datos1.Recordset = rs_datos1
    If swnuevo = 2 Then
        dtc_codigo1.Text = Txt_campo5.Text
    End If
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    
    'ac_bienes_unidad_medida
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If lbl_det = "43340" Then
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo = 'T' ", db, adOpenStatic   'order by descripcion
    Else
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' ", db, adOpenStatic   'order by descripcion
    End If
    If rs_datos2.RecordCount > 0 Then
    Set Ado_datos2.Recordset = rs_datos2
    If swnuevo = 2 Then
        dtc_codigo2.Text = Txt_campo14.Text
    End If
'    dtc_codigo2.Text = Ado_datos2.Recordset!unimed_codigo
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    End If
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

Private Sub Txt_campo10_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub Txt_campo10_LostFocus()
    Select Case dtc_codigo2.Text
       Case "MES"
           Txt_campo16.Text = "12"
       Case "BMES"
           Txt_campo16.Text = "6"
       Case "TMES"
           Txt_campo16.Text = "4"
       Case "CMES"
           Txt_campo16.Text = "3"
       Case "QMES"
           Txt_campo16.Text = "2.5"
        Case "SMES"
           Txt_campo16.Text = "2"
        Case "ANUAL"
           Txt_campo16.Text = "1"
    End Select
    Txt_campo19.Text = "2"
    Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
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
    If Txt_campo16.Text <> "" Then
        Txt_campo11.Caption = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
    End If
End Sub

Private Sub Txt_campo18_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo18.BoundText
    dtc_desc1.BoundText = Txt_campo18.BoundText
    Dtc_aux2.BoundText = Txt_campo18.BoundText
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
    Dtc_aux2.BoundText = Txt_campo2.BoundText
    dtc_aux3.BoundText = Txt_campo2.BoundText
    dtc_aux1.BoundText = Txt_campo2.BoundText
    Txt_campo3.BoundText = Txt_campo2.BoundText
    Txt_campo4.BoundText = Txt_campo2.BoundText
    Txt_campo18.BoundText = Txt_campo2.BoundText
End Sub

Private Sub Txt_campo3_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo3.BoundText
    dtc_desc1.BoundText = Txt_campo3.BoundText
    Dtc_aux2.BoundText = Txt_campo3.BoundText
    dtc_aux3.BoundText = Txt_campo3.BoundText
    Txt_campo2.BoundText = Txt_campo3.BoundText
    dtc_aux1.BoundText = Txt_campo3.BoundText
    Txt_campo4.BoundText = Txt_campo3.BoundText
    Txt_campo18.BoundText = Txt_campo3.BoundText
End Sub

Private Sub Txt_campo4_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo4.BoundText
    dtc_desc1.BoundText = Txt_campo4.BoundText
    Dtc_aux2.BoundText = Txt_campo4.BoundText
    dtc_aux3.BoundText = Txt_campo4.BoundText
    Txt_campo2.BoundText = Txt_campo4.BoundText
    dtc_aux1.BoundText = Txt_campo4.BoundText
    Txt_campo3.BoundText = Txt_campo4.BoundText
    Txt_campo18.BoundText = Txt_campo4.BoundText
End Sub

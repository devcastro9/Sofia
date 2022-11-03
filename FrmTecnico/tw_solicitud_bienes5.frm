VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tw_solicitud_bienes5 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Cotización - Detalle de Bienes"
   ClientHeight    =   8205
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   11235
      TabIndex        =   58
      Top             =   120
      Width           =   11295
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         Picture         =   "tw_solicitud_bienes5.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   70
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         Picture         =   "tw_solicitud_bienes5.frx":08EC
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   69
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE BIENES"
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
         Left            =   5925
         TabIndex        =   59
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   6975
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   11295
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
         Height          =   2655
         Left            =   120
         TabIndex        =   45
         Top             =   4215
         Width           =   11055
         Begin VB.CommandButton CmdCalcula2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Calcula Precio Unitario -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox Txt_campo22 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "bien_total_eur"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   7200
            TabIndex        =   74
            Text            =   "1"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo20 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "bien_total_compra"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   5520
            TabIndex        =   73
            Text            =   "1"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo11 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            DataField       =   "bien_total_venta"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   3840
            TabIndex        =   72
            Text            =   "1"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton CmdCalcula 
            BackColor       =   &H008080FF&
            Caption         =   "Calcula Precio Total  -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox Txt_campo21 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_eur"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   7200
            TabIndex        =   68
            Text            =   "1"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo19 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_compra"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   5520
            TabIndex        =   5
            Text            =   "1"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo10 
            Alignment       =   2  'Center
            DataField       =   "bien_precio_venta_base"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   3840
            TabIndex        =   4
            Text            =   "1"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo16 
            Alignment       =   2  'Center
            DataField       =   "bien_cantidad"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5520
            TabIndex        =   6
            Text            =   "1"
            Top             =   1440
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "tw_solicitud_bienes5.frx":10C2
            DataField       =   "tipo_moneda"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   4560
            TabIndex        =   81
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "tipo_moneda_descripcion"
            BoundColumn     =   "tipo_moneda"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "tw_solicitud_bienes5.frx":10DB
            DataField       =   "tipo_moneda"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   3840
            TabIndex        =   82
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "tipo_moneda"
            BoundColumn     =   "tipo_moneda"
            Text            =   ""
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   11040
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo de Moneda"
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
            Left            =   1920
            TabIndex        =   80
            Top             =   240
            Width           =   1650
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C00000&
            X1              =   3840
            X2              =   10695
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Euros"
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
            TabIndex        =   67
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dolares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   5640
            TabIndex        =   61
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bolivianos"
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
            Left            =   3960
            TabIndex        =   60
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label lbl_campo10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "<---      Precio Unitario"
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
            Left            =   8880
            TabIndex        =   48
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label lbl_campo16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "<---        X      Cantidad"
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
            Height          =   360
            Left            =   8880
            TabIndex        =   47
            Top             =   1440
            Width           =   1995
         End
         Begin VB.Label lbl_campo11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "<---      =  Precio Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   240
            Left            =   8880
            TabIndex        =   46
            Top             =   2160
            Width           =   1995
         End
      End
      Begin VB.TextBox Txt_campo14 
         DataField       =   "unimed_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
         Height          =   285
         Left            =   6000
         TabIndex        =   28
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Txt_campo15 
         DataField       =   "fosa_dimension_frente"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
         Height          =   285
         Left            =   7005
         TabIndex        =   27
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Txt_campo17 
         DataField       =   "venta_o_compra"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
         Height          =   285
         Left            =   8040
         TabIndex        =   26
         Text            =   "V"
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
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
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   11055
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   10545
            TabIndex        =   75
            Top             =   2055
            Width           =   255
         End
         Begin VB.TextBox Txt_campo9 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   3960
            TabIndex        =   57
            Text            =   "0"
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Txt_campo8 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   4920
            TabIndex        =   56
            Text            =   "0"
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Txt_campo6 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   7200
            TabIndex        =   55
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo7 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   5760
            TabIndex        =   54
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo5 
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   285
            Left            =   8760
            TabIndex        =   53
            Text            =   "0"
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtObservacion 
            DataField       =   "observacion"
            DataSource      =   "tw_identificacion_cliente.ado_detalle2"
            Height          =   525
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   2820
            Width           =   10575
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3105
            TabIndex        =   44
            Top             =   1275
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   6825
            TabIndex        =   36
            Top             =   1275
            Width           =   255
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   10545
            TabIndex        =   35
            Top             =   1275
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "tw_solicitud_bienes5.frx":10F4
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   8205
            TabIndex        =   1
            Top             =   525
            Width           =   2625
            _ExtentX        =   4630
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
            Bindings        =   "tw_solicitud_bienes5.frx":110D
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   9120
            TabIndex        =   34
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ListField       =   "unimed_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            Bindings        =   "tw_solicitud_bienes5.frx":1126
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   3600
            TabIndex        =   37
            Top             =   1260
            Width           =   3495
            _ExtentX        =   6165
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
            Bindings        =   "tw_solicitud_bienes5.frx":113F
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   7320
            TabIndex        =   38
            Top             =   1260
            Width           =   3495
            _ExtentX        =   6165
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
            Bindings        =   "tw_solicitud_bienes5.frx":1158
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   240
            TabIndex        =   41
            Top             =   1260
            Width           =   3135
            _ExtentX        =   5530
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
            Bindings        =   "tw_solicitud_bienes5.frx":1171
            DataField       =   "fosa_dimension_frente"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   2040
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "solicitud_tipo_descripcion"
            BoundColumn     =   "solicitud_tipo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "tw_solicitud_bienes5.frx":118A
            DataField       =   "fosa_dimension_frente"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   5880
            TabIndex        =   50
            Top             =   1680
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            ListField       =   "solicitud_tipo"
            BoundColumn     =   "solicitud_tipo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "tw_solicitud_bienes5.frx":11A3
            DataField       =   "bien_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   525
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux1 
            Bindings        =   "tw_solicitud_bienes5.frx":11BC
            DataField       =   "bien_codigo"
            DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   1680
            TabIndex        =   62
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
         Begin MSDataListLib.DataCombo dtc_aux2 
            Bindings        =   "tw_solicitud_bienes5.frx":11D5
            DataField       =   "bien_codigo"
            DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   3000
            TabIndex        =   63
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
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "tw_solicitud_bienes5.frx":11EE
            DataField       =   "bien_codigo"
            DataSource      =   "frm_to_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   4200
            TabIndex        =   64
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
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_solicitud_bienes5.frx":1207
            DataField       =   "unimed_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   8520
            TabIndex        =   76
            Top             =   2520
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "unimed_descripcion"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_solicitud_bienes5.frx":1220
            DataField       =   "unimed_codigo"
            DataSource      =   "tw_identificacion_cliente.ado_detalle5"
            Height          =   315
            Left            =   7800
            TabIndex        =   77
            Top             =   2520
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unimed_codigo"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin VB.Label lbl_desc2 
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
            Left            =   9120
            TabIndex        =   78
            Top             =   1755
            Width           =   1890
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Denominación del Repuesto - PARA EL CLIENTE"
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
            TabIndex        =   52
            Top             =   2520
            Width           =   4440
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "El Repuesto será utilizado para:"
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
            TabIndex        =   51
            Top             =   1755
            Width           =   2850
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
            TabIndex        =   42
            Top             =   975
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
            Left            =   3600
            TabIndex        =   40
            Top             =   975
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
            Left            =   7320
            TabIndex        =   39
            Top             =   975
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Caracteristicas Complementarias"
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
            Left            =   3720
            TabIndex        =   33
            Top             =   180
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.Label lbl_descripcion 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Denominación del Repuesto"
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
            TabIndex        =   32
            Top             =   240
            Width           =   2565
         End
         Begin VB.Label lbl_codigo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Código del Repuesto"
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
            Left            =   8280
            TabIndex        =   31
            Top             =   240
            Width           =   1920
         End
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
      Begin VB.PictureBox Img_Foto 
         Height          =   2055
         Left            =   5880
         ScaleHeight     =   1995
         ScaleWidth      =   1995
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   8640
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "estado_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
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
         Left            =   9600
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl_det 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "par_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
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
         Left            =   6240
         TabIndex        =   49
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
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
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
         Left            =   7320
         TabIndex        =   43
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
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
         Left            =   3480
         TabIndex        =   20
         Top             =   360
         Width           =   1095
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
         Left            =   4680
         TabIndex        =   22
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "tw_identificacion_cliente.ado_detalle5"
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
         Left            =   960
         TabIndex        =   21
         Top             =   360
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
         Left            =   2565
         TabIndex        =   19
         Top             =   240
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
         Left            =   195
         TabIndex        =   18
         Top             =   240
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
      ScaleWidth      =   11550
      TabIndex        =   11
      Top             =   8205
      Width           =   11550
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
      Top             =   8400
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
      Top             =   8400
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
      Top             =   8400
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   7320
      Top             =   8400
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
End
Attribute VB_Name = "tw_solicitud_bienes5"
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
        mw_solicitud.Ado_detalle1.Recordset.CancelUpdate
        Unload Me
    End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     Call CmdCalcula_Click
     Set rs_aux1 = New ADODB.Recordset
     SQL_FOR = "select * from ao_solicitud_bienes where unidad_codigo = '" & txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " and bien_codigo = '" & dtc_codigo1.Text & "'  "
     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
     If rs_aux1.RecordCount > 0 Then
        If swnuevo = 1 Then
            MsgBox "El bien ya existe, verifique los datos y vuelva a intentar ..."
            'var_cod = 0
            Exit Sub
        End If
     Else
     End If
     If lbl_det.Caption = "30000" Then
         If swnuevo = 1 Then
           ' tw_identificacion_cliente.Ado_detalle3.Recordset("ges_gestion").Value = glGestion
              tw_identificacion_cliente.Ado_detalle3.Recordset("ges_gestion").Value = tw_identificacion_cliente.Ado_datos.Recordset("ges_gestion")
            
            tw_identificacion_cliente.Ado_detalle3.Recordset("unidad_codigo").Value = txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle3.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle3.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle3.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle3.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle3.Recordset("bien_codigo").Value = IIf(txt_campo5.Text = "", "NA1", txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("marca_codigo").Value = IIf(txt_campo8.Text = "", "S/M", txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("modelo_codigo").Value = IIf(txt_campo9.Text = "", "S/M", txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle3.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("marca_codigo").Value = IIf(txt_campo2.Text = "", "S/M", txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle3.Recordset("modelo_codigo").Value = IIf(txt_campo3.Text = "", "S/M", txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle3.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("subgrupo_codigo").Value = IIf(Dtc_aux2.Text = "", "99900", Dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            
            If Txt_campo16.Text <> "" Then
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_precio_venta_base").Value = IIf(txt_campo10 = "", 0, txt_campo10)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_total_venta").Value = IIf(txt_campo11 = "", 0, txt_campo11)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_precio_compra").Value = IIf(Txt_campo19 = "", 0, Txt_campo19)
            tw_identificacion_cliente.Ado_detalle3.Recordset("bien_total_compra").Value = IIf(Txt_campo20 = "", 0, Txt_campo20)
            tw_identificacion_cliente.Ado_detalle3.Recordset!bien_precio_eur.Value = IIf(Txt_campo21 = "", 0, Txt_campo21.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset!bien_total_eur.Value = IIf(Txt_campo22.Text = "", 0, Txt_campo22.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle3.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = GlExtension
            tw_identificacion_cliente.Ado_detalle3.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle3.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle3.Recordset.UpdateBatch adAffectAll
     End If
     If lbl_det.Caption = "39800" Then
         If Txt_campo16.Text <> "" Then
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
         Else
                Txt_campo16.Text = 1
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
         End If
         If dtc_codigo2.Text = "" Then
                dtc_codigo2.Text = "PZA"
         End If

         If swnuevo = 1 Then
            'tw_identificacion_cliente.Ado_detalle3.Recordset!bien_precio_eur.Value = IIf(Txt_campo21 = "", 0, Txt_campo21.Text)
            'tw_identificacion_cliente.Ado_detalle3.Recordset!bien_total_eur.Value = IIf(Txt_campo22.Caption = "", 0, Txt_campo22.Caption)

            db.Execute "INSERT INTO ao_solicitud_bienes (ges_gestion, unidad_codigo, solicitud_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, marca_codigo, modelo_codigo, bien_cantidad, " & _
                "  bien_cantidad_aux, bien_precio_compra, bien_total_compra, bien_precio_venta_base, bien_total_venta, tipo_moneda, unimed_codigo, unimed_codigo_empaque, " & _
                "  bien_cantidad_por_empaque, venta_o_compra, fosa_dimension_frente, fosa_dimension_fondo, almacen_tipo, bien_codigo_padre, estado_codigo, observacion, " & _
                "  usr_codigo , fecha_registro, bien_precio_eur, bien_total_eur)  " & _
                " VALUES ('" & glGestion & "', '" & txt_campo1.Caption & "', " & txt_codigo.Caption & ", '" & dtc_codigo1.Text & "', '" & dtc_aux1.Text & "', '" & Dtc_aux2.Text & "', '" & dtc_aux3.Text & "', '" & txt_campo2.Text & "', '" & txt_campo3.Text & "', " & Txt_campo16.Text & ",  " & _
                " '0', '0', '0', " & txt_campo10.Text & ", " & txt_campo11.Text & ", '" & IIf(dtc_codigo4.Text = "", "BOB", dtc_codigo4.Text) & "', '" & txt_campo4.Text & "', '" & txt_campo4.Text & "', " & _
                " " & Txt_campo16.Text & ", 'V', '7', '0', 'R', '" & GlExtension & "', 'REG', '" & Trim(TxtObservacion.Text) & "', " & _
                " '" & glusuario & "', '" & Date & "', " & CDbl(Txt_campo21.Text) & ", " & CDbl(Txt_campo22.Text) & ") "
         Else
            db.Execute "UPDATE ao_solicitud_bienes SET bien_codigo='" & dtc_codigo1.Text & "', grupo_codigo='" & dtc_aux1.Text & "', subgrupo_codigo='" & Dtc_aux2.Text & "', par_codigo='" & dtc_aux3.Text & "', tipo_moneda = '" & IIf(dtc_codigo4.Text = "", "BOB", dtc_codigo4.Text) & "' ,  " & _
            " marca_codigo='" & txt_campo2.Text & "', modelo_codigo='" & txt_campo3.Text & "', bien_cantidad=" & Txt_campo16.Text & ", bien_precio_venta_base=" & txt_campo10.Text & ", unimed_codigo='" & txt_campo4.Text & "', unimed_codigo_empaque='" & txt_campo4.Text & "', observacion='" & Trim(TxtObservacion.Text) & "', " & _
            " bien_total_venta=" & txt_campo11.Text & ", usr_codigo='" & glusuario & "', bien_codigo_padre= '" & GlExtension & "',  bien_precio_eur=" & CDbl(Txt_campo21.Text) & ", bien_total_eur = " & CDbl(Txt_campo22.Text) & " WHERE unidad_codigo='" & txt_campo1.Caption & "'  AND solicitud_codigo=" & txt_codigo.Caption & " AND bien_codigo= '" & tw_identificacion_cliente.Ado_detalle5.Recordset!bien_codigo & "' "

            'tw_identificacion_cliente.Ado_detalle3.Recordset!observacion = TxtObservacion.Text
         End If
    End If
    If lbl_det.Caption = "34800" Then
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
                tw_identificacion_cliente.Ado_detalle6.Recordset("bien_codigo").Value = IIf(txt_campo5.Text = "", "NA1", txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("marca_codigo").Value = IIf(txt_campo8.Text = "", "S/M", txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("modelo_codigo").Value = IIf(txt_campo9.Text = "", "S/M", txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle6.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("marca_codigo").Value = IIf(txt_campo2.Text = "", "S/M", txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle6.Recordset("modelo_codigo").Value = IIf(txt_campo3.Text = "", "S/M", txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle6.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("subgrupo_codigo").Value = IIf(Dtc_aux2.Text = "", "99900", Dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_precio_venta_base").Value = IIf(txt_campo10 = "", 0, txt_campo10)
            tw_identificacion_cliente.Ado_detalle6.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_total_venta").Value = IIf(txt_campo11 = "", 0, txt_campo11)
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle6.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle2.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle6.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle6.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle6.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle6.Recordset.UpdateBatch adAffectAll
    End If
    If lbl_det.Caption = "24300" Then
         If swnuevo = 1 Then
            tw_identificacion_cliente.Ado_detalle7.Recordset("ges_gestion").Value = glGestion
            tw_identificacion_cliente.Ado_detalle7.Recordset("unidad_codigo").Value = txt_campo1.Caption
            tw_identificacion_cliente.Ado_detalle7.Recordset("solicitud_codigo").Value = txt_codigo.Caption
            tw_identificacion_cliente.Ado_detalle7.Recordset("estado_codigo").Value = "REG"
            tw_identificacion_cliente.Ado_detalle7.Recordset("venta_o_compra").Value = "V"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_foto_cargado").Value = "N"
    '        tw_identificacion_cliente.Ado_detalle6.Recordset("archivo_plano_cargado").Value = "N"
         End If
            If var_val2 = "1" Then
                tw_identificacion_cliente.Ado_detalle7.Recordset("bien_codigo").Value = IIf(txt_campo5.Text = "", "NA1", txt_campo5.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("marca_codigo").Value = IIf(txt_campo8.Text = "", "S/M", txt_campo8.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("modelo_codigo").Value = IIf(txt_campo9.Text = "", "S/M", txt_campo9.Text)
            Else
                'OJO FALTA GRABAR EN ac_bienes
                tw_identificacion_cliente.Ado_detalle7.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("marca_codigo").Value = IIf(txt_campo2.Text = "", "S/M", txt_campo2.Text)
                tw_identificacion_cliente.Ado_detalle7.Recordset("modelo_codigo").Value = IIf(txt_campo3.Text = "", "S/M", txt_campo3.Text)
            End If
            tw_identificacion_cliente.Ado_detalle7.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("subgrupo_codigo").Value = IIf(Dtc_aux2.Text = "", "99900", Dtc_aux2.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
            If Txt_campo16.Text <> "" Then
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            Else
                Txt_campo16.Text = 1
                txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
            End If
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_precio_venta_base").Value = IIf(txt_campo10 = "", 0, txt_campo10)
            tw_identificacion_cliente.Ado_detalle7.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_total_venta").Value = IIf(txt_campo11 = "", 0, txt_campo11)
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_precio_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle7.Recordset("bien_total_compra").Value = 0
            tw_identificacion_cliente.Ado_detalle2.Recordset!observacion = TxtObservacion.Text
            'Tipo de Solicitud ******************************
            tw_identificacion_cliente.Ado_detalle7.Recordset("fosa_dimension_frente").Value = IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            tw_identificacion_cliente.Ado_detalle7.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            tw_identificacion_cliente.Ado_detalle7.Recordset("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            tw_identificacion_cliente.Ado_detalle7.Recordset("usr_codigo").Value = glusuario
            tw_identificacion_cliente.Ado_detalle7.Recordset.UpdateBatch adAffectAll
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
            rs_aux4!modelo_codigo = txt_campo3.Text     'Ado_datos.Recordset!modelo_codigo
            rs_aux4!modelo_codigo_h = "0"        'Ado_datos.Recordset!modelo_codigo_h1
            rs_aux4!modelo_codigo_x = "0"       'Ado_datos.Recordset!modelo_codigo_x1
            rs_aux4!cotiza_fecha = Date
            rs_aux4!cotiza_cantidad = IIf(Txt_campo16 = "", 1, Txt_campo16)
            rs_aux4!cotiza_tdc_bol = GlTipoCambioOficial
            rs_aux4!cotiza_precio_fob_bs = IIf(txt_campo10 = "", 0, txt_campo10)
            rs_aux4!cotiza_precio_fob_dol = CDbl(txt_campo10) * GlTipoCambioOficial
            rs_aux4!cotiza_precio_total_bs = IIf(txt_campo11 = "", 0, txt_campo11)
            rs_aux4!cotiza_precio_total_dol = CDbl(txt_campo11) * GlTipoCambioOficial
            rs_aux4!costo_monto = IIf(txt_campo11 = "", 0, txt_campo11)
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
       
'     Frame1.Visible = False
'     Frame2.Visible = False

     Unload Me

   '  Call ABRIR_TABLA
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

Private Sub CmdCalcula_Click()
    'GlTipoCambioOficial        'USD
    'GlTipoCambioEuro           'EUR
    If dtc_codigo4.Text = "" Then dtc_codigo4.Text = "BOB"
    Select Case dtc_codigo4.Text
        Case "BOB"
            If txt_campo10.Text = "" Then txt_campo10.Text = "0"
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
            Txt_campo19.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioOficial, 2)
            Txt_campo21.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioEuro, 2)
            If Txt_campo16.Text = "" Then Txt_campo16.Text = "1"
            txt_campo11.Text = Round(CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)     'BOB (Total)
            Txt_campo20.Text = Round(CDbl(Txt_campo19.Text) * CDbl(Txt_campo16.Text), 2)     'USD (Total)
            Txt_campo22.Text = Round(CDbl(Txt_campo21.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Total)
        Case "USD"
            If Txt_campo19.Text = "" Then Txt_campo19.Text = "0"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = True      'USD
            Txt_campo21.Enabled = False     'EUR
            txt_campo10.Text = Round(CDbl(Txt_campo19.Text) * GlTipoCambioOficial, 2)
            Txt_campo21.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioEuro, 2)
            If Txt_campo16.Text = "" Then Txt_campo16.Text = "1"
            txt_campo11.Text = Round(CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)     'BOB (Total)
            Txt_campo20.Text = Round(CDbl(Txt_campo19.Text) * CDbl(Txt_campo16.Text), 2)     'USD (Total)
            Txt_campo22.Text = Round(CDbl(Txt_campo21.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Total)
        Case "EUR"
            If Txt_campo21.Text = "" Then Txt_campo21.Text = "0"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = True      'EUR
            txt_campo10.Text = Round(CDbl(Txt_campo21.Text) * GlTipoCambioEuro, 2)
            Txt_campo19.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioOficial, 2)
            If Txt_campo16.Text = "" Then Txt_campo16.Text = "1"
            txt_campo11.Text = Round(CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)     'BOB (Total)
            Txt_campo20.Text = Round(CDbl(Txt_campo19.Text) * CDbl(Txt_campo16.Text), 2)     'USD (Total)
            Txt_campo22.Text = Round(CDbl(Txt_campo21.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Total)
        Case Else
            If txt_campo10.Text = "" Then txt_campo10.Text = "0"
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
            Txt_campo19.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioOficial, 2)
            Txt_campo21.Text = Round(CDbl(txt_campo10.Text) / GlTipoCambioEuro, 2)
            If Txt_campo16.Text = "" Then Txt_campo16.Text = "1"
            txt_campo11.Text = Round(CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)     'BOB (Total)
            Txt_campo20.Text = Round(CDbl(Txt_campo19.Text) * CDbl(Txt_campo16.Text), 2)     'USD (Total)
            Txt_campo22.Text = Round(CDbl(Txt_campo21.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Total)
    End Select

End Sub

Private Sub CmdCalcula2_Click()
    If dtc_codigo4.Text = "" Then dtc_codigo4.Text = "BOB"
    Select Case dtc_codigo4.Text
        Case "BOB"
            If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then Txt_campo16.Text = "1"
            If txt_campo11.Text = "" Then txt_campo11.Text = "0"
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
            txt_campo11.Enabled = True      'BOB
            Txt_campo20.Enabled = False     'USD
            Txt_campo22.Enabled = False     'EUR
            Txt_campo20.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioOficial, 2)
            Txt_campo22.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioEuro, 2)
            
            txt_campo10.Text = Round(CDbl(txt_campo11.Text) / CDbl(Txt_campo16.Text), 2)     'BOB (Unitario)
            Txt_campo19.Text = Round(CDbl(Txt_campo20.Text) / CDbl(Txt_campo16.Text), 2)     'USD (Unitario)
            Txt_campo21.Text = Round(CDbl(Txt_campo22.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Unitario)
        Case "USD"
            If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then Txt_campo16.Text = "1"
            If Txt_campo20.Text = "" Then Txt_campo20.Text = "0"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = True      'USD
            Txt_campo21.Enabled = False     'EUR
            txt_campo11.Enabled = False     'BOB
            Txt_campo20.Enabled = True      'USD
            Txt_campo22.Enabled = False     'EUR
            txt_campo11.Text = Round(CDbl(Txt_campo20.Text) * GlTipoCambioOficial, 2)
            Txt_campo22.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioEuro, 2)
            
            txt_campo10.Text = Round(CDbl(txt_campo11.Text) / CDbl(Txt_campo16.Text), 2)     'BOB (Unitario)
            Txt_campo19.Text = Round(CDbl(Txt_campo20.Text) / CDbl(Txt_campo16.Text), 2)     'USD (Unitario)
            Txt_campo21.Text = Round(CDbl(Txt_campo22.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Unitario)
        Case "EUR"
            If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then Txt_campo16.Text = "1"
            If Txt_campo22.Text = "" Then Txt_campo22.Text = "0"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = True      'EUR
            txt_campo11.Enabled = False     'BOB
            Txt_campo20.Enabled = False     'USD
            Txt_campo22.Enabled = True      'EUR
            txt_campo11.Text = Round(CDbl(Txt_campo22.Text) * GlTipoCambioEuro, 2)
            Txt_campo20.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioOficial, 2)
            
            txt_campo10.Text = Round(CDbl(txt_campo11.Text) / CDbl(Txt_campo16.Text), 2)     'BOB (Unitario)
            Txt_campo19.Text = Round(CDbl(Txt_campo20.Text) / CDbl(Txt_campo16.Text), 2)     'USD (Unitario)
            Txt_campo21.Text = Round(CDbl(Txt_campo22.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Unitario)
        Case Else
            If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then Txt_campo16.Text = "1"
            If txt_campo11.Text = "" Then txt_campo11.Text = "0"
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
            Txt_campo20.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioOficial, 2)
            Txt_campo22.Text = Round(CDbl(txt_campo11.Text) / GlTipoCambioEuro, 2)
            
            txt_campo10.Text = Round(CDbl(txt_campo11.Text) / CDbl(Txt_campo16.Text), 2)     'BOB (Unitario)
            Txt_campo19.Text = Round(CDbl(Txt_campo20.Text) / CDbl(Txt_campo16.Text), 2)     'USD (Unitario)
            Txt_campo21.Text = Round(CDbl(Txt_campo22.Text) * CDbl(Txt_campo16.Text), 2)     'EUR (Unitario)
    End Select

End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux1.BoundText
    dtc_desc1.BoundText = dtc_aux1.BoundText
    Dtc_aux2.BoundText = dtc_aux1.BoundText
    dtc_aux3.BoundText = dtc_aux1.BoundText
    txt_campo2.BoundText = dtc_aux1.BoundText
    txt_campo3.BoundText = dtc_aux1.BoundText
    txt_campo4.BoundText = dtc_aux1.BoundText
    Txt_campo18.BoundText = dtc_aux1.BoundText
    'dtc_codigo2.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo1.BoundText = Dtc_aux2.BoundText
    dtc_desc1.BoundText = Dtc_aux2.BoundText
    dtc_aux1.BoundText = Dtc_aux2.BoundText
    dtc_aux3.BoundText = Dtc_aux2.BoundText
    txt_campo2.BoundText = Dtc_aux2.BoundText
    txt_campo3.BoundText = Dtc_aux2.BoundText
    txt_campo4.BoundText = Dtc_aux2.BoundText
    Txt_campo18.BoundText = Dtc_aux2.BoundText
    'dtc_codigo2.BoundText = dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_aux3.BoundText
    dtc_desc1.BoundText = dtc_aux3.BoundText
    Dtc_aux2.BoundText = dtc_aux3.BoundText
    dtc_aux1.BoundText = dtc_aux3.BoundText
    txt_campo2.BoundText = dtc_aux3.BoundText
    txt_campo3.BoundText = dtc_aux3.BoundText
    txt_campo4.BoundText = dtc_aux3.BoundText
    Txt_campo18.BoundText = dtc_aux3.BoundText
    'dtc_codigo2.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    txt_campo2.BoundText = dtc_codigo1.BoundText
    txt_campo3.BoundText = dtc_codigo1.BoundText
    txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    'dtc_codigo2.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo1_LostFocus()
    If Len(TxtObservacion.Text) = 0 Then
        'TxtObservacion.Text = TxtObservacion.Text
    'Else
        TxtObservacion.Text = dtc_desc1.Text + " - " + txt_campo4
    End If
End Sub

Private Sub dtc_codigo2_Change()
    dtc_desc2.BoundText = dtc_codigo2.Text
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.Text
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Change()
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
    txt_campo2.BoundText = dtc_desc1.BoundText
    txt_campo3.BoundText = dtc_desc1.BoundText
    txt_campo4.BoundText = dtc_desc1.BoundText
    Txt_campo18.BoundText = dtc_desc1.BoundText
'    dtc_codigo2.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
    txt_campo2.BoundText = dtc_desc1.BoundText
    txt_campo3.BoundText = dtc_desc1.BoundText
    txt_campo4.BoundText = dtc_desc1.BoundText
    Txt_campo18.BoundText = dtc_desc1.BoundText
'    dtc_codigo2.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
    If Len(TxtObservacion.Text) = 0 Then
        TxtObservacion.Text = dtc_desc1.Text + " - " + txt_campo4
    End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc4_LostFocus()
    'GlTipoCambioOficial        'USD
    'GlTipoCambioEuro           'EUR
    If dtc_codigo4.Text = "" Then dtc_codigo4.Text = "BOB"
    Select Case dtc_codigo4.Text
        Case "BOB"
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
        Case "USD"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = True      'USD
            Txt_campo21.Enabled = False     'EUR
        Case "EUR"
            txt_campo10.Enabled = False     'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = True      'EUR
        Case Else
            txt_campo10.Enabled = True      'BOB
            Txt_campo19.Enabled = False     'USD
            Txt_campo21.Enabled = False     'EUR
    End Select
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
	Call SeguridadSet(Me)
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
        'rs_datos1.Open "select * from ac_bienes where (par_codigo= '39810' or par_codigo= '39820') and estado_codigo ='APR' ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion 'and estado_codigo = 'APR'
        rs_datos1.Open "select * from ac_bienes where almacen_tipo = 'R' AND par_codigo <> '24300' and estado_codigo ='APR' ORDER BY bien_descripcion ", db, adOpenStatic
    End If
    If lbl_det = "24300" Then
            rs_datos1.Open "select * from ac_bienes where par_codigo= '24300'  ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    If lbl_det = "40000" Then
        '    rs_datos1.Open "select * from ac_bienes where grupo_codigo = '30000' OR par_codigo= '22210' OR (grupo_codigo = '40000' and par_codigo <> '43340') ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
        rs_datos1.Open "select * from ac_bienes where (grupo_codigo = '40000' and par_codigo <> '43340') ORDER BY bien_descripcion ", db, adOpenStatic   'order by descripcion
    End If
    'End If
        'lbl_det = " par_codigo = '" & 34110 & "' AND par_codigo = '" & 33100 & "' "
    
    'rs_datos1.Open "select * from ac_bienes where par_codigo = '" & 43340 & "' AND edif_codigo = '" & lbl_edif.Caption & "' ", db, adOpenStatic   'order by descripcion
    Set Ado_datos1.Recordset = rs_datos1
    If swnuevo = 2 Then
        dtc_codigo1.Text = txt_campo5.Text
    End If
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    txt_campo2.BoundText = dtc_codigo1.BoundText
    txt_campo3.BoundText = dtc_codigo1.BoundText
    txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    
    'ac_bienes_unidad_medida
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If lbl_det = "43340" Or lbl_det = "24300" Then
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo = 'T' order by unimed_descripcion ", db, adOpenStatic
    Else
        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' order by unimed_descripcion ", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_datos2
    If swnuevo = 2 Then
        dtc_codigo2.Text = Txt_campo14.Text
    End If
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    If lbl_det = "39800" Then
        rs_datos3.Open "Select * from gc_tipo_solicitud WHERE solicitud_tipo = 7 OR solicitud_tipo = 10 OR solicitud_tipo = 4 OR solicitud_tipo = 5 OR solicitud_tipo = 8 order by solicitud_tipo_descripcion ", db, adOpenStatic
    Else
        rs_datos3.Open "Select * from gc_tipo_solicitud  order by solicitud_tipo_descripcion ", db, adOpenStatic
    End If
    Set Ado_datos3.Recordset = rs_datos3
    If swnuevo = 2 Then
        dtc_codigo5.Text = Txt_campo15.Text
    End If
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    If lbl_det = "39800" Then
        rs_datos4.Open "Select * from gc_tipo_moneda WHERE estado_codigo = 'APR' ", db, adOpenStatic
    End If
    Set Ado_datos4.Recordset = rs_datos4
    'If swnuevo = 2 Then
    '    dtc_codigo4.Text = Txt_campo4.Text
    'End If
    dtc_desc4.BoundText = dtc_codigo4.BoundText
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
'    If Txt_campo16.Text = "" Then
'        Txt_campo16.Text = "1"
'    End If
'    If txt_campo10.Text <> "" Then
'        txt_campo11.Text = Round(CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)
'        If txt_campo11.Text = "0" Then
'            Txt_campo19.Text = "0"
'        Else
'            Txt_campo19.Text = Round(CDbl(txt_campo11.Text) / 6.96, 2)
'        End If
'        Txt_campo20.Text = Round(CDbl(Txt_campo19.Text) * CDbl(Txt_campo16.Text), 2)
'    End If
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
    If txt_campo10.Text = "" Then
        txt_campo10.Text = "1"
    End If
    If Txt_campo16.Text <> "" Then
    txt_campo11.Text = CDbl(txt_campo10.Text) * CDbl(Txt_campo16.Text)
    End If
End Sub

Private Sub Txt_campo18_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo18.BoundText
    dtc_desc1.BoundText = Txt_campo18.BoundText
    Dtc_aux2.BoundText = Txt_campo18.BoundText
    dtc_aux3.BoundText = Txt_campo18.BoundText
    txt_campo2.BoundText = Txt_campo18.BoundText
    dtc_aux1.BoundText = Txt_campo18.BoundText
    txt_campo4.BoundText = Txt_campo18.BoundText
    txt_campo3.BoundText = Txt_campo18.BoundText
    dtc_codigo2.BoundText = Txt_campo18.BoundText
End Sub

'Private Sub txt_campo3_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub Txt_campo2_Click(Area As Integer)
    dtc_codigo1.BoundText = txt_campo2.BoundText
    dtc_desc1.BoundText = txt_campo2.BoundText
    Dtc_aux2.BoundText = txt_campo2.BoundText
    dtc_aux3.BoundText = txt_campo2.BoundText
    dtc_aux1.BoundText = txt_campo2.BoundText
    txt_campo3.BoundText = txt_campo2.BoundText
    txt_campo4.BoundText = txt_campo2.BoundText
    Txt_campo18.BoundText = txt_campo2.BoundText
    'dtc_codigo2.BoundText = Txt_campo2.BoundText
End Sub

Private Sub Txt_campo3_Click(Area As Integer)
    dtc_codigo1.BoundText = txt_campo3.BoundText
    dtc_desc1.BoundText = txt_campo3.BoundText
    Dtc_aux2.BoundText = txt_campo3.BoundText
    dtc_aux3.BoundText = txt_campo3.BoundText
    txt_campo2.BoundText = txt_campo3.BoundText
    dtc_aux1.BoundText = txt_campo3.BoundText
    txt_campo4.BoundText = txt_campo3.BoundText
    Txt_campo18.BoundText = txt_campo3.BoundText
    'dtc_codigo2.BoundText = Txt_campo3.BoundText
End Sub

Private Sub Txt_campo4_Click(Area As Integer)
    dtc_codigo1.BoundText = txt_campo4.BoundText
    dtc_desc1.BoundText = txt_campo4.BoundText
    Dtc_aux2.BoundText = txt_campo4.BoundText
    dtc_aux3.BoundText = txt_campo4.BoundText
    txt_campo2.BoundText = txt_campo4.BoundText
    dtc_aux1.BoundText = txt_campo4.BoundText
    txt_campo3.BoundText = txt_campo4.BoundText
    Txt_campo18.BoundText = txt_campo4.BoundText
    'dtc_codigo2.BoundText = Txt_campo4.BoundText
End Sub

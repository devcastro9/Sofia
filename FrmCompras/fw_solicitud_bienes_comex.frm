VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form fw_solicitud_bienes_comex 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7290
   ClientLeft      =   1065
   ClientTop       =   -30
   ClientWidth     =   10965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   135
      ScaleHeight     =   675
      ScaleWidth      =   10680
      TabIndex        =   47
      Top             =   6495
      Width           =   10680
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5355
         Picture         =   "fw_solicitud_bienes_comex.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   49
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "fw_solicitud_bienes_comex.frx":08EC
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   48
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   13215
         TabIndex        =   50
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame SWS 
      BackColor       =   &H00C0C0C0&
      Height          =   6495
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   10440
         TabIndex        =   66
         Top             =   200
         Width           =   10440
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRO DE BIENES O SERVICIOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   360
            Left            =   2640
            TabIndex        =   68
            Top             =   0
            Width           =   5325
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   13215
            TabIndex        =   67
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Frame fra_almacen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ALMACEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   750
         Left            =   120
         TabIndex        =   54
         Top             =   4080
         Width           =   10485
         Begin MSDataListLib.DataCombo dtc_desc_alm 
            Bindings        =   "fw_solicitud_bienes_comex.frx":10C2
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   305
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_cod_alm 
            Bindings        =   "fw_solicitud_bienes_comex.frx":10DC
            DataField       =   "almacen_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   6720
            TabIndex        =   56
            Top             =   305
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_codigo"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
      End
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
         Height          =   1575
         Left            =   120
         TabIndex        =   33
         Top             =   4800
         Width           =   10455
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3120
            TabIndex        =   65
            Top             =   810
            Width           =   350
         End
         Begin MSDataListLib.DataCombo Txt_campo14 
            Bindings        =   "fw_solicitud_bienes_comex.frx":10F6
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   2355
            TabIndex        =   64
            Top             =   795
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unimed_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.TextBox Txt_tdc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "compra_tdc"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4800
            TabIndex        =   62
            Text            =   "6.96"
            Top             =   800
            Width           =   855
         End
         Begin VB.ComboBox cmd_moneda 
            DataField       =   "tipo_moneda"
            DataSource      =   "Ado_datos02"
            Height          =   315
            ItemData        =   "fw_solicitud_bienes_comex.frx":110F
            Left            =   3720
            List            =   "fw_solicitud_bienes_comex.frx":111C
            TabIndex        =   61
            Text            =   "BOB"
            Top             =   800
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "compra_precio_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "fw_compras_comex.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7080
            TabIndex        =   59
            Text            =   "1"
            Top             =   800
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "compra_precio_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "fw_compras_comex.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8760
            TabIndex        =   57
            Text            =   "1"
            Top             =   795
            Width           =   1455
         End
         Begin VB.TextBox Txt_campo19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "bien_cantidad_por_empaque"
            DataSource      =   "fw_solicitud_compras.ado_detalle2"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7080
            TabIndex        =   46
            Text            =   "1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Txt_campo10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "compra_precio_unitario_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "1"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Txt_campo16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "compra_cantidad"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5880
            TabIndex        =   39
            Text            =   "1"
            Top             =   800
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "fw_solicitud_bienes_comex.frx":112F
            DataField       =   "unimed_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   2760
            TabIndex        =   43
            Top             =   1155
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unimed_descripcion"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "fw_solicitud_bienes_comex.frx":1148
            DataField       =   "unimed_codigo"
            Height          =   315
            Left            =   1800
            TabIndex        =   44
            Top             =   1155
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483637
            ForeColor       =   0
            ListField       =   "unimed_codigo"
            BoundColumn     =   "unimed_codigo"
            Text            =   ""
         End
         Begin VB.Label Lbl_TDC 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "TDC"
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
            Left            =   4800
            TabIndex        =   69
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total Dol."
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
            Left            =   7080
            TabIndex        =   60
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Txt_estado 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "REG"
            DataField       =   "estado_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
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
            Left            =   8880
            TabIndex        =   42
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Txt_campo11A 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "compra_precio_unitario_dol"
            DataSource      =   "fw_compras_comex.ado_detalle1"
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
            Left            =   240
            TabIndex        =   41
            Top             =   1155
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Moneda"
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
            Left            =   3600
            TabIndex        =   38
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lbl_campo10 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Importe Unitario Bs. "
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
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1155
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
            Height          =   240
            Left            =   1800
            TabIndex        =   36
            Top             =   480
            Width           =   1770
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
            Height          =   240
            Left            =   5760
            TabIndex        =   35
            Top             =   480
            Width           =   1215
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
            Height          =   240
            Left            =   8760
            TabIndex        =   34
            Top             =   480
            Width           =   1455
         End
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
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   10455
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   7440
            TabIndex        =   58
            Top             =   540
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   9600
            TabIndex        =   51
            Top             =   540
            Width           =   360
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2865
            TabIndex        =   32
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   2115
            Width           =   255
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9840
            TabIndex        =   20
            Top             =   1320
            Width           =   360
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "fw_solicitud_bienes_comex.frx":1161
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   7800
            TabIndex        =   1
            Top             =   525
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "fw_solicitud_bienes_comex.frx":117A
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   525
            Width           =   7590
            _ExtentX        =   13388
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Txt_campo4 
            Bindings        =   "fw_solicitud_bienes_comex.frx":1193
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1300
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_descripcion_anterior"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo2 
            Bindings        =   "fw_solicitud_bienes_comex.frx":11AC
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   3360
            TabIndex        =   25
            Top             =   2100
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "marca_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo3 
            Bindings        =   "fw_solicitud_bienes_comex.frx":11C5
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   6840
            TabIndex        =   26
            Top             =   2100
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "modelo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Txt_campo18 
            Bindings        =   "fw_solicitud_bienes_comex.frx":11DE
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   2100
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "pais_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Txt_campo5 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "bien_codigo"
            DataSource      =   "fw_compras_comex.ado_detalle1"
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
            Left            =   8880
            TabIndex        =   70
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
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
            TabIndex        =   30
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   1800
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
            Left            =   240
            TabIndex        =   21
            Top             =   1020
            Width           =   2970
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
            TabIndex        =   19
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl_codigo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7800
            TabIndex        =   18
            Top             =   240
            Width           =   660
         End
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "fw_solicitud_bienes_comex.frx":11F7
         DataField       =   "bien_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle1"
         Height          =   315
         Left            =   8400
         TabIndex        =   16
         Top             =   600
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
      Begin MSDataListLib.DataCombo dtc_aux2 
         Bindings        =   "fw_solicitud_bienes_comex.frx":1210
         DataField       =   "bien_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle1"
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Top             =   600
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
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "fw_solicitud_bienes_comex.frx":1229
         DataField       =   "bien_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle1"
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   600
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "compra_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle1"
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
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label MOD_NEW 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         Height          =   300
         Left            =   2400
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label txt_gestion 
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
         ForeColor       =   &H00FFFF80&
         Height          =   300
         Left            =   1080
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_det 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "par_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle1"
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
         Left            =   8400
         TabIndex        =   45
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_edif 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "compra_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle1"
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
         Left            =   9360
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "0"
         DataField       =   "unidad_codigo"
         DataSource      =   "fw_compras_gral.ado_detalle1"
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
         Left            =   7200
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Txt_descripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "fw_compras_comex.ado_detalle1"
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
         TabIndex        =   13
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "fw_compras_comex.ado_detalle1"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad Ejecutora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   8
         Left            =   2685
         TabIndex        =   10
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lbl_codigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Tr�mite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   195
         TabIndex        =   9
         Top             =   720
         Width           =   870
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
      ScaleWidth      =   10965
      TabIndex        =   2
      Top             =   7290
      Width           =   10965
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   7
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   2520
      Top             =   7440
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
      Left            =   4920
      Top             =   7440
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
   Begin MSAdodcLib.Adodc Ado_clasif6 
      Height          =   330
      Left            =   0
      Top             =   7440
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
Attribute VB_Name = "fw_solicitud_bienes_comex"
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
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset

Dim rs_clasif6 As New ADODB.Recordset
Dim rs_UNIDAD As New ADODB.Recordset
Dim rs_compra_det As New ADODB.Recordset 'rs_compra_cab
Dim rs_compra_cab As New ADODB.Recordset
'BUSCADOR
Dim var_val2 As String
Dim var_cod5 As String
Dim VAR_VAL, CAL_DOL, CAL_BS As String
Dim VAR_BENEF As String

Dim var_ctm, var_itm As Double

Dim NRO_REG As Integer

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Est� Seguro de CANCELAR la operaci�n ? ", vbYesNo + vbQuestion, "Atenci�n")
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
    If swnuevo = 1 Then     'Or VAR_SW = "NEW"
        NRO_REG = fw_compras_comex.Ado_detalle1.Recordset.RecordCount + 1
        ' compra_codigo_det,            " & NRO_REG & ",
        db.Execute "insert into ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo,            compra_cantidad,            compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,             grupo_codigo,           subgrupo_codigo, " & _
                   " par_codigo,            tipo_descuento, almacen_codigo, usr_usuario,    fecha_registro, hora_registro, unimed_codigo,   estado_codigo, adjudica_monto_bs_87, compra_tdc, tipo_moneda) " & _
                   " VALUES ('" & txt_gestion.Caption & "', " & Val(lbl_edif.Caption) & ", '" & dtc_codigo1.Text & "', " & Val(Txt_campo16.Text) & ", " & CDbl(Txt_campo10.Text) & ",      '0',            " & CDbl(Txt_campo11.Text) & ", " & CDbl(Text2.Text) & ",           '0',              " & CDbl(Text2.Text) & ",       '" & dtc_desc1.Text & "', '" & dtc_aux1.Text & "', '" & Dtc_aux2.Text & "',  " & _
                   " '" & dtc_aux3.Text & "', '0',          '1',        '" & glusuario & "', '" & Date & "', '0',       '" & Txt_campo14 & "', 'REG', '0', '6.96', 'BOB'  )"
'
'        db.Execute "insert into ao_compra_detalle (ges_gestion, cobranza_detalle, cobranza_codigo, beneficiario_codigo_resp,    cobranza_bs,       cobranza_dol,           cobranza_fecha,             cobranza_observaciones,             cta_codigo,             cmpbte_deposito,            doc_numero,              trans_codigo,              literal,    estado_codigo, estado_codigo_bco, usr_codigo,           fecha_registro,                     tipo_moneda, usr_codigo_mod, usr_codigo_apr, cobranza_tdc) " & _
                " VALUES ('" & Ado_datos01.Recordset!ges_gestion & "', " & correldet & ", " & NRO_COBR & ", '" & DataCombo1.Text & "', " & COBR_BS & ", " & VAR_DOL2 & ", '" & CDate(DTPicker1.Value) & "', '" & txt_observaciones.Text & "', '" & dtc_cta2.Text & "', '" & Txt_deposito.Text & "', " & Txt_docnro.Text & ", '" & DataCombo9.Text & "', '" & var_literal & "', 'REG',      'REG', '" & glusuario & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & cmd_moneda.Text & "', '" & glusuario & "', '" & glusuario & "', " & TxtTDC.Text & " )"
        
'        fw_compras_comex.Ado_detalle1.Recordset("ges_gestion").Value = txt_gestion.Caption
'        fw_compras_comex.Ado_detalle1.Recordset("unidad_codigo").Value = Txt_campo1.Caption 'fw_compras_comex.VAR_UNI
'        fw_compras_comex.Ado_detalle1.Recordset("solicitud_codigo").Value = txt_codigo.Caption
'        fw_compras_comex.Ado_detalle1.Recordset("estado_codigo").Value = "REG"
'        fw_compras_comex.Ado_detalle1.Recordset("venta_o_compra").Value = "C" 'C = PAGOS PERIODICOS � CREDITO y E = EFECTIVO (Al Contado)
''        fw_compras_comex.Ado_detalle1.Recordset("archivo_foto_cargado").Value = "N"
''        fw_compras_comex.Ado_detalle1.Recordset("archivo_plano_cargado").Value = "N"
'
'        fw_compras_comex.Ado_detalle1.Recordset("bien_codigo").Value = IIf(dtc_codigo1.Text = "", "NA1", dtc_codigo1.Text)
'        fw_compras_comex.Ado_detalle1.Recordset("marca_codigo").Value = IIf(Txt_campo8.Text = "", "S/M", Txt_campo8.Text)
'        fw_compras_comex.Ado_detalle1.Recordset("modelo_codigo").Value = IIf(Txt_campo9.Text = "", "S/M", Txt_campo9.Text)
'
'        fw_compras_comex.Ado_detalle1.Recordset!compra_concepto = dtc_desc1.Text            'Txt_campo4.Text
'        fw_compras_comex.Ado_detalle1.Recordset("grupo_codigo").Value = IIf(dtc_aux1.Text = "", "90000", dtc_aux1.Text)
'        fw_compras_comex.Ado_detalle1.Recordset("subgrupo_codigo").Value = IIf(dtc_aux2.Text = "", "99900", dtc_aux2.Text)
'        fw_compras_comex.Ado_detalle1.Recordset("par_codigo").Value = IIf(dtc_aux3.Text = "", "99990", dtc_aux3.Text)
'        fw_compras_comex.Ado_detalle1.Recordset("bien_precio_compra").Value = IIf(Txt_campo10 = "", 0, Txt_campo10)
'        'fw_compras_comex.Ado_detalle1.Recordset("unimed_codigo").Value = IIf(dtc_codigo2 = "", "MES", dtc_codigo2)
'        fw_compras_comex.Ado_detalle1.Recordset("bien_cantidad").Value = IIf(Txt_campo16 = "", 1, Txt_campo16)
'        fw_compras_comex.Ado_detalle1.Recordset("bien_total_compra").Value = IIf(Txt_campo11 = "", 0, Txt_campo11)
'        fw_compras_comex.Ado_detalle1.Recordset("bien_cantidad_por_empaque").Value = IIf(Txt_campo19 = "", 2, Txt_campo19)
'        'fw_compras_comex.Ado_detalle3.Recordset("bien_total_compra").Value = 0    '
'
'        Select Case VAR_UNI
'           Case "DNMAN"
'               fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_frente").Value = "10"
'           Case "DNREP"
'               fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_frente").Value = "7"
'           Case "DNINS"
'               fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_frente").Value = "4"
'           Case "DNAJS"
'               fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_frente").Value = "5"
'           Case "DNMOD"
'               fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_frente").Value = "9"
'           Case Else
'           fw_compras_comex.Ado_detalle1.Recordset("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
'
'        End Select
'            'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
'        'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
'        fw_compras_comex.Ado_detalle1.Recordset("fecha_registro").Value = Date
'        'aw_p_ao_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
'        fw_compras_comex.Ado_detalle1.Recordset("usr_codigo").Value = glusuario
'        fw_compras_comex.Ado_detalle1.Recordset.UpdateBatch adAffectAll
    End If
    If swnuevo = 2 Then
'      Set rs_compra_cab = New Recordset
'    If rs_compra_cab.State = 1 Then rs_compra_cab.Close
'    rs_compra_cab.Open "Select * from ao_compra_cabecera where ges_gestion = '" & Year(Date) & "' and solicitud_codigo =" & rs_datos!solicitud_codigo & " and unidad_codigo = '" & rs_datos!unidad_codigo & "' and unidad_codigo_adm ='" & rs_datos!unidad_codigo & "'", db, adOpenKeyset, adLockOptimistic
    
    'rs_compra_det
     Set rs_compra_det = New Recordset
     If rs_compra_det.State = 1 Then rs_compra_det.Close
     
     If MOD_NEW.Caption = "NEW" Then
   
        rs_compra_det.Open "Select * from ao_compra_detalle", db, adOpenKeyset, adLockOptimistic
        rs_compra_det.AddNew
     
        If rs_aux8.State = 1 Then rs_aux8.Close
        rs_aux8.Open "Select max(compra_codigo_det) as correla from ao_compra_detalle WHERE compra_codigo = " & fw_compras_comex.Ado_datos.Recordset!compra_codigo & "", db, adOpenKeyset, adLockOptimistic
        If IsNull(rs_aux8!correla) = True Then
            rs_compra_det!compra_codigo_det = "1"
        Else
            rs_compra_det!compra_codigo_det = rs_aux8!correla + 1
        End If
      
        Set rs_aux7 = New ADODB.Recordset
        If rs_aux7.State = 1 Then rs_aux7.Close
        rs_aux7.Open "Select * from ao_compra_detalle WHERE compra_codigo = " & fw_compras_comex.Ado_datos.Recordset!compra_codigo & " AND bien_codigo = '" & dtc_codigo1.Text & "' AND compra_codigo_det <> " & rs_compra_det!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
        If rs_aux7.RecordCount > 0 Then
            sino = MsgBox("Este Item ya Existe en esta Compra", vbCritical, "SOFIA")
            rs_compra_det.Cancel
            Exit Sub
        End If
      
     Else
        rs_compra_det.Open "Select * from ao_compra_detalle WHERE compra_codigo = " & fw_compras_comex.Ado_datos.Recordset!compra_codigo & " AND compra_codigo_det = " & fw_compras_comex.Ado_detalle1.Recordset!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
      
        Set rs_aux7 = New ADODB.Recordset
        If rs_aux7.State = 1 Then rs_aux7.Close
        rs_aux7.Open "Select * from ao_compra_detalle WHERE compra_codigo = " & fw_compras_comex.Ado_datos.Recordset!compra_codigo & " AND bien_codigo = '" & dtc_codigo1.Text & "' AND compra_codigo_det <> " & fw_compras_comex.Ado_detalle1.Recordset!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
        If rs_aux7.RecordCount > 0 Then
            sino = MsgBox("Este Item Ya Existe En Esta Compra", vbCritical, "SOFIA")
            'rs_compra_det.Cancel
            Exit Sub
        End If
     End If
     
     Dim correlativo As Integer
     'correlativo = 1
'    Set rs_correl = New Recordset
'    If rs_correl.State = 1 Then rs_correl.Close
'    rs_correl.Open "Select MAX(compra_codigo_det) AS CORREL from ao_compra_detalle WHERE compra_codigo = " & rs_compra_cab!compra_codigo & ", db, adOpenStatic"
'    If rs_correl!CORREL <> "NULL" Then
'    correlativo = rs_correl!CORREL + 1
'    Else
'    correlativo = "1"
'    End If
    
    
    'rs_aux2.MoveFirst
      'While Not rs_aux2.EOF
      
        rs_compra_det!ges_gestion = fw_compras_comex.Ado_datos.Recordset!ges_gestion
        rs_compra_det!compra_codigo = fw_compras_comex.Ado_datos.Recordset!compra_codigo
        
        rs_compra_det!bien_codigo = dtc_codigo1.Text
        rs_compra_det!compra_cantidad = Txt_campo16.Text
        rs_compra_det!compra_precio_unitario_bs = Txt_campo10.Text
        rs_compra_det!compra_precio_unitario_dol = CDbl(Txt_campo10.Text) / GlTipoCambioOficial
        rs_compra_det!compra_descuento_bs = CDbl(Txt_campo10.Text) * 0.87
        rs_compra_det!compra_descuento_dol = CDbl(rs_compra_det!compra_precio_unitario_dol) * 0.87
        rs_compra_det!compra_precio_total_bs = Txt_campo11.Text
        rs_compra_det!compra_precio_total_dol = Text2.Text
        rs_compra_det!compra_concepto = dtc_desc1.Text            'Txt_campo4.Text
        rs_compra_det!grupo_codigo = dtc_aux1.Text
        rs_compra_det!subgrupo_codigo = dtc_aux1.Text
        rs_compra_det!par_codigo = dtc_aux3.Text
        'rs_compra_det!almacen_codigo = "0"
        rs_compra_det!unimed_codigo = IIf(dtc_codigo2.Text = "", Txt_campo14.Text, dtc_codigo2.Text)
        'bien_descripcion
        rs_compra_det!almacen_codigo = IIf(dtc_cod_alm.Text = "", "15", dtc_cod_alm.Text)
        rs_compra_det!compra_tdc = IIf(Txt_tdc.Text = "0", GlTipoCambioOficial, CDbl(Txt_tdc))
        rs_compra_det!tipo_moneda = IIf(cmd_moneda.Text = "", "BOB", cmd_moneda.Text)

        rs_compra_det!usr_usuario = glusuario
        rs_compra_det!fecha_registro = Date
        rs_compra_det!adjudica_monto_bs_87 = CDbl(Txt_campo11.Text) * 0.87
        rs_compra_det.Update
'        rs_aux2.MoveNext
        'fw_compras_comex.ABRIR_TABLA_DET
      'Wend
        'fw_compras_comex.Ado_detalle2.Update
   End If

        'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
        'aw_p_ao_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
'        fw_compras_comex.Ado_detalle2.Recordset("fecha_registro").Value = Date
'        'aw_p_ao_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
'        fw_compras_comex.Ado_detalle2.Recordset("usr_codigo").Value = glusuario
'        fw_compras_comex.Ado_detalle2.Recordset.UpdateBatch adAffectAll
     
'     Set rs_aux1 = New ADODB.Recordset
'     SQL_FOR = "select * from ao_solicitud_edificacion where unidad_codigo = '" & aw_p_ao_solicitud.Ado_datos.Recordset("unidad_codigo") & "' and solicitud_codigo = " & aw_p_ao_solicitud.Ado_datos.Recordset("solicitud_codigo") & " and edif_codigo = '" & dtc_codigo1.Text & "'  "
'     rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'     If rs_aux1.RecordCount > 0 Then
'        MsgBox "El c�digo ya existe, consulte con el administrador del Sistema..."
'        var_cod = 0
'        Exit Sub
'     Else
'        aw_p_ao_solicitud.Ado_detalle1.Recordset("edif_codigo").Value = dtc_codigo1.Text
'     End If
     
     
'     var_cod = aw_p_ao_solicitud.Ado_detalle1.Recordset.RecordCount
'     db.Execute "Update ao_solicitud Set correl_edificacion = " & var_cod & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "  "
'    If lbl_det = "43340" Then
'     'Graba en Cotiza    1
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        SQL_FOR = "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " AND bien_codigo = '" & dtc_codigo1.Text & "'    "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If swnuevo = 1 Then
'            'Call cotiza_codigo
'            Set rs_aux5 = New ADODB.Recordset
'            If rs_aux5.State = 1 Then rs_aux5.Close
'            rs_aux5.Open "Select max(cotiza_codigo) as Codigo from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    ", db, adOpenStatic
'            If Not rs_aux5.EOF Then
'                 var_cod5 = IIf(IsNull(rs_aux5!Codigo), 1, rs_aux5!Codigo + 1)
'            End If
'            rs_aux4.AddNew
'            rs_aux4!ges_gestion = Year(Date)
'            rs_aux4!unidad_codigo = Txt_campo1.Caption
'            rs_aux4!solicitud_codigo = txt_codigo.Caption
'            rs_aux4!edif_codigo = frm_to_identificacion_cliente.dtc_codigo3.Text
'            rs_aux4!trafico_codigo = "0"  'Ado_datos.Recordset!trafico_codigo
'            rs_aux4!cotiza_codigo = var_cod5
'            rs_aux4!pais_continente = "NN"
'            'Call correl_bien
'            rs_aux4!bien_codigo = IIf(dtc_codigo1.Text = "", Txt_campo5.Text, dtc_codigo1.Text) '"MAN-002"       'CODIGO Servicio de Mantenimeitno
'            rs_aux4!proceso_codigo = "TEC"
'            rs_aux4!subproceso_codigo = "TEC-02"
'            rs_aux4!etapa_codigo = "TEC-02-01"
'            rs_aux4!poa_codigo = "3.2.3"
'            rs_aux4!clasif_codigo = "TEC"
'            rs_aux4!doc_codigo = "R-362"        'OJO - CAMBIAR R-xxx   OJO 28-DIC-2014
'            rs_aux4!doc_numero = "0"
'            rs_aux4!estado_codigo = "APR"
'
'            rs_aux4!modelo_codigo = Txt_campo3.Text     'Ado_datos.Recordset!modelo_codigo
'            rs_aux4!modelo_codigo_h = "0"        'Ado_datos.Recordset!modelo_codigo_h1
'            rs_aux4!modelo_codigo_x = "0"       'Ado_datos.Recordset!modelo_codigo_x1
'            rs_aux4!cotiza_fecha = Date
'            rs_aux4!cotiza_cantidad = IIf(Txt_campo16 = "", 1, Txt_campo16)
'            rs_aux4!cotiza_tdc_bol = GlTipoCambioOficial
'            rs_aux4!cotiza_precio_fob_bs = IIf(Txt_campo10 = "", 0, Txt_campo10)
'            rs_aux4!cotiza_precio_fob_dol = CDbl(Txt_campo10) * GlTipoCambioOficial
'            rs_aux4!cotiza_precio_total_bs = IIf(Txt_campo11 = "", 0, Txt_campo11)
'            rs_aux4!cotiza_precio_total_dol = CDbl(Txt_campo11) * GlTipoCambioOficial
'            rs_aux4!costo_monto = IIf(Txt_campo11 = "", 0, Txt_campo11)
'            rs_aux4!fecha_registro = Date
'            rs_aux4!usr_codigo = glusuario
'            rs_aux4.Update
'        Else
'            db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_cantidad = " & CDbl(Txt_campo16) & ", cotiza_precio_fob_bs = " & CDbl(Txt_campo10.Text) & ", cotiza_precio_total_bs = " & CDbl(Txt_campo11.Caption) & ", costo_monto = " & CDbl(Txt_campo11.Caption) & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & " and bien_codigo = '" & dtc_codigo1.Text & "'    "
'        End If
'
'        If swnuevo = 1 Then
'            db.Execute "Update ao_solicitud Set correl_cotiza = " & var_cod5 & " Where unidad_codigo = '" & Txt_campo1.Caption & "' and solicitud_codigo = " & txt_codigo.Caption & "    "
'        End If
'    End If
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
'  End If
    End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'dtc_cod_alm
  If Text2.Text = "0" Then
    MsgBox "Debe registrar el " + Label1.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If Txt_campo11.Text = "0" Then
    MsgBox "Debe registrar el " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar el " + lbl_codigo1.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo16.Text = "" Or Txt_campo16.Text = "0") Then
    MsgBox "Debe registrar ..." + lbl_campo16.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_tdc.Text = "" Then
    MsgBox "Debe registrar el " + Lbl_TDC.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_cod_alm.Text = "") Then
'    MsgBox "Debe registrar ... el ALMACEN", vbCritical + vbExclamation, "Validaci�n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
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
'        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atenci�n"
'    '    db.RollbackTrans
'        Screen.MousePointer = vbDefault
'    End If
End Sub

Private Sub cmd_moneda_LostFocus()
    Text2.Visible = True
    Txt_campo11.Visible = True
    If cmd_moneda.Text = "BOB" Then
        Txt_campo11.Enabled = True
        Txt_campo11.Visible = True
        Text2.Enabled = False
        'lbl_moneda.Caption = "Bolivianos"
    Else
        Txt_campo11.Enabled = False
        Txt_campo11.Visible = False
        Text2.Enabled = True
        'lbl_moneda.Caption = "Dolares"
    End If
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
    Txt_campo14.BoundText = dtc_aux1.BoundText
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
    Txt_campo14.BoundText = Dtc_aux2.BoundText
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
    Txt_campo14.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_cod_alm_Click(Area As Integer)
    dtc_desc_alm.BoundText = dtc_cod_alm.BoundText
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
    dtc_codigo2.BoundText = dtc_codigo1.BoundText
    Txt_campo14.BoundText = dtc_codigo1.BoundText
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

Private Sub dtc_codigo2_Change()
dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_desc_alm_Change()
    dtc_cod_alm.BoundText = dtc_desc_alm.BoundText
End Sub

Private Sub dtc_desc1_Change()
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Dtc_aux2.BoundText = dtc_desc1.BoundText
    dtc_aux3.BoundText = dtc_desc1.BoundText
    Txt_campo2.BoundText = dtc_desc1.BoundText
    Txt_campo3.BoundText = dtc_desc1.BoundText
    Txt_campo4.BoundText = dtc_desc1.BoundText
    Txt_campo18.BoundText = dtc_desc1.BoundText
    Txt_campo14.BoundText = dtc_desc1.BoundText
'    dtc_codigo2.BoundText = dtc_desc1.BoundText
'    dtc_desc2.BoundText = dtc_codigo2.Text
    
End Sub

Private Sub dtc_desc1_LostFocus()
    If dtc_aux1.Text = "20000" Then
        fra_almacen.Visible = False
        Select Case parametro
            Case "UALMR"
                dtc_cod_alm = "10"
            Case "ALMRB"
                dtc_cod_alm = "22"
            Case "ALMRS"
                dtc_cod_alm = "23"
            Case "ALMRC"     'REPUESTOS
                dtc_cod_alm = "26"
            Case Else
                dtc_cod_alm = "1"
        End Select
        dtc_desc_alm.BoundText = dtc_cod_alm.BoundText
        Txt_campo18.Visible = False
        Txt_campo2.Visible = False
        Txt_campo3.Visible = False
        Label11.Visible = False
        lbl_campo2.Visible = False
        lbl_campo5.Visible = False
    Else
        fra_almacen.Visible = True
        Txt_campo18.Visible = True
        Txt_campo2.Visible = True
        Txt_campo3.Visible = True
        Label11.Visible = True
        lbl_campo2.Visible = True
        lbl_campo5.Visible = True
    End If
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
    'dtc_codigo2.BoundText = dtc_desc2.BoundText
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
    Txt_campo11.Text = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
End Sub

Private Sub Form_Activate()
    mbDataChanged = False
    var_val2 = "2"
    If lbl_det = "43340" Then
        'Label1.Caption = "DETALLE DE BIENES (Equipos)"
        If Txt_campo16.Text = "0" Or Txt_campo16.Text = "" Then
            Txt_campo16.Text = "12"
        End If
    Else
        'Label1.Caption = "DETALLE DE BIENES "
    End If
'    If fw_compras_comex.Ado_detalle1!compra_tdc < 2 Then
        Txt_tdc.Text = "6.96"
'    Else
'        Txt_tdc.Text = fw_compras_comex.Ado_detalle1!compra_tdc
'    End If
    Text2.Visible = False
    Txt_campo11.Visible = False
'    Set rs_unidad = New ADODB.Recordset
'    If rs_unidad.State = 1 Then rs_unidad.Close
'    rs_unidad.Open "Select * from gc_unidad_ejecutora = '" & fw_compras_comex.VAR_UNI & "' order by unidad_descripcion", db, adOpenStatic
'    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_unidad
'    Txt_descripcion.Caption = rs_unidad!unidad_descripcion
End Sub

Private Sub Form_Load()
    CAL_DOL = "NO"
    CAL_BS = "NO"

    'VAR_UNI = parametro
    txt_codigo.Caption = fw_compras_comex.txt_codigo.Caption
    Set rs_UNIDAD = New ADODB.Recordset
    If rs_UNIDAD.State = 1 Then rs_UNIDAD.Close
    rs_UNIDAD.Open "Select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "' order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    'Set Ado_datos1.Recordset = rs_UNIDAD
    If rs_UNIDAD.RecordCount > 0 Then
        Txt_descripcion.Caption = rs_UNIDAD!unidad_descripcion
        'RESPONSABLE ALMACEN
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
        If rs_aux3.RecordCount > 0 Then
            VAR_BENEF = rs_aux3!beneficiario_codigo
            'VAR_DPTO = rs_aux3!depto_codigo
        Else
            VAR_BENEF = "0"
            'VAR_DPTO = "2"
        End If

        'If parametro = "DCONT" Then
        '    VAR_BENEF = "4999270"
        'Else
        '    VAR_BENEF = fw_compras_comex.dtc_codigo11.Text
        'End If
    End If
    
    Call ABRIR_TABLA
    'Call ABRIR_TABLA
    mbDataChanged = False
    Frame1.Visible = True
'    Frame2.Visible = False
    var_val2 = "2"
    
If parametro <> "COMEX" Then
    Txt_campo10.Visible = True
    lbl_campo10.Visible = True
    Set rs_clasif6 = New ADODB.Recordset
    If rs_clasif6.State = 1 Then rs_clasif6.Close
    'Select Case Glaux
    ' VERIFICAR CON REPUESTOS       JQA-2021
    'rs_clasif6.Open "SELECT * FROM ac_almacenes where beneficiario_codigo = '" & VAR_BENEF & "' ORDER BY almacen_descripcion ", db, adOpenStatic
    rs_clasif6.Open "SELECT * FROM ac_almacenes where almacen_codigo = 1", db, adOpenStatic
     Set Ado_clasif6.Recordset = rs_clasif6
     If rs_clasif6.RecordCount = 0 Then
     End If
     dtc_desc_alm.Enabled = True
     dtc_desc1.Enabled = True
     dtc_desc1.backColor = &HFFFFFF
     Text1(5).Visible = False
     fra_almacen.Visible = True
Else
     Txt_campo10.Visible = False
     lbl_campo10.Visible = False
     Set rs_clasif6 = New ADODB.Recordset
     If rs_clasif6.State = 1 Then rs_clasif6.Close
    'Select Case Glaux
     rs_clasif6.Open "SELECT * FROM ac_almacenes where almacen_codigo = 1", db, adOpenStatic
     Set Ado_clasif6.Recordset = rs_clasif6
     dtc_desc_alm.Enabled = False
     dtc_desc_alm.BoundText = rs_clasif6!almacen_codigo
     dtc_cod_alm.Text = rs_clasif6!almacen_codigo
     dtc_desc1.Locked = False
     dtc_desc1.backColor = &HFFFFFF
     Text1(5).Visible = False
     fra_almacen.Visible = False
End If
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
    'Bienes por almacen
        'ac_bienes
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    Select Case parametro
        Case "UALMI", "ALMIB", "ALMIS", "ALMIC"     'INSUMOS
            VAR_DET = " WHERE almacen_tipo = 'I' "
        Case "UALMR", "ALMRB", "ALMRS", "ALMRC"     'REPUESTOS
            VAR_DET = " where almacen_tipo = 'R' "
        Case "UALMH", "ALMHB", "ALMHS", "ALMHC"     'HERRAMIENTAS
            VAR_DET = " where almacen_tipo = 'H' "
        Case "DCONT"    'ALMACEN GENERAL
            VAR_DET = " where almacen_tipo = 'A' OR almacen_tipo = 'S' "
        Case "COMEX"    'ALMACEN GENERAL
            'VAR_DET = " where almacen_tipo = 'Q' "
            Select Case Glaux
                Case "PROVI"
                    VAR_DET = " where (almacen_tipo = 'S' AND bien_codigo_anterior = 'BANCO') or (almacen_tipo = 'R') "
                    'VAR_DET = VAR_DET & " AND bien_codigo_anterior = 'BANCO'"
                    'VAR_DET = VAR_DET & " AND bien_codigo_anterior = "
                Case "TRANS"
                      VAR_DET = " where almacen_tipo = 'S' "
                      VAR_DET = VAR_DET & " AND bien_codigo_anterior = 'TRANS'"
                Case "ADUAN"
                      VAR_DET = " where almacen_tipo = 'S' "
                      VAR_DET = VAR_DET & " AND bien_codigo_anterior = 'ADUAN'"
                Case "DESCA"
                     VAR_DET = " where almacen_tipo = 'S' "
                     VAR_DET = VAR_DET & " AND bien_codigo_anterior = 'DESCA'"
                Case "CONTR"
                     VAR_DET = " where almacen_tipo = 'S' "
                     VAR_DET = VAR_DET & " AND bien_codigo_anterior = 'CONTR'"
            End Select
        Case Else
            VAR_DET = " where almacen_tipo = 'A' "
    End Select
    rs_datos1.Open "select * from ac_bienes " & VAR_DET & " AND estado_codigo = 'APR' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly    'order by descripcion
    
    Set Ado_datos1.Recordset = rs_datos1
    sino = rs_datos1.RecordCount
    If swnuevo = 2 Then
        dtc_codigo1.Text = Txt_campo5.Caption
    End If
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    Dtc_aux2.BoundText = dtc_codigo1.BoundText
    dtc_aux3.BoundText = dtc_codigo1.BoundText
    Txt_campo2.BoundText = dtc_codigo1.BoundText
    Txt_campo3.BoundText = dtc_codigo1.BoundText
    Txt_campo4.BoundText = dtc_codigo1.BoundText
    Txt_campo18.BoundText = dtc_codigo1.BoundText
    Txt_campo14.BoundText = dtc_codigo1.BoundText
    'ac_bienes_unidad_medida
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
'    If lbl_det = "43340" Then
'        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo = 'T' ", db, adOpenStatic   'order by descripcion
'    Else
'        rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' ", db, adOpenStatic   'order by descripcion
'    End If
    rs_datos2.Open "select * from ac_bienes_unidad_medida where unimed_tipo <> 'T' ", db, adOpenStatic   'order by descripcion
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

Private Sub Text2_Change()
'
' On Error GoTo UpdateErr
'If CAL_DOL = "SI" Then
'    If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then
'        Txt_campo16.Text = "1"
'    End If
'    '
'    If Text2.Text <> "" And Text2.Text <> "," Then
'        'Txt_campo11.Text = Format(Round(CDbl(Text2.Text) * GlTipoCambioOficial, 2), "###,###,##0.00")
'        Txt_campo11.Text = Format(Round(CDbl(Text2.Text) * IIf(Txt_tdc.Text = "0", GlTipoCambioOficial, CDbl(Txt_tdc.Text)), 2), "###,###,##0.00")
'        Txt_campo10.Text = Format(Round(CDbl(Txt_campo11) / CDbl(Txt_campo16), 2), "###,###,##0.00")
'    Else
'        Txt_campo11.Text = "0"
'    End If
'End If
'CAL_DOL = "NO"
'CAL_BS = "NO"
'Exit Sub
'UpdateErr:
'  MsgBox Err.Description
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
CAL_DOL = "SI"
CAL_BS = "NO"
If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Text2_LostFocus()
 On Error GoTo UpdateErr
'If CAL_DOL = "SI" Then
If cmd_moneda = "USD" Then
    If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then
        Txt_campo16.Text = "1"
    End If
    '
    If Text2.Text <> "" And Text2.Text <> "," Then
        'Txt_campo11.Text = Format(Round(CDbl(Text2.Text) * GlTipoCambioOficial, 2), "###,###,##0.00")
        Txt_campo11.Text = Format(Round(CDbl(Text2.Text) * IIf(Txt_tdc.Text = "0", GlTipoCambioOficial, CDbl(Txt_tdc.Text)), 2), "###,###,##0.00")
        Txt_campo10.Text = Format(Round(CDbl(Txt_campo11) / CDbl(Txt_campo16), 2), "###,###,##0.00")
    Else
        Txt_campo11.Text = "0"
    End If
End If
CAL_DOL = "NO"
CAL_BS = "NO"
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Txt_campo10_Change()
'If Txt_campo16.Text = "" Then
'        Txt_campo16.Text = "1"
'
'    End If
'    If Txt_campo10.Text <> "" Then
'    Txt_campo11.Text = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
'    End If
End Sub

Private Sub Txt_campo10_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
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
     If Txt_campo10.Text <> "" And Txt_campo10.Text <> "." Then
    Txt_campo11.Text = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
    End If
End Sub

Private Sub Txt_campo11_Change()
'On Error GoTo UpdateErr
'If CAL_BS = "SI" Then
'    If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then
'        Txt_campo16.Text = "1"
'    End If
'
'    If Txt_campo11.Text <> "" And Txt_campo11.Text <> "," Then
'        Txt_campo10.Text = Format(Round(CDbl(Txt_campo11) / CDbl(Txt_campo16), 2), "###,###,##0.00")
'        Text2.Text = Format(Round(CDbl(Txt_campo11.Text) / IIf(Txt_tdc.Text = "0", GlTipoCambioOficial, CDbl(Txt_tdc.Text)), 2), "###,###,##0.00")
'    Else
'        Text2.Text = "0"
'    End If
'End If
'CAL_DOL = "NO"
'CAL_BS = "NO"
'
'Exit Sub
'UpdateErr:
'  MsgBox Err.Description
End Sub


Private Sub Txt_campo11_KeyPress(KeyAscii As Integer)
CAL_DOL = "NO"
CAL_BS = "SI"

  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txt_campo11_LostFocus()
On Error GoTo UpdateErr
'If CAL_BS = "SI" Then
If cmd_moneda = "BOB" Then
    If Txt_campo16.Text = "" Or Txt_campo16.Text = "0" Then
        Txt_campo16.Text = "1"
    End If
    
    If Txt_campo11.Text <> "" And Txt_campo11.Text <> "," Then
        Txt_campo10.Text = Format(Round(CDbl(Txt_campo11) / CDbl(Txt_campo16), 2), "###,###,##0.00")
        Text2.Text = Format(Round(CDbl(Txt_campo11.Text) / IIf(Txt_tdc.Text = "0", GlTipoCambioOficial, CDbl(Txt_tdc.Text)), 2), "###,###,##0.00")
    Else
        Text2.Text = "0"
    End If
End If
CAL_DOL = "NO"
CAL_BS = "NO"

Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Txt_campo14_Click(Area As Integer)
    dtc_codigo1.BoundText = Txt_campo14.BoundText
    dtc_desc1.BoundText = Txt_campo14.BoundText
    Dtc_aux2.BoundText = Txt_campo14.BoundText
    dtc_aux3.BoundText = Txt_campo14.BoundText
    Txt_campo2.BoundText = Txt_campo14.BoundText
    dtc_aux1.BoundText = Txt_campo14.BoundText
    Txt_campo4.BoundText = Txt_campo14.BoundText
    Txt_campo3.BoundText = Txt_campo14.BoundText
    Txt_campo18.BoundText = Txt_campo14.BoundText
End Sub

Private Sub Txt_campo16_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'  '? . , 09
'  ',.01234856789

If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
Exit Sub
Else
KeyAscii = 0
End If

End Sub


Private Sub Txt_campo16_Change()

 On Error GoTo UpdateErr
'     If Txt_campo10.Text = "" Then
'        Txt_campo10.Text = "1"
'    End If
'    If Txt_campo16.Text <> "" Then
'    Txt_campo11.Text = CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text)
'    End If

'    If Txt_campo16.Text <> "" And Txt_campo16.Text <> "." Then
'    If dtc_codigo1.Text = "479" Then
'       Txt_campo11.Caption = Round(CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text), 0)
'       Else
'       Txt_campo11.Caption = Round(CDbl(Txt_campo10.Text) * CDbl(Txt_campo16.Text), 2)
'    End If
'    End If

If Txt_campo11.Text = "" Then
Txt_campo11.Text = "0"
End If
If Txt_campo16.Text <> "" And Txt_campo16.Text <> "." Then
Txt_campo10.Text = Round(CDbl(Txt_campo11) / CDbl(Txt_campo16), 2)
End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
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
    Txt_campo14.BoundText = Txt_campo18.BoundText
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
    Txt_campo14.BoundText = Txt_campo2.BoundText
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
    Txt_campo14.BoundText = Txt_campo3.BoundText
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
    Txt_campo14.BoundText = Txt_campo4.BoundText
End Sub


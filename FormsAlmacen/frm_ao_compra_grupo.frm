VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ao_compra_grupo 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Administrativos - Compras Grupales"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frm_ao_compra_grupo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   120
      Picture         =   "frm_ao_compra_grupo.frx":0A02
      ScaleHeight     =   1425
      ScaleWidth      =   1875
      TabIndex        =   47
      Top             =   4680
      Width           =   1935
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cotiza"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_compra_grupo.frx":6AE20
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   740
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_compra_grupo.frx":6C5A2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Adiciona Producto"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_compra_grupo.frx":6C9E4
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_compra_grupo.frx":6CE26
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   740
         Width           =   765
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "BIEN A COTIZAR"
      ForeColor       =   &H00FFFFC0&
      Height          =   1935
      Left            =   2160
      TabIndex        =   45
      Top             =   4560
      Width           =   12855
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "frm_ao_compra_grupo.frx":6D268
         Height          =   1305
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   2302
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "grupo_codigo"
            Caption         =   "Grupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "subgrupo_codigo"
            Caption         =   "Sub-Grupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien/Serv"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "compra_concepto"
            Caption         =   "Descripcion Bien o Servicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "compra_cantidad"
            Caption         =   "Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "compra_precio_unitario_bs"
            Caption         =   "Precio.BOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "compra_precio_total_bs"
            Caption         =   "Precio.Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "compra_precio_unitario_dol"
            Caption         =   "Precio.USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "compra_precio_total_dol"
            Caption         =   "Total USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "tipo_eqp_descripcion"
            Caption         =   "Tipo.Bien/Serv."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "marca_descripcion"
            Caption         =   "Marca"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4575.118
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1604.976
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Caption         =   "Elije una o mas Opciones para Agrupar la Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3360
      Left            =   6120
      TabIndex        =   19
      Top             =   1200
      Width           =   8895
      Begin VB.ComboBox Txt02 
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_ao_compra_grupo.frx":6D283
         Left            =   180
         List            =   "frm_ao_compra_grupo.frx":6D2AB
         TabIndex        =   44
         Text            =   "ENERO"
         Top             =   2100
         Width           =   1695
      End
      Begin VB.ComboBox TxtGestion 
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_ao_compra_grupo.frx":6D314
         Left            =   1680
         List            =   "frm_ao_compra_grupo.frx":6D32A
         TabIndex        =   43
         Text            =   "2015"
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "compra_descripcion_grupo"
         DataSource      =   "Ado_datos"
         Height          =   435
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2715
         Width           =   7425
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   6960
         TabIndex        =   20
         Top             =   1005
         Visible         =   0   'False
         Width           =   290
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_ao_compra_grupo.frx":6D352
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5880
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "fecha_proceso"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7185
         TabIndex        =   22
         Top             =   2100
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   90832897
         CurrentDate     =   41678
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_ao_compra_grupo.frx":6D36B
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_ao_compra_grupo.frx":6D384
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_ao_compra_grupo.frx":6D39D
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   1275
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_ao_compra_grupo.frx":6D3B6
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4620
         TabIndex        =   27
         Top             =   1275
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_compra_grupo.frx":6D3CF
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_compra_grupo.frx":6D3E8
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2895
         TabIndex        =   29
         Top             =   600
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_ao_compra_grupo.frx":6D401
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2100
         TabIndex        =   30
         Top             =   2100
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   7875
         TabIndex        =   42
         Top             =   340
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Correlativo"
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
         Left            =   180
         TabIndex        =   41
         Top             =   340
         Width           =   975
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7785
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Registro"
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
         Index           =   12
         Left            =   7185
         TabIndex        =   39
         Top             =   1830
         Width           =   1380
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "grupo_compra"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   180
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Solicitante"
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
         Left            =   2925
         TabIndex        =   37
         Top             =   345
         Width           =   1755
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cliente"
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
         Left            =   4620
         TabIndex        =   36
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Mes "
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
         Left            =   180
         TabIndex        =   35
         Top             =   1830
         Width           =   435
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Concepto:"
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
         Left            =   180
         TabIndex        =   34
         Top             =   2730
         Width           =   915
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Edificio del Cliente"
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
         Left            =   180
         TabIndex        =   33
         Top             =   1005
         Width           =   1650
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gestión"
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
         Left            =   1710
         TabIndex        =   32
         Top             =   345
         Width           =   690
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Bien (Insumo)"
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
         Left            =   2160
         TabIndex        =   31
         Top             =   1830
         Width           =   1215
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   3360
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   5895
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   3600
         TabIndex        =   18
         Top             =   2985
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1320
         TabIndex        =   17
         Top             =   2985
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   2610
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4604
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "grupo_compra"
            Caption         =   "Correlativo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "unidad_codigo"
            Caption         =   "U.Solicitante"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "fecha_proceso"
            Caption         =   "Fecha.Reg."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "bien_codigo"
            Caption         =   "Bien.(Insumo)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2925
         Width           =   5625
         _ExtentX        =   9922
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
         Caption         =   ""
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
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_ao_compra_grupo.frx":6D41A
      ScaleHeight     =   960
      ScaleWidth      =   14835
      TabIndex        =   4
      Top             =   120
      Width           =   14900
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_ao_compra_grupo.frx":D944C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ao_compra_grupo.frx":D988E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_ao_compra_grupo.frx":D9A98
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_ao_compra_grupo.frx":DA050
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_ao_compra_grupo.frx":DA60D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_ao_compra_grupo.frx":DA817
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_ao_compra_grupo.frx":DB4E1
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_ao_compra_grupo.frx":DBAC1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_ao_compra_grupo.frx":DC0E5
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO1"
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
         Left            =   10080
         TabIndex        =   14
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_ao_compra_grupo.frx":DC2EF
      ScaleHeight     =   915
      ScaleWidth      =   14835
      TabIndex        =   0
      Top             =   120
      Width           =   14900
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_ao_compra_grupo.frx":148321
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_ao_compra_grupo.frx":14852B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   9900
         TabIndex        =   3
         Top             =   300
         Width           =   1305
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   8880
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
      Left            =   2280
      Top             =   8880
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
      Left            =   4440
      Top             =   8880
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
      Left            =   6720
      Top             =   8880
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
Attribute VB_Name = "frm_ao_compra_grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset

'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_DET As String

Dim VAR_REG As Integer
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant

Private Sub BtnAñadir_Click()
  On Error GoTo AddErr
    VAR_SW = "ADD"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    
    Ado_datos.Recordset.AddNew
    dtc_desc1.SetFocus
    
    'dtc_codigo1.Text = parametro
'    Select Case parametro
'        Case "DVTA"        'INI COMERCIAL
'            dtc_codigo2.Text = 3
'        Case "COMEX"        'INI COMEX
'            dtc_codigo2.Text = 3
'        Case "DNINS"                        'INI GRABA INSTALACIONES
'            '
'            dtc_codigo2.Text = 4
'        Case "DNAJS"
'            '
'            dtc_codigo2.Text = 4
'        Case "DNMAN"
'            '
'            dtc_codigo2.Text = 4
'        Case Else
'            dtc_codigo2.Text = 5
'    End Select
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
  Exit Sub
AddErr:
  MsgBox Err.Description

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

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    parametro = Aux

    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True

    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
	Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from ao_compra_grupo where estado_codigo = 'REG' "
        Case "UALMR"    'REPUESTOS
            queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'"
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from ao_compra_grupo "
        Case "UALMR"    'REPUESTOS
            queryinicial = "Select * from av_solicitud_repuestos "
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_unidad_ejecutora WHERE estado_codigo = 'APR' order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "Select * from gv_unidad_solicitante_30000 estado_codigo = 'APR' order by unidad_descripcion", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'ac_bienes
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from ac_bienes WHERE (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800') order by bien_descripcion ", db, adOpenStatic
    'rs_datos2.Open "Select * from av_solicitud_lista_30000 WHERE (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800') order by bien_descripcion ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones WHERE estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "Select * from av_solicitud_edificio_30000 order by edif_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_beneficiario WHERE estado_codigo = 'APR' order by beneficiario_denominacion", db, adOpenStatic
    'rs_datos4.Open "Select * from av_solicitud_beneficiario_30000 order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
    If VAR_SW <> "ADD" Then
'        Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'                Call ABRIR_TABLA_DET
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'                Call ABRIR_TABLA_DET
'            Case "10"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'                Call ABRIR_TABLA_DET
'            Case "5"    ' SERVICIO MODERNIZACION
'        End Select
        Call ABRIR_TABLA_DET
        Call ABRIR_TABLAS_AUX
    Else
        'Set rs_det1 = New ADODB.Recordset
        Call ABRIR_TABLA_DET
'        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    'FraDet1.Caption = "BITÁCORA DE: " + dtc_desc1.Text
    If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet1.Visible = False
            FrmABMDet2.Visible = False
    Else
            FrmABMDet1.Visible = True
'            FrmABMDet2.Visible = True
    End If
  End If
End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            If TxtGestion.Text = "" Then
                TxtGestion.Text = "%"
            End If
            'rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and grupo_codigo = '30000' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1.Open "select * from av_compra_30000 where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and grupo_codigo = '30000' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            'select * from av_compra_30000 where ges_gestion like '" & TxtGestion & "' and  unidad_codigo like '" & dtc_codigo1 & "' and  edif_codigo like '" & dtc_codigo3 & "' and beneficiario_codigo like '" & dtc_codigo4 & "' and bien_codigo like '" && "'
            'grupo_mes like '" & & "',
        Case "UALMR"    'REPUESTOS
            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and (par_codigo = '39810' or par_codigo = '39820') ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "UALMH"    'HERRAMIENTAS
            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '34800'   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "UGADM"    'ADMINISTRATIVA
            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "    ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case Else
            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '43340'   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    End Select
    'rs_aux2.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and grupo_codigo = '30000' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset

'    Set rs_aux2 = New ADODB.Recordset
'    If rs_aux2.State = 1 Then rs_aux2.Close
'    'rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    rs_aux2.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle2.Recordset = rs_aux2
'    Set dg_det2.DataSource = Ado_detalle2.Recordset
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        var_cod = rs_datos!grupo_compra
        txt_codigo.Caption = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
     End If
     rs_datos!ges_gestion = TxtGestion.Text     ' glGestion
     rs_datos!unidad_codigo = dtc_codigo1
     rs_datos!unidad_codigo_adm = parametro
     rs_datos!fecha_proceso = DTPfecha1.Value
     rs_datos!edif_codigo = dtc_codigo3.Text
     rs_datos!beneficiario_codigo = dtc_codigo4.Text
     rs_datos!bien_codigo = dtc_codigo2.Text
     rs_datos!compra_descripcion = Txt_descripcion.Text
     
     'var_cod = rs_datos!grupo_compra

     Select Case Txt02.Text
        Case "ENERO"
            rs_datos!grupo_mes = 1
        Case "FEBRERO"
            rs_datos!grupo_mes = 2
        Case "MARZO"
            rs_datos!grupo_mes = 3
        Case "ABRIL"
            rs_datos!grupo_mes = 4
        Case "MAYO"
            rs_datos!grupo_mes = 5
        Case "JUNIO"
            rs_datos!grupo_mes = 6
        Case "JULIO"
            rs_datos!grupo_mes = 7
        Case "AGOSTO"
            rs_datos!grupo_mes = 8
        Case "SEPTIEMBRE"
            rs_datos!grupo_mes = 9
        Case "OCTUBRE"
            rs_datos!grupo_mes = 10
        Case "NOVIEMBRE"
            rs_datos!grupo_mes = 11
        Case "DICIEMBRE"
            rs_datos!grupo_mes = 12
     End Select
     'rs_datos!usr_codigo_aprueba = ""
     'rs_datos!fecha_registro_aprueba = Date
     
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
     Else
        Call OptFilGral2_Click
     End If
      
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     VAR_SW = ""
     
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo11.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
''  If (dtc_codigo8.Text = "") Then
''    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
''    VAR_VAL = "ERR"
''    Exit Sub
''  End If
''  If (dtc_codigo9.Text = "") Then
''    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
''    VAR_VAL = "ERR"
''    Exit Sub
''  End If
  If (TxtGestion.Text = "") Then
    MsgBox "Debe registrar la Gestión ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub


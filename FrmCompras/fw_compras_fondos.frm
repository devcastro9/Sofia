VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_compras_fondos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Fondos a Rendir"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "fw_compras_fondos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox BtnSalir 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   17640
      Picture         =   "fw_compras_fondos.frx":0A02
      ScaleHeight     =   615
      ScaleWidth      =   1245
      TabIndex        =   84
      ToolTipText     =   "Cierra la Ventana Activa"
      Top             =   360
      Width           =   1245
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   5280
      Left            =   4440
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox lbl_total_dol 
         DataField       =   "compra_monto_dol"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "0"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox lbl_total_bs 
         DataField       =   "compra_monto_bs"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3840
         TabIndex        =   92
         Text            =   "0"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   10080
         TabIndex        =   85
         Top             =   120
         Width           =   10080
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRO DE SOLICITUD DE FONDOS"
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
            Left            =   2775
            TabIndex        =   86
            Top             =   120
            Width           =   4545
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   10080
         TabIndex        =   58
         Top             =   4440
         Visible         =   0   'False
         Width           =   10080
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3480
            Picture         =   "fw_compras_fondos.frx":13C1
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   60
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5235
            Picture         =   "fw_compras_fondos.frx":1B97
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   59
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9810
         TabIndex        =   56
         Top             =   1290
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "compra_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   555
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   2880
         Width           =   8865
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7545
         TabIndex        =   16
         Top             =   975
         Width           =   290
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4860
         TabIndex        =   15
         Top             =   1290
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2505
         TabIndex        =   14
         Top             =   3555
         Visible         =   0   'False
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "fw_compras_fondos.frx":2483
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3480
         TabIndex        =   18
         Top             =   2160
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
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "fw_compras_fondos.frx":249D
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_sigla"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "fw_compras_fondos.frx":24B6
         DataField       =   "venta_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "venta_tipo"
         BoundColumn     =   "venta_tipo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "compra_fecha"
         Height          =   300
         Left            =   8505
         TabIndex        =   21
         Top             =   4020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   110362625
         CurrentDate     =   44934
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "fw_compras_fondos.frx":24CF
         DataField       =   "codigo_empresa"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4680
         TabIndex        =   22
         Top             =   2460
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "denominacion_empresa"
         BoundColumn     =   "codigo_empresa"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "fw_compras_fondos.frx":24E9
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   23
         Top             =   3960
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "fw_compras_fondos.frx":2502
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "fw_compras_fondos.frx":251B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "fw_compras_fondos.frx":2534
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Top             =   1635
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc_ben 
         Bindings        =   "fw_compras_fondos.frx":254D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   1635
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "fw_compras_fondos.frx":2566
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   29
         Top             =   600
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
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "fw_compras_fondos.frx":257F
         DataField       =   "venta_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1635
         TabIndex        =   30
         Top             =   3540
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "venta_tipo_descripcion"
         BoundColumn     =   "venta_tipo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "fw_compras_fondos.frx":2598
         DataField       =   "unidad_codigo_adm"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   31
         Top             =   960
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "fw_compras_fondos.frx":25B1
         DataField       =   "codigo_empresa"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "codigo_empresa"
         BoundColumn     =   "codigo_empresa"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "fw_compras_fondos.frx":25CB
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   2460
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label TxtCompra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "compra_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         TabIndex        =   91
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "#Compra"
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
         TabIndex        =   90
         Top             =   680
         Width           =   825
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Dol"
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
         Index           =   5
         Left            =   6255
         TabIndex        =   57
         Top             =   3675
         Width           =   825
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_numero"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9000
         TabIndex        =   52
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total Bs"
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
         Left            =   4035
         TabIndex        =   34
         Top             =   3675
         Width           =   750
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "#Tr?mite"
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
         Index           =   0
         Left            =   7980
         TabIndex        =   49
         Top             =   680
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   12
         Left            =   8625
         TabIndex        =   48
         Top             =   3675
         Width           =   1380
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa"
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
         Index           =   13
         Left            =   4680
         TabIndex        =   47
         Top             =   2190
         Width           =   825
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7920
         TabIndex        =   46
         Top             =   960
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10320
         Y1              =   2100
         Y2              =   2160
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   10320
         Y1              =   3555
         Y2              =   3600
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3165
         TabIndex        =   45
         Top             =   680
         Width           =   1560
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solicitante"
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
         Left            =   5340
         TabIndex        =   44
         Top             =   1365
         Width           =   930
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable del Proceso:"
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
         Left            =   180
         TabIndex        =   43
         Top             =   2190
         Width           =   2415
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registro ISO"
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
         Left            =   405
         TabIndex        =   42
         Top             =   3675
         Width           =   1140
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   41
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
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
         Left            =   180
         TabIndex        =   40
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO"
         DataField       =   "unidad_codigo_ant"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1350
         TabIndex        =   39
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cite.Tr?mite"
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
         Index           =   6
         Left            =   1695
         TabIndex        =   38
         Top             =   680
         Width           =   1080
      End
      Begin VB.Label dtc_codigo9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "doc_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   345
         TabIndex        =   37
         Top             =   4005
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "#Documento"
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
         Index           =   3
         Left            =   9000
         TabIndex        =   36
         Top             =   680
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "36NO"
         DataField       =   "compra_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2460
         TabIndex        =   35
         Top             =   3990
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4365
      Left            =   0
      TabIndex        =   12
      Top             =   4140
      Width           =   19215
      Begin VB.PictureBox BtnAddDetalle3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9000
         Picture         =   "fw_compras_fondos.frx":25E5
         ScaleHeight     =   735
         ScaleWidth      =   1080
         TabIndex        =   83
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.PictureBox fraOpcionesDet 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   8760
         TabIndex        =   73
         Top             =   240
         Width           =   8760
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5520
            Picture         =   "fw_compras_fondos.frx":31A4
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnAnlDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "fw_compras_fondos.frx":3A71
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   76
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnModDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1305
            Picture         =   "fw_compras_fondos.frx":41BD
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   75
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnAddDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_compras_fondos.frx":4AD2
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   74
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture2AA 
         FillColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   10200
         ScaleHeight     =   600
         ScaleWidth      =   8835
         TabIndex        =   51
         Top             =   240
         Width           =   8895
         Begin VB.PictureBox BtnAddDetalle4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_compras_fondos.frx":5291
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   94
            Top             =   20
            Width           =   1200
         End
         Begin VB.PictureBox BtnImprimir3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3720
            Picture         =   "fw_compras_fondos.frx":5A50
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "fw_compras_fondos.frx":637E
         Height          =   3180
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5609
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   13
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
            DataField       =   "bien_descripcion"
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
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         BeginProperty Column07 
            DataField       =   "compra_precio_total_bs"
            Caption         =   "Precio.Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "compra_precio_unitario_dol"
            Caption         =   "Precio.USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "compra_precio_total_dol"
            Caption         =   "Total USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
            DataField       =   "almacen_descripcion"
            Caption         =   "Almacen"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4754.835
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1964.976
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_det1A 
         Bindings        =   "fw_compras_fondos.frx":6399
         Height          =   3300
         Left            =   10200
         TabIndex        =   53
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5821
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   12
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
            DataField       =   "adjudica_cantidad"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         BeginProperty Column06 
            DataField       =   "bien_precio_adjudica_bs"
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
         BeginProperty Column07 
            DataField       =   "bien_total_adjudica_bs"
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
         BeginProperty Column08 
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
         BeginProperty Column09 
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4649.953
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton BtnAprobar3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "Envia a Almacen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   920
         Left            =   9000
         MaskColor       =   &H80000014&
         Picture         =   "fw_compras_fondos.frx":63B5
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "A?adir a proveedor"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnDesAprobar3 
         BackColor       =   &H80000018&
         Caption         =   "Retorna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9000
         Picture         =   "fw_compras_fondos.frx":67F7
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Quitar de Proveedor"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DESCARGOS - REGISTRO FACTURAS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3960
      Left            =   9360
      TabIndex        =   7
      Top             =   120
      Width           =   9855
      Begin VB.PictureBox FrmABMDet2 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   9600
         TabIndex        =   78
         Top             =   240
         Width           =   9600
         Begin VB.PictureBox BtnAprobar4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4080
            Picture         =   "fw_compras_fondos.frx":6C39
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   89
            Top             =   0
            Width           =   1320
         End
         Begin VB.CommandButton BtnImprimir4 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   6720
            Picture         =   "fw_compras_fondos.frx":7503
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Imprime Nota de Venta"
            Top             =   0
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.PictureBox BtnAprobar1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5400
            Picture         =   "fw_compras_fondos.frx":7E31
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   82
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnAnlDetalle2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "fw_compras_fondos.frx":8748
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   81
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnModDetalle2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1305
            Picture         =   "fw_compras_fondos.frx":8F40
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   80
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAddDetalle2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_compras_fondos.frx":9955
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   79
            Top             =   0
            Width           =   1200
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "fw_compras_fondos.frx":A205
         Height          =   2880
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5080
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc."
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
         BeginProperty Column01 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cod.Proveedor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "adjudica_descripcion"
            Caption         =   "Denominaci?n.Proveedor"
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
            DataField       =   "adjudica_monto_dol"
            Caption         =   "Importe.USD"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "adjudica_monto_bs"
            Caption         =   "Importe.BOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fecha_compra"
            Caption         =   "Fecha Factura"
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
            DataField       =   "fecha_inicio_contrato"
            Caption         =   "Fecha.Inicio"
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
         BeginProperty Column07 
            DataField       =   "fecha_fin_contrato"
            Caption         =   "Fecha.Finalizacion"
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
         BeginProperty Column08 
            DataField       =   "Fecha_envio_proveedor"
            Caption         =   "Fecha.Entrega/Salida"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4169.764
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3960
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9135
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   100
         ScaleHeight     =   660
         ScaleWidth      =   8940
         TabIndex        =   61
         Top             =   240
         Width           =   8940
         Begin VB.CommandButton BtnVer 
            BackColor       =   &H00808000&
            Caption         =   "Digitaliza"
            Height          =   600
            Left            =   10800
            Picture         =   "fw_compras_fondos.frx":A220
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   0
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H00808080&
            Height          =   600
            Left            =   11760
            Picture         =   "fw_compras_fondos.frx":A662
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.PictureBox BtnA?adir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "fw_compras_fondos.frx":A86C
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   69
            Top             =   0
            Width           =   1200
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1305
            Picture         =   "fw_compras_fondos.frx":B02B
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   68
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "fw_compras_fondos.frx":B940
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   67
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6960
            Picture         =   "fw_compras_fondos.frx":C08C
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4080
            Picture         =   "fw_compras_fondos.frx":C8BF
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   65
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5520
            Picture         =   "fw_compras_fondos.frx":D074
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   64
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnSalirA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   8520
            Picture         =   "fw_compras_fondos.frx":D941
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   63
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   255
            Left            =   8640
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl_titulo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CRONOGRAMA"
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
            Left            =   12855
            TabIndex        =   72
            Top             =   195
            Width           =   1815
         End
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   2415
         Left            =   105
         TabIndex        =   9
         Top             =   960
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "doc_numero"
            Caption         =   "Nro. Doc"
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
         BeginProperty Column01 
            DataField       =   "compra_codigo"
            Caption         =   "#Compra"
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
            DataField       =   "compra_fecha"
            Caption         =   "Fecha.Solicitud"
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
         BeginProperty Column03 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Tr?mite"
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
            DataField       =   "compra_descripcion"
            Caption         =   "Concepto"
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
         BeginProperty Column05 
            DataField       =   "estado_codigo_eqp"
            Caption         =   "Estado1"
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
         BeginProperty Column06 
            DataField       =   "unidad_codigo"
            Caption         =   "U.Ejecutora"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "#Tr?mite"
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
         BeginProperty Column09 
            DataField       =   "estado_codigo_tra"
            Caption         =   "Etapa2"
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
            DataField       =   "estado_codigo_nac"
            Caption         =   "Etapa3"
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
         BeginProperty Column11 
            DataField       =   "estado_codigo_des"
            Caption         =   "Etapa5"
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
         BeginProperty Column12 
            DataField       =   "estado_codigo"
            Caption         =   "Estado.Gral"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3809.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               ColumnWidth     =   1260.284
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   10
         Top             =   3585
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6480
         TabIndex        =   11
         Top             =   3600
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   105
         Top             =   3480
         Width           =   8955
         _ExtentX        =   15796
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
         BackColor       =   16777215
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   21360
      TabIndex        =   0
      Top             =   12915
      Width           =   21360
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
         Left            =   0
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9960
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
   Begin Crystal.CrystalReport CR01 
      Left            =   7200
      Top             =   11040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   9960
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
      Top             =   9960
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
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   9000
      Top             =   9960
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
      Caption         =   "Ado_datos5"
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   11280
      Top             =   9960
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
      Caption         =   "Ado_datos6"
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
      Left            =   13560
      Top             =   9960
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   120
      Top             =   10320
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
      Caption         =   "Ado_datos8"
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   2280
      Top             =   10320
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
      Caption         =   "Ado_datos9"
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   6360
      Top             =   10920
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
      Caption         =   "Ado_datos10"
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   120
      Top             =   10680
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
      Caption         =   "Ado_detalle1"
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2400
      Top             =   9600
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
      Caption         =   "Ado_detalle2"
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   4440
      Top             =   10320
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
      Caption         =   "Ado_datos11"
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
   Begin MSAdodcLib.Adodc Ado_datos12 
      Height          =   330
      Left            =   6720
      Top             =   10320
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
      Caption         =   "Ado_datos12"
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   4680
      Top             =   10680
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
      Caption         =   "Ado_detalle3"
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
   Begin Crystal.CrystalReport CR02 
      Left            =   9000
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_detalle1A 
      Height          =   330
      Left            =   7200
      Top             =   10680
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
      Caption         =   "Ado_detalle1A"
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
Attribute VB_Name = "fw_compras_fondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos  As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim i As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det1A As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod, DETALLE1, DETALLE2, DETALLE3 As String
Dim VAR_VAL, VAR_DOCF, VAR_CODF As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI, VAR_UORIGEN As String
Dim sino As String
Dim VAR_PAIS, VAR_BENEF, VAR_DA, VAR_TIPO_ALM As String
Public w_nuevo As String
Dim VAR_FONDO, VAR_TIPO As String

Dim COUNTER As Integer
Dim VAR_CMPBTE, VAR_COMPRA As Integer
Dim CORRELARTIVO1, CORRELATIVO2 As Integer

Dim VAR_AUX, VAR_CONT2, SUMbs, SUMdol, VAR_DPTO_AUX As Double
Dim VAR_FOBSEG, VAR_FOBSEG2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Public Function Literal(Cadena As String) As String
Dim SW As Integer
Dim sw1 As Integer
Dim swc, i As Integer
Dim VEC(20) As Long
SW = 0
      '*********PARTE DECIMAL*********
            If Cadena < 0 Then Cadena = Cadena * (-1)
            Cadena = Round(Cadena, 2)
             x = Len(Cadena)
              For k = 1 To x
                  Z = Mid(Cadena, k, 1)
                  If (Z = ".") Or SW = 1 Then
                    d = d + Mid(Cadena, k, 1)
                    SW = 1
                  End If
              Next k
              
              d = Mid(d, 2, Len(d))
              
              'Para la parte decimal del monto
              If d = "00" Or d = "" Then
                 d = d & " 00/100"
              Else
                 If d >= 0 And d <= 9 And Len(d) = 1 Then
                    d = " " & d & "0" & "/100"
                 Else
                    d = " " & d & "/100 "
                 End If
              End If
      '*********PARTE ENTERA*********
 If Cadena <> "" Then
    Cadena = Int(Cadena)
 Else
    MsgBox "No existe monto"
 End If
   s = ""
   Z = ""
   c = 0
   k = 0
   sw1 = 0
   swc = 0
   
   
   x = Len(Cadena)
   For i = 1 To x
       a = Mid(Cadena, i, 1)
       VEC(i) = Mid(Cadena, i, 1)
   Next i
j = x
While j <> 0
k = k + 1
If k <> 8 Then
  If c <> 3 Then
       c = c + 1
      
       If c = 1 And (VEC(j - 1) <> 1 And VEC(j - 1) <> 2) Then
            Select Case VEC(j)
                Case 0: s = " " + s
                Case 1:
                   If sw1 <> 1 Then
                      s = "UNO " + Z + s
                   End If
                   If sw1 = 1 Then
                      s = "UN " + Z + s
                   End If
                   
                Case 2: s = "DOS " + Z + s
                Case 3: s = "TRES " + Z + s
                Case 4: s = "CUATRO " + Z + s
                Case 5: s = "CINCO " + Z + s
                Case 6: s = "SEIS " + Z + s
                Case 7: s = "SIETE " + Z + s
                Case 8: s = "OCHO " + Z + s
                Case 9: s = "NUEVE " + Z + s
          End Select
          
           'If J + 1 <> "" And sw1 <> 1 And VEC(J - 1) <> 0 And VEC(J) <> 0 Then
           If VEC(j - 1) <> 0 And VEC(j) <> 0 Then
                 s = "Y " + s
           End If
           
        End If
        
         If c = 2 And VEC(j) = 1 Then
               swc = 1
                Select Case VEC(j + 1)
                      Case 0: s = "DIEZ " + Z + s
                      Case 1: s = "ONCE " + Z + s
                      Case 2: s = "DOCE " + Z + s
                      Case 3: s = "TRECE " + Z + s
                      Case 4: s = "CATORCE " + Z + s
                      Case 5: s = "QUINCE " + Z + s
                      Case 6: s = "DIECISEIS " + Z + s
                      Case 7: s = "DIECISIETE " + Z + s
                      Case 8: s = "DIECIOCHO " + Z + s
                      Case 9: s = "DIECINUEVE " + Z + s
                End Select
          End If
          
          If c = 2 And VEC(j) = 2 Then
                Select Case VEC(j + 1)
                      Case 0: s = "VEINTE " + Z + s
                      Case 1: s = "VEINTIUNO " + Z + s
                      Case 2: s = "VEINTIDOS " + Z + s
                      Case 3: s = "VEINTITRES " + Z + s
                      Case 4: s = "VEINTICUATRO " + Z + s
                      Case 5: s = "VEINTICINCO " + Z + s
                      Case 6: s = "VEINTISEIS " + Z + s
                      Case 7: s = "VEINTISIETE " + Z + s
                      Case 8: s = "VEINTIOCHO " + Z + s
                      Case 9: s = "VEINTINUEVE " + Z + s
                End Select
          End If
   
        If c = 2 Then
            Select Case VEC(j)
                Case 3: s = "TREINTA " + Z + s
                Case 4: s = "CUARENTA " + Z + s
                Case 5: s = "CINCUENTA " + Z + s
                Case 6: s = "SESENTA " + Z + s
                Case 7: s = "SETENTA " + Z + s
                Case 8: s = "OCHENTA " + Z + s
                Case 9: s = "NOVENTA " + Z + s
            End Select
            
        End If
        
        If c = 3 Then
            Select Case VEC(j)
                Case 1:
                If j = 1 Then
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                Else
                    If VEC(j + 1) = 0 And VEC(j + 2) = 0 Then
                       s = "CIEN " + Z + s
                    Else
                       s = "CIENTO " + Z + s
                    End If
                       'S = "CIENTO " + z + S
                End If
                Case 2: s = "DOSCIENTOS " + Z + s
                Case 3: s = "TRESCIENTOS " + Z + s
                Case 4: s = "CUATROCIENTOS " + Z + s
                Case 5: s = "QUINIENTOS " + Z + s
                Case 6: s = "SEISCIENTOS " + Z + s
                Case 7: s = "SETECIENTOS " + Z + s
                Case 8: s = "OCHOCIENTOS " + Z + s
                Case 9: s = "NOVECIENTOS " + Z + s
            End Select
        End If
   Else
     If j >= 3 Then
            If VEC(j) = 0 And VEC(j - 1) = 0 And VEC(j - 2) = 0 Then
            Else
              s = "MIL " + s
            End If
    Else
              s = "MIL " + s
    End If
        j = j + 1
        c = 0
        sw1 = 1
   End If
 Else
    If VEC(j) <> 1 Then
      s = "MILLONES " + s
    Else
'      If K > 7 Then
      If k <> 8 Then
        s = "MILLONES " + s
      Else
        s = "MILLON " + s
      End If
    End If
      j = j + 1
      c = 0
      sw1 = 1
 End If
   j = j - 1
   
Wend

Literal = s + d
End Function

Private Sub BIENES()
'   Set rs_compra_det = New Recordset
'     If rs_compra_det.State = 1 Then rs_compra_det.Close
'     rs_compra_det.Open "Select * from ao_compra_adjudica_bienes", db, adOpenKeyset, adLockOptimistic
''     If MOD_NEW.Caption = "NEW" Then
''
''     rs_compra_det.Open "Select * from ao_compra_adjudica_bienes", db, adOpenKeyset, adLockOptimistic
''     rs_compra_det.AddNew
''     rs_compra_det!compra_codigo_det = ao_compra_adjudica_bienes.Ado_detalle1.Recordset.RecordCount + 1
''     Else
''      rs_compra_det.Open "Select * from ao_compra_adjudica_bienes WHERE compra_codigo = " & fw_compras_fondos.Ado_datos.Recordset!compra_codigo & " AND compra_codigo_det = " & fw_compras_fondos.Ado_detalle1.Recordset!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
''
''     End If
    DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
    VAR_COD2 = Ado_datos.Recordset!compra_codigo
    CodBien = Ado_detalle1.Recordset!bien_codigo
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & VAR_COD2 & " AND bien_codigo = '" & CodBien & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    If rs_det1.RecordCount > 0 Then
        rs_det1.MoveFirst
        While Not rs_det1.EOF
        
            Set rs_compra_det = New Recordset
            If rs_compra_det.State = 1 Then rs_compra_det.Close
            rs_compra_det.Open "Select * from ao_compra_adjudica_bienes WHERE compra_codigo = " & VAR_COD2 & " AND bien_codigo = '" & rs_det1!bien_codigo & "'", db, adOpenKeyset, adLockOptimistic
        
            If rs_compra_det.RecordCount = 0 Then
                db.Execute "INSERT INTO ao_compra_adjudica_bienes (ges_gestion, compra_codigo, adjudica_codigo, bien_codigo, compra_codigo_det,                         grupo_codigo,                   subgrupo_codigo,                    par_codigo,                 adjudica_cantidad,              bien_cantidad_adjudica,           bien_precio_adjudica_bs,                  bien_total_adjudica_bs,             tipo_moneda, unimed_codigo,                 unimed_codigo_empaque, bien_cantidad_por_empaque, " & _
                    " marca_codigo, modelo_codigo, bien_nro_lote, bien_fecha_vencimiento, estado_codigo, usr_codigo, fecha_registro, hora_registro, compra_concepto, almacen_codigo, adjudica_monto_bs_87 ) " & _
                    " VALUES ('" & rs_det1!ges_gestion & "', " & VAR_COD2 & ", " & DETALLE2 & ", '" & rs_det1!bien_codigo & "', " & rs_det1!compra_codigo_det & ", '" & rs_det1!grupo_codigo & "', '" & rs_det1!subgrupo_codigo & "', '" & rs_det1!par_codigo & "', " & rs_det1!compra_cantidad & ", " & rs_det1!compra_cantidad & ", " & rs_det1!compra_precio_unitario_bs & ", " & rs_det1!compra_precio_total_bs & ", 'BOB', '" & rs_det1!unimed_codigo & "', '" & rs_det1!unimed_codigo & "', " & rs_det1!compra_cantidad & ", " & _
                    " '', '', '0', '" & Date & "', 'REG', '" & glusuario & "', '" & Date & "', '', '" & rs_det1!compra_concepto & "', " & rs_det1!almacen_codigo & ", " & rs_det1!adjudica_monto_bs_87 & "  )"
                'rs_compra_det.AddNew
            End If
            
            'rs_compra_det!compra_codigo_det = rs_det1!compra_codigo_det
            'rs_compra_det!ges_gestion = rs_det1!ges_gestion
            'rs_compra_det!compra_codigo = VAR_COD2                                      'rs_det1!compra_codigo
            'rs_compra_det!adjudica_codigo = DETALLE2        'Ado_detalle2.Recordset!adjudica_codigo
            'rs_compra_det!bien_codigo = rs_det1!bien_codigo
            'rs_compra_det!adjudica_cantidad = rs_det1!compra_cantidad
            'rs_compra_det!bien_precio_adjudica_bs = rs_det1!compra_precio_unitario_bs
            'rs_compra_det!compra_descuento_bs = "0"
            'rs_compra_det!compra_descuento_dol = "0"
            'rs_compra_det!bien_total_adjudica_bs = rs_det1!compra_precio_total_bs
            'rs_compra_det!compra_concepto = rs_det1!compra_concepto
            'rs_compra_det!grupo_codigo = rs_det1!grupo_codigo
            'rs_compra_det!subgrupo_codigo = rs_det1!subgrupo_codigo
            'rs_compra_det!par_codigo = rs_det1!par_codigo
            'rs_compra_det!almacen_codigo = "0"
            'rs_compra_det!unimed_codigo = rs_det1!unimed_codigo
            'bien_descripcion
            'rs_compra_det!almacen_codigo = rs_det1!almacen_codigo
            'rs_compra_det!estado_codigo = "APR"
            'rs_compra_det!usr_codigo = glusuario
            'rs_compra_det!fecha_registro = Date
            'rs_compra_det!adjudica_monto_bs_87 = IIf(IsNull(rs_det1!adjudica_monto_bs_87), 0, rs_det1!adjudica_monto_bs_87)
            'rs_compra_det.Update
            rs_det1.MoveNext
        Wend
'
'         Call OptFilGral1_Click
'
'     If (dg_datos.SelBookmarks.Count <> 0) Then
'        dg_datos.SelBookmarks.Remove 0
'     End If
'     If Ado_datos.Recordset.RecordCount > 0 Then
'        rs_datos.Find "compra_codigo = " & VAR_COD2 & "   ", , , 1
'        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
''         If rs_det1.RecordCount > 0 Then
''         rs_det1.MoveLast
''        End If
'     Else
'        rs_datos.MoveLast
'     End If
'
'     If (dg_det2.SelBookmarks.Count <> 0) Then
'        dg_det2.SelBookmarks.Remove 0
'     End If
'     If Ado_detalle2.Recordset.RecordCount > 0 Then
'        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
'        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
'     Else
'        rs_det2.MoveLast
'     End If
'
      End If
End Sub

Private Sub Ado_detalle1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If parametro <> "COMEX" Then
        BtnAprobar1.Visible = True
'        BtnAprobar3.Visible = True
    End If

    If Ado_detalle1.Recordset.RecordCount > 0 Then
        If Ado_detalle1.Recordset!estado_codigo = "REG" Then
'            BtnAprobar3.Visible = True
        Else
'            BtnAprobar3.Visible = False
        End If
    End If

End Sub

Private Sub Ado_detalle1A_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Ado_detalle1A.Recordset.RecordCount > 0 Then
        If Ado_detalle1A.Recordset!estado_codigo = "REG" Then
'            BtnDesAprobar3.Visible = True
        Else
'            BtnDesAprobar3.Visible = False
        End If
    End If
End Sub

Private Sub Ado_detalle2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If parametro <> "COMEX" Then
        BtnAprobar1.Visible = True
'        BtnAprobar3.Visible = True
    End If
    If Ado_detalle2.Recordset.RecordCount > 0 Then

        If Ado_detalle2.Recordset!adjudica_codigo <> "" Then
            Set rs_det1A = New ADODB.Recordset
            If rs_det1A.State = 1 Then rs_det1A.Close
            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1A.Sort = "compra_codigo_det"
            Set Ado_detalle1A.Recordset = rs_det1A
            If Ado_detalle1A.Recordset.RecordCount > 0 Then
                dg_det1A.Visible = True
                Set dg_det1A.DataSource = Ado_detalle1A.Recordset
                'Command1.Visible = True
            Else
                dg_det1A.Visible = False
                ' Set Ado_detalle1A.Recordset = rsNada
                Set dg_det1A.DataSource = rsNada
                 'Command1.Visible = False
            End If
        End If
    Else
        dg_det1A.Visible = False
    End If
End Sub

Private Sub BtnAddDetalle1_Click()
On Error GoTo UpdateErr

 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount > 0 Then
       If Ado_detalle2.Recordset!estado_codigo = "APR" Then
          sino = MsgBox("No Se Puede Agregar M?s Items Por Que La Factura Ya Fu? Aprobada (APR)", vbCritical, "SOFIA")
          Exit Sub
       End If
    End If
 
    marca1 = Ado_datos.Recordset.Bookmark
    If rs_datos!estado_codigo = "REG" Then
      If Ado_detalle1.Recordset.RecordCount = 0 Then
        var_cod = Ado_datos.Recordset!compra_codigo   'Codigo Llave de la Tabla
        VAR_COMPRA = Ado_datos.Recordset!solicitud_codigo  'Codigo Llave de la Tabla
            Select Case Aux
                Case "30"                           'Aux = "30" 'CCD
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"              'R-160   Formulario Cargo de Cuenta
                    CodBien = "FONDO03"
                Case "29"                           'Aux = "29" 'VIAJES
                    'rs_datos!trans_codigo_egr = "46"
                    VAR_DOCF = "R-162"              'R-162   Formularios para Viajes
                    CodBien = "FONDO02"
                Case "28"                           'Aux = "28" 'CAJA CHICA
                    'rs_datos!trans_codigo_egr = "45"
                    VAR_DOCF = "R-161"              'R-161   Formulario Caja Chica
                    CodBien = "FONDO01"
                Case Else
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"
                    CodBien = "FONDO03"
            End Select

'        db.Execute "insert into ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,             grupo_codigo, subgrupo_codigo, " & _
'             " par_codigo, tipo_descuento, almacen_codigo, usr_usuario,    fecha_registro, hora_registro, unimed_codigo, estado_codigo, adjudica_monto_bs_87, compra_tdc, tipo_moneda) " & _
'             " VALUES ('" & glGestion & "', " & var_cod & ", '" & CodBien & "', '1', " & CDbl(lbl_total_bs.Text) & ",      '0',            " & CDbl(lbl_total_bs.Text) & ", " & CDbl(lbl_total_dol.Text) & ",           '0',   " & CDbl(lbl_total_dol.Text) & ", '" & Txt_descripcion.Text & "', '20000', '26000',  " & _
'             " '26990',    '0',            '15',        '" & glusuario & "', '" & Date & "', '0',       'UNI',            'REG', " & CDbl(lbl_total_bs.Text) * 0.87 & ", '6.96', 'BOB'  )"
             
        db.Execute "insert into ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,             grupo_codigo, subgrupo_codigo, " & _
             " par_codigo, tipo_descuento, almacen_codigo, usr_usuario,    fecha_registro, hora_registro, unimed_codigo, estado_codigo, adjudica_monto_bs_87, compra_tdc, tipo_moneda) " & _
             " VALUES ('" & glGestion & "', " & var_cod & ", '" & CodBien & "', '1', " & CDbl(lbl_total_bs.Text) & ",      '0',            " & CDbl(lbl_total_bs.Text) & ", " & CDbl(lbl_total_dol.Text) & ",           '0',   " & CDbl(lbl_total_dol.Text) & ", '" & Txt_descripcion.Text & "', '20000', '26000',  " & _
             " '26990',    '0',            '15',        '" & glusuario & "', '" & Date & "', '0',       'UNI',            'REG', " & CDbl(lbl_total_bs.Text) * 0.87 & ", '6.96', 'BOB'  )"
     Else
     End If

'        swnuevo = 1
'        VAR_SW = "NEW"
'        fraOpciones.Visible = False
'        fraOpcionesDet.Visible = False
'        FraNavega.Enabled = False
'        FraDet2.Enabled = False
'        FrmABMDet2.Visible = False
'        FraDet1.Enabled = False
'            'Ado_detalle1.Recordset.AddNew
'            frm_solicitud_bienes_gral.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes_gral.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes_gral.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes_gral.lbl_edif.Caption = Label1.Caption
'            frm_solicitud_bienes_gral.lbl_det.Caption = Glaux
'            frm_solicitud_bienes_gral.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes_gral.MOD_NEW.Caption = "NEW"
'            frm_solicitud_bienes_gral.dtc_desc1.Text = ""
'            frm_solicitud_bienes_gral.txt_gestion.Caption = Year(DTPfecha1.Value)
'            frm_solicitud_bienes_gral.Show vbModal
''
''    swnuevo = 0
'    fraOpciones.Visible = True
'    fraOpcionesDet.Visible = True
'    FraNavega.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Visible = True
'    FraDet1.Enabled = True
'    BtnSalir.Visible = True
        
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya est? Aprobado!! ", vbExclamation
  End If

    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
         If rs_det1.RecordCount > 0 Then
         rs_det1.MoveLast
        End If
     Else
        rs_datos.MoveLast
     End If

End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAddDetalle2_Click()
On Error GoTo UpdateErr
  If parametro <> "COMEX" Then
    BtnAprobar1.Visible = True
'    BtnAprobar3.Visible = True
  End If
  VAR_SW = "NEW"
  'sw_nuevo = "NEW"
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo = "APR" Then       'estado_codigo_eqp
        MsgBox "No se puede Adicionar un nuevo registro, porque este ya est? Aprobado!! ", vbExclamation
        Exit Sub
    End If
    If Ado_detalle1.Recordset.RecordCount = 0 Then
        MsgBox "No puede Registrar la Factura, debe registrar previamente el DETALLE DE BIENES, Vuelva a Intentar !! ", vbExclamation
        Exit Sub
    End If
    swnuevo = 1
    fraOpciones.Visible = False
    fraOpcionesDet.Visible = False
    FraNavega.Enabled = False
    FraDet2.Visible = False
    FrmABMDet2.Visible = False
        'QUITAR LA GENERACION DE ao_compra_adjudica     'WWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'    Set rs_aux4 = New Recordset
'    If rs_aux4.State = 1 Then rs_aux4.Close
'    rs_aux4.Open "select max(adjudica_codigo) as correla from ao_compra_adjudica where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "", db, adOpenKeyset, adLockOptimistic
'                'Call ABRIR_TABLA_DET
        VAR_PAIS = Ado_datos.Recordset!beneficiario_codigo_resp
        
        Ado_detalle2.Recordset.AddNew
        fw_adjudica_fondos.txt_codigo.Caption = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
        fw_adjudica_fondos.Txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
        fw_adjudica_fondos.Txt_descripcion.Caption = Me.dtc_desc1.Text
        fw_adjudica_fondos.txtCodigo1.Caption = Me.Ado_datos.Recordset!compra_codigo
'                If rs_aux4!correla > 0 Then
'                    fw_adjudica_fondos.lbl_adjudica.Caption = rs_aux4!correla + 1
'                Else
'                    fw_adjudica_fondos.lbl_adjudica.Caption = "1"
'                End If
        fw_adjudica_fondos.txtSW.Text = swnuevo             '"C"
        fw_adjudica_fondos.txt_total_dol = VAR_FOBSEG
        fw_adjudica_fondos.txt_total_bs = VAR_FOBSEG2
        fw_adjudica_fondos.txt_pais.Text = VAR_PAIS
        fw_adjudica_fondos.txtEstado.Text = "REG"
        fw_adjudica_fondos.txtfecha_compra.Value = Date
        fw_adjudica_fondos.txt_total_bs = "0"
        fw_adjudica_fondos.cmd_unimed2 = "MES"
        fw_adjudica_fondos.txt_tipo_cambio = GlTipoCambioOficial
        fw_adjudica_fondos.opt_bs.Value = True
        fw_adjudica_fondos.cmb_mes_ini.Text = UCase(MonthName(Month(Date)))
'            fw_adjudica_fondos.txtFecha.Visible = False
'            fw_adjudica_fondos.txtFecha2.Visible = False
'            fw_adjudica_fondos.txtFecha3.Visible = False
'            fw_adjudica_fondos.txt_nro_dui.Enabled = False
'            fw_adjudica_fondos.lblbien(2).Visible = False
'            fw_adjudica_fondos.lblbien(3).Visible = False
'            fw_adjudica_fondos.lblbien(4).Visible = False
        fw_adjudica_fondos.txt_total_bs = SUMbs
        fw_adjudica_fondos.txt_total_dol = SUMdol
        'Set rs_aux8 = New Recordset
'       If rs_aux8.State = 1 Then rs_aux8.Close
'       rs_aux8.Open "select sum(bien_total_adjudica_bs) as total from ao_compra_adjudica_bienes where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic
'       fw_adjudica_fondos.txt_total_bs = rs_aux8!total
        fw_adjudica_fondos.Show vbModal
            
'    '        Case "V"    'FACTURACION LOCAL - COMEX
'    '    End Select
'        swnuevo = 0
        fraOpciones.Visible = True
        FraNavega.Enabled = True
        FraDet2.Visible = True
        FrmABMDet2.Visible = True
        fraOpcionesDet.Visible = True
        BtnSalir.Visible = True
'    End If
'  Else 'ESTADO
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya est? Aprobado!! ", vbExclamation
'    Exit Sub
'  End If 'ESTADO
'  DETALLE2 = Ado_detalle1.Recordset!adjudica_codigo
    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
'
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
'        If rs_det2.RecordCount > 0 Then
'         rs_det2.MoveLast
'        End If
     End If
     
 If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 And VAR_SW <> "NEW" Then
        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     det2.MoveLast
     End If
End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
  
  
  
  
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Est? Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atenci?n")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado ...", vbExclamation, "Validaci?n de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro no fue identificado correctamente ...", vbExclamation, "Validaci?n de Registro"
 End If
End Sub

Private Sub BtnAddDetalle3_Click()
On Error GoTo UpdateErr
If parametro <> "COMEX" Then
    BtnAprobar1.Visible = True
    BtnAprobar3.Visible = True
End If
If Ado_detalle2.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset.RecordCount > 0 Then

        If Ado_detalle2.Recordset!estado_codigo = "APR" Then
            sino = MsgBox("No Se Puede Agregar Este ITEM, La factura Ya Esta APROBADA(APR)", vbCritical, "ERROR")
            Exit Sub
        End If

        If Ado_detalle1.Recordset!almacen_codigo = "" Then
            sino = MsgBox("Debe registrar el almacen en el bien", vbCritical, "SOFIA")
            Exit Sub
        End If
     Set rs_compra_det = New Recordset
     If rs_compra_det.State = 1 Then rs_compra_det.Close
     rs_compra_det.Open "Select * from ao_compra_adjudica_bienes WHERE compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' And compra_codigo_det = " & Ado_detalle1.Recordset!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
     If rs_compra_det.RecordCount > 0 Then
        MsgBox "Este bien ya esta con un proveedor", vbExclamation, "Validaci?n de Registro"
        Exit Sub
     End If
     
     Set rs_compra_det = New Recordset
     If rs_compra_det.State = 1 Then rs_compra_det.Close
     rs_compra_det.Open "Select * from ao_compra_adjudica_bienes", db, adOpenKeyset, adLockOptimistic
'     If MOD_NEW.Caption = "NEW" Then
'
'     rs_compra_det.Open "Select * from ao_compra_adjudica_bienes", db, adOpenKeyset, adLockOptimistic
'     rs_compra_det.AddNew
'     rs_compra_det!compra_codigo_det = ao_compra_adjudica_bienes.Ado_detalle1.Recordset.RecordCount + 1
'     Else
'      rs_compra_det.Open "Select * from ao_compra_adjudica_bienes WHERE compra_codigo = " & fw_compras_fondos.Ado_datos.Recordset!compra_codigo & " AND compra_codigo_det = " & fw_compras_fondos.Ado_detalle1.Recordset!compra_codigo_det & "", db, adOpenKeyset, adLockOptimistic
'
'     End If
        rs_compra_det.AddNew
        rs_compra_det!compra_codigo_det = Ado_detalle1.Recordset!compra_codigo_det
        rs_compra_det!ges_gestion = Ado_detalle1.Recordset!ges_gestion
        rs_compra_det!compra_codigo = Ado_detalle1.Recordset!compra_codigo
        rs_compra_det!adjudica_codigo = Ado_detalle2.Recordset!adjudica_codigo
        rs_compra_det!bien_codigo = Ado_detalle1.Recordset!bien_codigo
        rs_compra_det!adjudica_cantidad = Ado_detalle1.Recordset!compra_cantidad
        rs_compra_det!bien_precio_adjudica_bs = Ado_detalle1.Recordset!compra_precio_unitario_bs
        'rs_compra_det!compra_descuento_bs = "0"
        'rs_compra_det!compra_descuento_dol = "0"
        rs_compra_det!bien_total_adjudica_bs = Ado_detalle1.Recordset!compra_precio_total_bs
        rs_compra_det!compra_concepto = Ado_detalle1.Recordset!compra_concepto
        rs_compra_det!grupo_codigo = Ado_detalle1.Recordset!grupo_codigo
        rs_compra_det!subgrupo_codigo = Ado_detalle1.Recordset!subgrupo_codigo
        rs_compra_det!par_codigo = Ado_detalle1.Recordset!par_codigo
        'rs_compra_det!almacen_codigo = "0"
        rs_compra_det!unimed_codigo = Ado_detalle1.Recordset!unimed_codigo
        'bien_descripcion
        rs_compra_det!almacen_codigo = Ado_detalle1.Recordset!almacen_codigo
         rs_compra_det!estado_codigo = "APR"
        rs_compra_det!usr_codigo = glusuario
        rs_compra_det!fecha_registro = Date
        rs_compra_det.Update
        DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
        DETALLE1 = Ado_detalle1.Recordset!compra_codigo_det
        Call ABRIR_TABLA_DET

   If (dg_det1.SelBookmarks.Count <> 0) Then
        dg_det1.SelBookmarks.Remove 0
     End If
     If Ado_detalle1.Recordset.RecordCount > 0 Then
        rs_det1.Find "compra_codigo_det = " & DETALLE1 & "   ", , , 1
        dg_det1.SelBookmarks.Add (rs_det1.Bookmark)
     Else
        rs_det1.MoveLast
     End If
     
     
     If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     Else
        rs_det2.MoveLast
     End If
     
Else
sino = MsgBox("Primero llene La ADJUDICACION (Proveedores)", vbExclamation, "Atenci?n")
End If
End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description


End Sub

Private Sub BtnAddDetalle4_Click()
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount > 0 Then
       If Ado_detalle2.Recordset!estado_codigo = "APR" Then
          sino = MsgBox("No se puede Agregar m?s Items porque la(s) Factura(s) ya fueron Aprobada(s) (APR)", vbCritical, "SOFIA")
          Exit Sub
       End If
    End If
 
'    marca1 = Ado_datos.Recordset.Bookmark
    If Ado_detalle2.Recordset!estado_codigo = "REG" Then
      'If Ado_detalle1.Recordset.RecordCount = 0 Then
        var_cod = Ado_detalle2.Recordset!compra_codigo          'Codigo Llave de la Tabla
        VAR_COMPRA = Ado_detalle2.Recordset!adjudica_codigo     'Codigo Llave de la Tabla
            Select Case Aux
                Case "30"                           'Aux = "30" 'CCD
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"              'R-160   Formulario Cargo de Cuenta
                    CodBien = "FONDO03"
                Case "29"                           'Aux = "29" 'VIAJES
                    'rs_datos!trans_codigo_egr = "46"
                    VAR_DOCF = "R-162"              'R-162   Formularios para Viajes
                    CodBien = "FONDO02"
                Case "28"                           'Aux = "28" 'CAJA CHICA
                    'rs_datos!trans_codigo_egr = "45"
                    VAR_DOCF = "R-161"              'R-161   Formulario Caja Chica
                    CodBien = "FONDO01"
                Case Else
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"
                    CodBien = "FONDO03"
            End Select


'ges_gestion, compra_codigo, adjudica_codigo, bien_codigo, compra_codigo_det, adjudica_old, grupo_codigo, subgrupo_codigo, par_codigo, adjudica_cantidad, bien_cantidad_adjudica, bien_precio_adjudica_bs, bien_total_adjudica_bs, tipo_moneda, unimed_codigo,
'                  unimed_codigo_empaque , bien_cantidad_por_empaque, marca_codigo, modelo_codigo, bien_nro_lote, bien_fecha_vencimiento, compra_concepto, almacen_codigo, adjudica_monto_bs_87, solicitud_tipo, estado_codigo, usr_codigo, fecha_registro, hora_registro


        db.Execute "insert into ao_compra_adjudica_bienes (ges_gestion, compra_codigo, adjudica_codigo, bien_codigo, compra_codigo_det, adjudica_old, grupo_codigo, subgrupo_codigo, par_codigo, adjudica_cantidad, bien_cantidad_adjudica, bien_precio_adjudica_bs, bien_total_adjudica_bs, tipo_moneda, unimed_codigo, " & _
             " unimed_codigo_empaque , bien_cantidad_por_empaque, marca_codigo, modelo_codigo, bien_nro_lote, bien_fecha_vencimiento, compra_concepto, almacen_codigo, adjudica_monto_bs_87, solicitud_tipo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
             " VALUES ('" & glGestion & "', " & var_cod & ", " & VAR_COMPRA & ", '1', " & CDbl(lbl_total_bs.Text) & ",      '0',            " & CDbl(lbl_total_bs.Text) & ", " & CDbl(lbl_total_dol.Text) & ",           '0',   " & CDbl(lbl_total_dol.Text) & ", '" & Txt_descripcion.Text & "', '20000', '26000',  " & _
             " '26990',    '0',            '15',        '" & glusuario & "', '" & Date & "', '0',       'UNI',            'REG', " & CDbl(lbl_total_bs.Text) * 0.87 & ", '6.96', 'BOB'  )"
             
'        db.Execute "insert into ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,             grupo_codigo, subgrupo_codigo, " & _
'             " par_codigo, tipo_descuento, almacen_codigo, usr_usuario,    fecha_registro, hora_registro, unimed_codigo, estado_codigo, adjudica_monto_bs_87, compra_tdc, tipo_moneda) " & _
'             " VALUES ('" & glGestion & "', " & var_cod & ", '" & CodBien & "', '1', " & CDbl(lbl_total_bs.Text) & ",      '0',            " & CDbl(lbl_total_bs.Text) & ", " & CDbl(lbl_total_dol.Text) & ",           '0',   " & CDbl(lbl_total_dol.Text) & ", '" & Txt_descripcion.Text & "', '20000', '26000',  " & _
'             " '26990',    '0',            '15',        '" & glusuario & "', '" & Date & "', '0',       'UNI',            'REG', " & CDbl(lbl_total_bs.Text) * 0.87 & ", '6.96', 'BOB'  )"
     'Else
     'End If

'        swnuevo = 1
'        VAR_SW = "NEW"
'        fraOpciones.Visible = False
'        fraOpcionesDet.Visible = False
'        FraNavega.Enabled = False
'        FraDet2.Enabled = False
'        FrmABMDet2.Visible = False
'        FraDet1.Enabled = False
'            'Ado_detalle1.Recordset.AddNew
'            frm_solicitud_bienes_gral.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes_gral.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes_gral.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes_gral.lbl_edif.Caption = Label1.Caption
'            frm_solicitud_bienes_gral.lbl_det.Caption = Glaux
'            frm_solicitud_bienes_gral.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes_gral.MOD_NEW.Caption = "NEW"
'            frm_solicitud_bienes_gral.dtc_desc1.Text = ""
'            frm_solicitud_bienes_gral.txt_gestion.Caption = Year(DTPfecha1.Value)
'            frm_solicitud_bienes_gral.Show vbModal
''
''    swnuevo = 0
'    fraOpciones.Visible = True
'    fraOpcionesDet.Visible = True
'    FraNavega.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Visible = True
'    FraDet1.Enabled = True
'    BtnSalir.Visible = True
        
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya est? Aprobado!! ", vbExclamation
  End If

    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
         If rs_det1.RecordCount > 0 Then
         rs_det1.MoveLast
        End If
     Else
        rs_datos.MoveLast
     End If

End If

End Sub

Private Sub BtnAnlDetalle1_Click()

'If rs_datos!estado_codigo <> "REG" Or Ado_detalle2.Recordset!estado_codigo = "REG" Then
'sino = MsgBox("NO se puede eliminar este registro si ya esta aprobado o anulado")
'Exit Sub
'End If
 Select Case Glaux
             Case "PROVI"
             If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "TRANS"
             If Ado_datos.Recordset!estado_codigo_tra <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "ADUAN"
             If Ado_datos.Recordset!estado_codigo_nac <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "DESCA"
             If Ado_datos.Recordset!estado_codigo_des <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "CONTR"
             If Ado_datos.Recordset!estado_codigo <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             
             Case Else
              If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
    End Select
    
Set rs_aux8 = New ADODB.Recordset
'            If rs_aux8.State = 1 Then rs_aux8.Close
'            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_aux8.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
'            'rs_det1A.Sort = "compra_codigo_det"
      If Ado_detalle2.Recordset.RecordCount > 0 Then
      If Ado_detalle2.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("La factura ya fue aprobada, no se puede eliminar este item ", vbInformation, "AVISO")
      Exit Sub
      End If
      Else
      sino = MsgBox("Est? Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbExclamation, "Atenci?n")
      If sino = vbYes Then
      'Ado_detalle1.Recordset.Delete adAffectCurrent
      db.Execute "delete ao_compra_detalle where compra_codigo = " & Ado_detalle1.Recordset!compra_codigo & " AND bien_codigo = " & Ado_detalle1.Recordset!bien_codigo & "AND compra_codigo_det = " & Ado_detalle1.Recordset!compra_codigo_det
      Call ABRIR_TABLA_DET
      End If
      End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   sino = MsgBox("Est? Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atenci?n")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validaci?n de Registro"
   End If
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  
     sino = MsgBox("Desea APROBAR el Registro ? (Ya no podr? modificarlo)", vbYesNo + vbInformation, "Atenci?n")
     If sino = vbYes Then
        If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
            MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
            Exit Sub
        End If
        rs_datos!estado_codigo_eqp = "APR"
        'rs_datos!estado_codigo = "APR"
        rs_datos!fecha_registro = Date
        rs_datos!usr_codigo = glusuario
        rs_datos.UpdateBatch adAffectAll
        
        var_cod = Ado_datos.Recordset!compra_codigo   'Codigo Llave de la Tabla
        VAR_COMPRA = Ado_datos.Recordset!solicitud_codigo  'Codigo Llave de la Tabla
        VAR_UNI = Ado_datos.Recordset!unidad_codigo
            Select Case Aux
                Case "30"                           'Aux = "30" 'CCD
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"              'R-160   Formulario Cargo de Cuenta
                    CodBien = "FONDO03"
                Case "29"                           'Aux = "29" 'VIAJES
                    'rs_datos!trans_codigo_egr = "46"
                    VAR_DOCF = "R-162"              'R-162   Formularios para Viajes
                    CodBien = "FONDO02"
                Case "28"                           'Aux = "28" 'CAJA CHICA
                    'rs_datos!trans_codigo_egr = "45"
                    VAR_DOCF = "R-161"              'R-161   Formulario Caja Chica
                    CodBien = "FONDO01"
                Case Else
                    'rs_datos!trans_codigo_egr = "47"
                    VAR_DOCF = "R-160"
                    CodBien = "FONDO03"
            End Select
        
        'graba ADJUDICA PARA PAGAR EL FONDO EN AVANCE
        db.Execute "Insert INTO ao_compra_adjudica (ges_gestion, compra_codigo, unidad_codigo, solicitud_codigo, fecha_compra, adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo,  doc_codigo,                         doc_numero, " & _
                   " nro_nota_remision, beneficiario_codigo, adjudica_descripcion, adjudica_cantidad_total, adjudica_monto_bs, tipo_moneda,     adjudica_monto_dol,                                 fecha_inicio_contrato,                  fecha_fin_contrato,                             fecha_envio_proveedor, " & _
                   " fecha_recibe_almacen, almacen_codigo, poa_codigo, mes_inicio_crono, cantidad_cuotas_pag, unimed_codigo_pag, correl_pagos_prog, compra_codigo_det, observaciones,                           nro_autorizacion, codigo_control, nro_dui, " & _
                   " tasas_ice_iehd, grabado_tasa_cero, importe_no_credito_fisc, sub_total, descuento, importe_cred_fisc,       credito_fiscal_13,                                  adjudica_monto_bs_87,                                   adjudica_monto_dol_87,                                  tipo_compra, tipo_cambio,           Literal, literal_neto, factura, " & _
                   " doc_codigo_alm, doc_numero_alm, estado_almacen, estado_codigo, usr_codigo, fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba, nit_empresa, nit_beneficiario, trans_codigo, trans_codigo_fac, beneficiario_codigo_resp, solicitud_tipo )  " & _
        " VALUES ('" & glGestion & "', " & var_cod & ",  '" & VAR_UNI & "', " & VAR_COMPRA & ", '" & Ado_datos.Recordset!compra_fecha & "', '" & Date & "', 'FIN',      'FIN-03',           'FIN-03-02', 'ADM', '" & Ado_datos.Recordset!doc_codigo & "', " & Ado_datos.Recordset!doc_numero & ", " & _
              " '0', '0', '" & Ado_datos.Recordset!compra_DESCRIPCION & "', '1', " & CDbl(Ado_datos.Recordset!compra_monto_bs) & ", 'BOB', " & CDbl(Ado_datos.Recordset!compra_monto_DOL) & ", '" & Ado_datos.Recordset!compra_fecha & "', '" & Ado_datos.Recordset!compra_fecha & "', '" & Ado_datos.Recordset!compra_fecha & "', " & _
              " '" & Date & "',             '15',           '4.2.3',    'ENERO',        '0',                    'MES',              '1',            '1',            '" & Ado_datos.Recordset!compra_DESCRIPCION & "', '0',          '0',            '0', " & _
              " '0',                    '0',            '0', " & Ado_datos.Recordset!compra_monto_bs & ", '0', '0', " & Round(Ado_datos.Recordset!compra_monto_bs * 0.13, 2) & ", " & Round(Ado_datos.Recordset!compra_monto_bs * 0.87, 2) & ", " & Round(Ado_datos.Recordset!compra_monto_DOL * 0.87, 2) & ", '1', " & GlTipoCambioOficial & ", 'Bolivianos', 'Bolivianos', 'NO', " & _
              " '0', '0', 'REG', 'REG', '" & glusuario & "', '" & Date & "', '', '" & glusuario & "', '" & Date & "', '0', '0', '" & Ado_datos.Recordset!trans_codigo_EGR & "', '38', '" & Ado_datos.Recordset!beneficiario_codigo_resp & "', '1' ) "

     End If

'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validaci?n de Registro"
'   End If
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atenci?n!"
'  End If
Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar1_Click()
On Error GoTo AddErr
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount = 0 Then
        MsgBox "No se puede APROBAR, debe registrar al menos una Factura del Proveedor ", vbExclamation
        Exit Sub
    End If
    VAR_COD2 = Ado_datos.Recordset!compra_codigo
    If parametro = "COMEX" Then
     sino = MsgBox("Desea APROBAR el Registro ? (Ya no podr? modificarlo)", vbYesNo + vbInformation, "Atenci?n")
     If sino = vbYes Then
        Select Case Glaux
             Case "PROVI"
                If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
                    MsgBox "No se puede modificar este registro, porque este ya est? Aprobado o Anulado (ANL)!! ", vbExclamation
                    Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo_eqp = "APR"
                    Ado_datos.Recordset!estado_codigo_tra = "REG"
                    Ado_datos.Recordset.Update

                End If
             Case "TRANS"
                If Ado_datos.Recordset!estado_codigo_tra <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                Ado_datos.Recordset!estado_codigo_tra = "APR"
                 Ado_datos.Recordset!estado_codigo_nac = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "ADUAN"
                If Ado_datos.Recordset!estado_codigo_nac <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                Ado_datos.Recordset!estado_codigo_nac = "APR"
                Ado_datos.Recordset!estado_codigo_des = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "DESCA"
                If Ado_datos.Recordset!estado_codigo_des <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo_des = "APR"
                    Ado_datos.Recordset!estado_codigo = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "CONTR"
                If Ado_datos.Recordset!estado_codigo <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo = "APR"
                    Ado_datos.Recordset.Update
                End If
             Case "GADM"
                If Ado_datos.Recordset!estado_codigo <> "REG" Then
                    MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                    Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo = "APR"
                    Ado_datos.Recordset.Update
                End If

             Case Else
                If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub

                End If
        End Select
        Ado_detalle2.Recordset("estado_codigo") = "APR"
        Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
        Ado_detalle2.Recordset("fecha_aprueba") = Date
        Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
        Ado_detalle2.Recordset.Update
     End If
    Else
        If Format(Ado_detalle2.Recordset!adjudica_monto_bs, "###,###,##0.00") > CDbl(lbl_total_bs.Text) Then
            sino = MsgBox("El monto introducido en la factura es mayor al monto de la Suma del precio de cada Item, Revise Por Favor", vbCritical, "SOFIA")
            Exit Sub
        End If
       Ado_detalle2.Recordset("estado_codigo") = "APR"
       Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
       Ado_detalle2.Recordset("fecha_aprueba") = Date
       Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
       Ado_detalle2.Recordset.Update
    
    End If
End If

'If parametro <> "COMEX" Then
'    If Format(Ado_detalle2.Recordset!adjudica_monto_bs, "###,###,##0.00") > CDbl(lbl_total_bs.text) Then
'        sino = MsgBox("El monto introducido en la factura es mayor al monto de la Suma del precio de cada Item, Revise Por Favor", vbCritical, "SOFIA")
'        Exit Sub
'    End If
'
'    If parametro <> "COMEX" Then
'        BtnAprobar1.Visible = True
'    End If
'    If Ado_detalle2.Recordset.RecordCount = 0 Then
'        MsgBox "No se puede APROBAR, debe registrar al menos una Factura del Proveedro ...", vbCritical, "SOFIA"
'        Exit Sub
'    End If
'    If Ado_detalle2.Recordset("estado_codigo") <> "REG" Then
'        sino = MsgBox("No se puede APROBAR un registro ANULADO o Aprobado", vbCritical, "SOFIA")
'        Exit Sub
'    End If
'    If Ado_detalle1.Recordset.RecordCount > 0 Then
'        sino = MsgBox("Desea APROBAR el Registro ? (Ya no podr? modificarlo)", vbYesNo + vbInformation, "Atenci?n")
'        If sino = vbYes Then
'            DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
'            VAR_COD2 = Ado_datos.Recordset!compra_codigo
'            'If VAR_TIPO_ALM = "" Then
'            '    VAR_TIPO_ALM = "R"
'            'End If
'
'            Select Case Glaux
'                Case "UALMI"
'                    VAR_TIPO_ALM = "I"
'                Case "UALMR"
'                    VAR_TIPO_ALM = "R"
'                Case "UALMH"
'                    VAR_TIPO_ALM = "H"
'                Case "GADM"
'                    VAR_TIPO_ALM = "M"
'                Case Else
'                    VAR_TIPO_ALM = "M"
'            End Select
'            'correlativo ALMACEN
'          Set rs_aux7 = New ADODB.Recordset
'          If rs_aux7.State = 1 Then rs_aux7.Close 'VAR_TIPO_ALM
'          rs_aux7.Open "Select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & VAR_DPTO_AUX & "' and cta_codigo2 = '" & VAR_TIPO_ALM & "' ) ", db, adOpenKeyset, adLockOptimistic
'          If rs_aux7.RecordCount > 0 Then
'             CORRELARIVO1 = IIf(IsNull(rs_aux7!numero_correlativo), 1, rs_aux7!numero_correlativo + 1)
'          Else
'             MsgBox "No se puede generar el Correlativo por los privilegios del USUARIO", vbCritical, "sofia"
'             Exit Sub
'          End If
'          If parametro <> "COMEX" Then
'             rs_datos!doc_numero_alm = CORRELARIVO1
'          End If
'
'         Call BIENES
'
'       Set rs_det1A = New ADODB.Recordset
'       If rs_det1A.State = 1 Then rs_det1A.Close
'       'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'       rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
'       rs_det1A.MoveFirst
'       'sino = rs_det1A.RecordCount
'        While Not rs_det1A.EOF
'
'            'ac_tipo_compra_venta
'            Set rs_aux12 = New ADODB.Recordset
'            If rs_aux12.State = 1 Then rs_aux12.Close
'            rs_aux12.Open "select * from ao_almacen_ingresos where ges_gestion='" & rs_det1A!ges_gestion & "' AND almacen_codigo=" & rs_det1A!almacen_codigo & " AND doc_codigo='" & Ado_datos.Recordset!doc_codigo_alm & "' AND doc_numero=" & CORRELARIVO1 & " AND bien_codigo='" & rs_det1A!bien_codigo & "' ", db, adOpenStatic
'            If rs_aux12.RecordCount > 0 Then
'            Else
'                db.Execute "ap_compras_grla 2,'" & rs_det1A!ges_gestion & "'," & rs_det1A!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "' ," & CORRELARIVO1 & ",'" & rs_det1A!bien_codigo & "','" & Ado_datos.Recordset!edif_codigo & "'," & VAR_COD2 & ",'" & Ado_detalle2.Recordset!beneficiario_codigo & "','" & Ado_detalle2.Recordset!fecha_compra & "'," & rs_det1A!adjudica_cantidad & "," & rs_det1A!bien_total_adjudica_bs & "," & CDbl(rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!compra_DESCRIPCION & "'," & rs_det1A!bien_precio_adjudica_bs & ""
'            End If
'
'            Set rs_aux6 = New ADODB.Recordset
'            If rs_aux6.State = 1 Then rs_aux6.Close
'            rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & rs_det1A!almacen_codigo & " AND bien_codigo = '" & rs_det1A!bien_codigo & "'", db, adOpenStatic
'            If rs_aux6.RecordCount > 0 Then
'                db.Execute "ap_almacen_totales 2," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
'            Else
'                db.Execute "ap_almacen_totales 1," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG', '" & glusuario & "' "
'                '15, '1.6.1.11', 1, 0, 1, 100, 0,0, 100/6.96, 0, 0, 'REG', 'ADMIN'
'            End If
'            rs_det1A.MoveNext
'        Wend
'
'        If Ado_detalle2.Recordset!adjudica_codigo <> "" Then
'            Set rs_det1A = New ADODB.Recordset
'            If rs_det1A.State = 1 Then rs_det1A.Close
'            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_det1A.Sort = "compra_codigo_det"
'            Set Ado_detalle1A.Recordset = rs_det1A
'                If Ado_detalle1A.Recordset.RecordCount > 0 Then
'                'dg_det1A.Visible = True
'
'                Set dg_det1A.DataSource = Ado_detalle1A.Recordset
'                'Command1.Visible = True
'                    Else
'                 'dg_det1A.Visible = False
'                ' Set Ado_detalle1A.Recordset = rsNada
'                Set dg_det1A.DataSource = rsNada
'
'                 'Command1.Visible = False
'                End If
'        End If
'
''        rs_datos!correl_bitacora = 0
'       'Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
''       Ado_detalle2.Recordset("estado_codigo") = "APR"
''       Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
''       Ado_detalle2.Recordset("fecha_aprueba") = Date
''       Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
''       Ado_detalle2.Recordset.Update
'
'    If IsNull(rs_datos!doc_numero) Then
'        rs_datos!doc_numero = CORRELARTIVO1
'    End If
'
'     rs_aux7!numero_correlativo = CORRELARIVO1
'
'     rs_aux7.Update
'     rs_datos.Update
     Call OptFilGral1_Click
    
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "compra_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
'         If rs_det1.RecordCount > 0 Then
'         rs_det1.MoveLast
'        End If
     Else
        rs_datos.MoveLast
     End If
     
     If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     Else
        rs_det2.MoveLast
     End If
        
'End If
'  Else
'    sino = MsgBox("El proveedor no tiene ningun bien", vbCritical, "SOFIA")
'  End If
'End If
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar3_Click()
On Error GoTo AddErr
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount = 0 Then
        MsgBox "No se puede APROBAR, debe registrar al menos una Factura del Proveedor ", vbExclamation
        Exit Sub
    End If
   If parametro = "COMEX" Then
     sino = MsgBox("Desea APROBAR el Registro ? (Ya no podr? modificarlo)", vbYesNo + vbInformation, "Atenci?n")
     If sino = vbYes Then
        Select Case Glaux
             Case "PROVI"
                If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
                    MsgBox "No se puede modificar este registro, porque este ya est? Aprobado o Anulado (ANL)!! ", vbExclamation
                    Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo_eqp = "APR"
                    Ado_datos.Recordset!estado_codigo_tra = "REG"
                    Ado_datos.Recordset.Update

                End If
             Case "TRANS"
                If Ado_datos.Recordset!estado_codigo_tra <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                Ado_datos.Recordset!estado_codigo_tra = "APR"
                 Ado_datos.Recordset!estado_codigo_nac = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "ADUAN"
                If Ado_datos.Recordset!estado_codigo_nac <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                Ado_datos.Recordset!estado_codigo_nac = "APR"
                Ado_datos.Recordset!estado_codigo_des = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "DESCA"
                If Ado_datos.Recordset!estado_codigo_des <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo_des = "APR"
                    Ado_datos.Recordset!estado_codigo = "REG"
                Ado_datos.Recordset.Update
                End If
             Case "CONTR"
                If Ado_datos.Recordset!estado_codigo <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo = "APR"
                    Ado_datos.Recordset.Update
                End If
             Case "GADM"
                If Ado_datos.Recordset!estado_codigo <> "REG" Then
                    MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                    Exit Sub
                Else
                    Ado_datos.Recordset!estado_codigo = "APR"
                    Ado_datos.Recordset.Update
                End If

             Case Else
                If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub

                End If
        End Select
        'Ado_detalle2.Recordset("estado_codigo") = "APR"
        'Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
        'Ado_detalle2.Recordset("fecha_aprueba") = Date
        'Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
        'Ado_detalle2.Recordset.Update
    End If
  Else
  
       'Ado_detalle2.Recordset("estado_codigo") = "APR"
       'Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
       'Ado_detalle2.Recordset("fecha_aprueba") = Date
       'Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
       'Ado_detalle2.Recordset.Update
    
  End If
End If

If parametro <> "COMEX" Then
    If Ado_detalle2.Recordset.RecordCount = 0 Then
        MsgBox "No se puede ENVIAR, debe registrar al menos una Factura del Proveedor ...", vbCritical, "SOFIA"
        Exit Sub
    End If
    'If Format(Ado_detalle2.Recordset!adjudica_monto_bs, "###,###,##0.00") > CDbl(lbl_total_bs.text) Then
    If (Ado_detalle1.Recordset!compra_precio_total_bs) > CDbl(Ado_detalle2.Recordset!adjudica_monto_bs) Then
        sino = MsgBox("El monto del Item elegido es mayor al importe de la factura, Revise Por Favor", vbCritical, "SOFIA")
        Exit Sub
    End If
     
    If parametro <> "COMEX" Then
        BtnAprobar1.Visible = True
    End If
    'If Ado_detalle2.Recordset.RecordCount = 0 Then
    '    MsgBox "No se puede APROBAR, debe registrar al menos una Factura del Proveedro ...", vbCritical, "SOFIA"
    '    Exit Sub
    'End If
    If Ado_detalle2.Recordset!estado_codigo <> "REG" Then
        sino = MsgBox("No se puede APROBAR un registro ANULADO o Aprobado", vbCritical, "SOFIA")
        Exit Sub
    End If
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        If Ado_detalle1.Recordset!estado_codigo <> "REG" Then
            sino = MsgBox("El registro ya fue ENVIADO o fue ANULADO ... Elija otro registro y vuelva a intentar....", vbCritical, "SOFIA")
            Exit Sub
        End If
        sino = MsgBox("Desea Envia como Ingreso a Almacen y adem?s, ser? parte de la Factura Elegida... ? (Ya no podr? modificarlo)", vbYesNo + vbInformation, "Atenci?n")
        If sino = vbYes Then
            DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
            VAR_COD2 = Ado_datos.Recordset!compra_codigo
            'If VAR_TIPO_ALM = "" Then
            '    VAR_TIPO_ALM = "R"
            'End If

            Select Case Glaux
                Case "UALMI"
                    VAR_TIPO_ALM = "I"
                Case "UALMR"
                    VAR_TIPO_ALM = "R"
                Case "UALMH"
                    VAR_TIPO_ALM = "H"
                Case "GADM"
                    VAR_TIPO_ALM = "M"
                Case Else
                    VAR_TIPO_ALM = "M"
            End Select
'            'INI correlativo ALMACEN
'          Set rs_aux7 = New ADODB.Recordset
'          If rs_aux7.State = 1 Then rs_aux7.Close 'VAR_TIPO_ALM
'          rs_aux7.Open "Select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & VAR_DPTO_AUX & "' and cta_codigo2 = '" & VAR_TIPO_ALM & "' ) ", db, adOpenKeyset, adLockOptimistic
'          If rs_aux7.RecordCount > 0 Then
'             CORRELARIVO1 = IIf(IsNull(rs_aux7!numero_correlativo), 1, rs_aux7!numero_correlativo + 1)
'          Else
'             MsgBox "No se puede generar el Correlativo por los privilegios del USUARIO", vbCritical, "sofia"
'             Exit Sub
'          End If
'          rs_aux7!numero_correlativo = CORRELARIVO1
'          rs_aux7.Update
'
'          If parametro <> "COMEX" Then
'            If IsNull(rs_datos!doc_numero) Or (rs_datos!doc_numero = 0) Then
'                rs_datos!doc_numero = CORRELARTIVO1
'                rs_datos.Update
'            End If
'             'rs_datos!doc_numero_alm = CORRELARIVO1
'          End If
'
'          If IsNull(Ado_detalle2.Recordset!doc_numero_alm) Or (Ado_detalle2.Recordset!doc_numero_alm = 0) Then
'                Ado_detalle2.Recordset!doc_numero_alm = CORRELARIVO1
'                Ado_detalle2.Recordset.Update
'          End If
         Call BIENES
        If (Ado_datos.Recordset!edif_codigo = "20101-3") Or (Ado_datos.Recordset!edif_codigo = "30101-3") Or (Ado_datos.Recordset!edif_codigo = "70101-3") Or (Ado_datos.Recordset!edif_codigo = "10101-3") Then
           CORRELARIVO1 = "0"  'CORRELATIVO PARA SALDOS INICIALES
           Set rs_det1A = New ADODB.Recordset
           If rs_det1A.State = 1 Then rs_det1A.Close
           'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
           rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
           rs_det1A.MoveFirst
           sino = rs_det1A.RecordCount
            While Not rs_det1A.EOF
                'ac_tipo_compra_venta
                Set rs_aux12 = New ADODB.Recordset
                If rs_aux12.State = 1 Then rs_aux12.Close
                rs_aux12.Open "select * from ao_almacen_ingresos where ges_gestion='" & rs_det1A!ges_gestion & "' AND almacen_codigo=" & rs_det1A!almacen_codigo & " AND doc_codigo='" & Ado_datos.Recordset!doc_codigo_alm & "' AND doc_numero=" & CORRELARIVO1 & " AND bien_codigo='" & rs_det1A!bien_codigo & "' ", db, adOpenStatic
                If rs_aux12.RecordCount > 0 Then
                Else
                    db.Execute "ap_compras_grla 2,'" & rs_det1A!ges_gestion & "'," & rs_det1A!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "' ," & CORRELARIVO1 & ",'" & rs_det1A!bien_codigo & "','" & Ado_datos.Recordset!edif_codigo & "'," & VAR_COD2 & ",'" & Ado_detalle2.Recordset!beneficiario_codigo & "','" & Ado_detalle2.Recordset!fecha_compra & "'," & rs_det1A!adjudica_cantidad & "," & rs_det1A!bien_total_adjudica_bs & "," & CDbl(rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!compra_DESCRIPCION & "'," & rs_det1A!bien_precio_adjudica_bs & ""
                End If
    
                Set rs_aux6 = New ADODB.Recordset
                If rs_aux6.State = 1 Then rs_aux6.Close
                rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & rs_det1A!almacen_codigo & " AND bien_codigo = '" & rs_det1A!bien_codigo & "'", db, adOpenStatic
                If rs_aux6.RecordCount > 0 Then
                    db.Execute "ap_almacen_totales 2," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
                Else
                    db.Execute "ap_almacen_totales 1," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG', '" & glusuario & "' "
                    '15, '1.6.1.11', 1, 0, 1, 100, 0,0, 100/6.96, 0, 0, 'REG', 'ADMIN'
                End If
                rs_det1A.MoveNext
            Wend
    
            If Ado_detalle2.Recordset!adjudica_codigo <> "" Then
                Set rs_det1A = New ADODB.Recordset
                If rs_det1A.State = 1 Then rs_det1A.Close
                'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
                rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
                rs_det1A.Sort = "compra_codigo_det"
                Set Ado_detalle1A.Recordset = rs_det1A
                    If Ado_detalle1A.Recordset.RecordCount > 0 Then
                    'dg_det1A.Visible = True
    
                    Set dg_det1A.DataSource = Ado_detalle1A.Recordset
                    'Command1.Visible = True
                        Else
                     'dg_det1A.Visible = False
                    ' Set Ado_detalle1A.Recordset = rsNada
                    Set dg_det1A.DataSource = rsNada
    
                     'Command1.Visible = False
                    End If
            End If
        End If
        Ado_detalle1.Recordset!estado_codigo = "APR"
        Ado_detalle1.Recordset.Update
        
'        rs_datos!correl_bitacora = 0
       'Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
'       Ado_detalle2.Recordset("estado_codigo") = "APR"
'       Ado_detalle2.Recordset("usr_codigo_aprueba") = glusuario
'       Ado_detalle2.Recordset("fecha_aprueba") = Date
'       Ado_detalle2.Recordset("fecha_recibe_almacen") = Date
'       Ado_detalle2.Recordset.Update
       
'    If IsNull(rs_datos!doc_numero) Then
'        rs_datos!doc_numero = CORRELARTIVO1
'    End If
'
'     rs_aux7!numero_correlativo = CORRELARIVO1
'
'     rs_aux7.Update
'     rs_datos.Update
     Call OptFilGral1_Click
    
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "compra_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
'         If rs_det1.RecordCount > 0 Then
'         rs_det1.MoveLast
'        End If
     Else
        rs_datos.MoveLast
     End If
     
     If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     Else
        rs_det2.MoveLast
     End If
        
End If
Else
    sino = MsgBox("El proveedor no tiene ningun bien", vbCritical, "SOFIA")
End If
End If
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar4_Click()
 If (Ado_datos.Recordset.RecordCount > 0) And (Ado_datos.Recordset!edif_codigo <> "20101-3") And (Ado_datos.Recordset!edif_codigo <> "70101-3") And (Ado_datos.Recordset!edif_codigo <> "30101-3") And (Ado_datos.Recordset!edif_codigo <> "10101-3") Then
    If (Ado_detalle2.Recordset.RecordCount > 0) And (IsNull(Ado_detalle2.Recordset!doc_numero_alm) Or (Ado_detalle2.Recordset!doc_numero_alm = 0)) Then
        VAR_COD2 = Ado_datos.Recordset!compra_codigo
        'INI correlativo ALMACEN
          Set rs_aux7 = New ADODB.Recordset
          If rs_aux7.State = 1 Then rs_aux7.Close 'VAR_TIPO_ALM
          rs_aux7.Open "Select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & VAR_DPTO_AUX & "' and cta_codigo2 = '" & VAR_TIPO_ALM & "' ) ", db, adOpenKeyset, adLockOptimistic
          If rs_aux7.RecordCount > 0 Then
             CORRELARTIVO1 = IIf(IsNull(rs_aux7!numero_correlativo), 1, rs_aux7!numero_correlativo + 1)
          Else
             MsgBox "No se puede generar el Correlativo por los privilegios del USUARIO", vbCritical, "sofia"
             Exit Sub
          End If
          rs_aux7!numero_correlativo = CORRELARTIVO1
          rs_aux7.Update
          'CORRELARTIVO1 = VAR_COD2          '"96"
          If parametro <> "COMEX" Then
            If IsNull(rs_datos!doc_numero) Or (rs_datos!doc_numero = 0) Then
                rs_datos!doc_numero = CORRELARTIVO1
                rs_datos.Update
            End If
             'rs_datos!doc_numero_alm = CORRELARIVO1
          End If
          'CORRELARTIVO1 = "526"
          If IsNull(Ado_detalle2.Recordset!doc_numero_alm) Or (Ado_detalle2.Recordset!doc_numero_alm = 0) Then
                Ado_detalle2.Recordset!doc_codigo_alm = "R-114"
                Ado_detalle2.Recordset!doc_numero_alm = CORRELARTIVO1
                Ado_detalle2.Recordset!estado_almacen = "APR"
                Ado_detalle2.Recordset.Update
                db.Execute "update ao_compra_cabecera set doc_numero_alm = " & CORRELARTIVO1 & " where compra_codigo = " & VAR_COD2 & " "
          End If
          'VAR_COD2 = Ado_datos.Recordset!compra_codigo
        Set rs_det1A = New ADODB.Recordset
        If rs_det1A.State = 1 Then rs_det1A.Close
        'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1A.MoveFirst
        'sino = rs_det1A.RecordCount
        'CORRELARIVO1 = "57"
        While Not rs_det1A.EOF
            'ac_tipo_compra_venta
            Set rs_aux12 = New ADODB.Recordset
            If rs_aux12.State = 1 Then rs_aux12.Close
            rs_aux12.Open "select * from ao_almacen_ingresos where ges_gestion='" & rs_det1A!ges_gestion & "' AND almacen_codigo=" & rs_det1A!almacen_codigo & " AND doc_codigo='" & Ado_datos.Recordset!doc_codigo_alm & "' AND doc_numero=" & CORRELARTIVO1 & " AND bien_codigo='" & rs_det1A!bien_codigo & "' ", db, adOpenStatic     '
            If rs_aux12.RecordCount > 0 Then
            Else
                db.Execute "ap_compras_grla 2,'" & rs_det1A!ges_gestion & "'," & rs_det1A!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "' ," & CORRELARTIVO1 & ",'" & rs_det1A!bien_codigo & "','" & Ado_datos.Recordset!edif_codigo & "'," & VAR_COD2 & ",'" & Ado_detalle2.Recordset!beneficiario_codigo & "','" & Ado_detalle2.Recordset!fecha_compra & "'," & rs_det1A!adjudica_cantidad & "," & rs_det1A!bien_total_adjudica_bs & "," & CDbl(rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!compra_DESCRIPCION & "'," & rs_det1A!bien_precio_adjudica_bs & ""
            End If

            Set rs_aux6 = New ADODB.Recordset
            If rs_aux6.State = 1 Then rs_aux6.Close
            rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & rs_det1A!almacen_codigo & " AND bien_codigo = '" & rs_det1A!bien_codigo & "'", db, adOpenStatic
            If rs_aux6.RecordCount > 0 Then
                db.Execute "ap_almacen_totales 2," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
            Else
                db.Execute "ap_almacen_totales 1," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG', '" & glusuario & "' "
                '15, '1.6.1.11', 1, 0, 1, 100, 0,0, 100/6.96, 0, 0, 'REG', 'ADMIN'
            End If
            rs_det1A!estado_codigo = "APR"
            rs_det1A.Update
            rs_det1A.MoveNext
        Wend
        
        If Ado_detalle2.Recordset!adjudica_codigo <> "" Then
            Set rs_det1A = New ADODB.Recordset
            If rs_det1A.State = 1 Then rs_det1A.Close
            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1A.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & VAR_COD2 & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1A.Sort = "compra_codigo_det"
            Set Ado_detalle1A.Recordset = rs_det1A
                If Ado_detalle1A.Recordset.RecordCount > 0 Then
                    'dg_det1A.Visible = True
                    Set dg_det1A.DataSource = Ado_detalle1A.Recordset
                    'Command1.Visible = True
                Else
                    'dg_det1A.Visible = False
                    ' Set Ado_detalle1A.Recordset = rsNada
                    Set dg_det1A.DataSource = rsNada
                    'Command1.Visible = False
                End If
        End If
          
    Else
        MsgBox "No se puede ACEPTAR, debe registrar al menos una Factura del Proveedor o Ya Fue Aceptado previamente ...", vbExclamation
    End If
  Else
    MsgBox "No se puede ACEPTAR, porque es SALDO INICIAL o no registr? la Factura del Proveedor o Ya Fue Aceptado previamente ...", vbExclamation
    
  End If
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexi?n = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atenci?n!"
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Est? Seguro de CANCELAR la operaci?n ? ", vbYesNo + vbQuestion, "Atenci?n")
   If sino = vbYes Then
        'rs_datos.CancelUpdate
        rs_datos.CancelBatch
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
'        If Ado_datos.Recordset!estado_codigo = "REG" Then
'            Call OptFilGral1_Click
'        Else
'            Call OptFilGral2_Click
'        End If
        'rs_datos.MoveFirst
        
    If VAR_SW = "MOD" Then
       var_cod = Ado_datos.Recordset!solicitud_codigo   'Codigo Llave de la Tabla
     End If


     'Call ABRIR_TABLA
     
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
        rs_datos.MoveFirst
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = '" & var_cod & "' ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
        mbDataChanged = False
        Fra_datos.Enabled = False
        Fra_datos.Visible = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        FraDet2.Visible = True
        FraDet1.Visible = True
'        FrmABMDet3.Visible = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True
        BtnSalir.Visible = True
        Call ABRIR_TABLA_DET
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset.RecordCount > 0 Then
    Select Case Glaux
             Case "PROVI"
             If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "TRANS"
             If Ado_datos.Recordset!estado_codigo_tra <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "ADUAN"
             If Ado_datos.Recordset!estado_codigo_nac <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "DESCA"
             If Ado_datos.Recordset!estado_codigo_des <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             Case "CONTR"
             If Ado_datos.Recordset!estado_codigo <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
             
             Case Else
              If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
             MsgBox "No se puede modificar este registro, porque este ya est? Anulado (ANL)!! ", vbExclamation
             Exit Sub
             End If
    End Select
   
       sino = MsgBox("Est? seguro de ANULAR el Registro ?" & vbCrLf & "Nota: Este registro no podra ser Reutilizado", vbYesNo + vbExclamation, "Atenci?n")
       If sino = vbYes Then
       
       var_cod = Ado_datos.Recordset!compra_codigo
       
       Select Case Glaux
             Case "PROVI"
               Ado_datos.Recordset!estado_codigo_eqp = "ANL"
               
             Case "TRANS"
               Ado_datos.Recordset!estado_codigo_tra = "ANL"
            
             Case "ADUAN"
               Ado_datos.Recordset!estado_codigo_nac = "ANL"
            
             Case "DESCA"
               Ado_datos.Recordset!estado_codigo_des = "ANL"
             
             Case "CONTR"
               Ado_datos.Recordset!estado_codigo = "ANL"
             
             Case Else
               Ado_datos.Recordset!estado_codigo_eqp = "ANL"
       End Select
          'rs_datos!estado_codigo_eqp = "ANL"
       rs_datos!fecha_registro = Date
       rs_datos!usr_codigo = glusuario
       rs_datos.UpdateBatch adAffectAll
       db.Execute "ap_compras_grla 1 ,'',0, '' ,0,'',''," & Ado_datos.Recordset!compra_codigo & ",'','',0,0,0, 'REG', '" & glusuario & "','',0"
       
'        Set rs_aux6 = New ADODB.Recordset
'        If rs_aux6.State = 1 Then rs_aux6.Close
'        rs_aux6.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo =" & rs_det1A!almacen_codigo & " AND bien_codigo = '" & rs_det1A!bien_codigo & "'", db, adOpenStatic
'        If rs_aux6.RecordCount > 0 Then
'            db.Execute "ap_almacen_totales 2," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
'        Else
'            db.Execute "ap_almacen_totales 1," & rs_det1A!almacen_codigo & ", '" & rs_det1A!bien_codigo & "', " & rs_det1A!adjudica_cantidad & ", '0', " & rs_det1A!adjudica_cantidad & ", " & rs_det1A!bien_total_adjudica_bs & ", 0, 0, " & rs_det1A!bien_total_adjudica_bs / GlTipoCambioOficial & ", 0, 0, 'REG', '" & glusuario & "' "
'        End If
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
        Call OptFilGral2_Click        'TODOS
        OptFilGral2.Value = True
'     End If

     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "compra_codigo = '" & var_cod & "' ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
          Call ABRIR_TABLA_DET
          
        VAR_SW = ""
     Else
        VAR_SW = ""
        rs_datos.MoveLast
     End If
          
       End If
       
    End If
  Exit Sub

UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Est? Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atenci?n")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validaci?n de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
 On Error GoTo UpdateErr
' db.BeginTrans
   VAR_UNI = dtc_codigo1.Text         'Aux
   VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        'db.BeginTrans
        Set rs_correl = New Recordset
        If rs_correl.State = 1 Then rs_correl.Close
        rs_correl.Open "Select * from gc_unidad_ejecutora WHERE unidad_codigo = '" & VAR_UNI & "'", db, adOpenKeyset, adLockOptimistic
        If rs_correl!correl_solicitud > 0 Then
            var_cod = rs_correl!correl_solicitud + 1
            rs_datos!solicitud_codigo_adm = var_cod
            rs_correl!correl_solicitud = var_cod
            rs_datos!solicitud_codigo = var_cod
        Else
            var_cod = "1"
            rs_datos!solicitud_codigo_adm = var_cod
            rs_correl!correl_solicitud = var_cod
            rs_datos!solicitud_codigo = var_cod
        End If
        rs_correl.Update
        'VAR_COMPRA = Ado_datos.Recordset!compra_codigo  'Codigo Llave de la Tabla
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(CDbl(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' Year(Date)   'no cambia
        rs_datos!unidad_codigo = VAR_UNI
        rs_datos!unidad_codigo_adm = VAR_UNI
        'Actualiza correaltivo ...
        'db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!edif_codigo = dtc_codigo3.Text
        rs_datos!depto_codigo = Left(Trim(dtc_codigo3.Text), 1)
        'PARAMETROS POR TIPO DE TRAMITE
        Select Case Aux
            Case "30"                           'Aux = "30" 'CCD
                rs_datos!trans_codigo_EGR = "47"
                VAR_TIPO = "CGOCTA"
                VAR_DOCF = "R-160"              'R-160   Formulario Cargo de Cuenta
                CodBien = "FONDO03"
            Case "29"                           'Aux = "29" 'VIAJES
                rs_datos!trans_codigo_EGR = "46"
                VAR_TIPO = "FVIAJE"
                VAR_DOCF = "R-162"              'R-162   Formularios para Viajes
                CodBien = "FONDO02"
            Case "28"                           'Aux = "28" 'CAJA CHICA
                rs_datos!trans_codigo_EGR = "45"
                VAR_TIPO = "CCHICA"
                VAR_DOCF = "R-161"              'R-161   Formulario Caja Chica
                CodBien = "FONDO01"
            Case Else
                rs_datos!trans_codigo_EGR = "47"
                VAR_TIPO = "CGOCTA"
                VAR_DOCF = "R-160"
                CodBien = "FONDO03"
        End Select
        rs_datos!codigo_empresa = dtc_codigo10.Text
        'CORRELATIVO DEL TIPO DE TRAMITE
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOCF & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            If dtc_codigo10.Text = "2" Then
                rs_aux2!doc_nro_copias = rs_aux2!doc_nro_copias + 1
                Txt_campo1.Caption = rs_aux2!doc_nro_copias
                VAR_CODF = rs_aux2!doc_nro_copias
            Else
                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                Txt_campo1.Caption = rs_aux2!correl_doc
                VAR_CODF = rs_aux2!correl_doc
            End If
            rs_aux2.Update
        End If
        'db.Execute "UPDATE ao_compra_cabecera SET doc_numero = " & rs_aux7!correl_ing + 1 & " WHERE compra_codigo = " & fw_compras_fondos.Ado_detalle2.Recordset!compra_codigo & ""
        rs_datos!doc_codigo = VAR_DOCF
        rs_datos!doc_numero = VAR_CODF
        'rs_datos!venta_tipo = VAR_TIPO_ALM
        If VAR_CODF < 10 Then
           rs_datos!unidad_codigo_ant = VAR_TIPO + "-0000" + Trim(txt_codigo)
        End If
        If VAR_CODF > 9 And VAR_CODF < 100 Then
           rs_datos!unidad_codigo_ant = VAR_TIPO + "-000" + Trim(txt_codigo)
        End If
        If VAR_CODF > 99 And VAR_CODF < 1000 Then
           rs_datos!unidad_codigo_ant = VAR_TIPO + "-00" + Trim(txt_codigo)
        End If
        If VAR_CODF > 999 And VAR_CODF < 10000 Then
           rs_datos!unidad_codigo_ant = VAR_TIPO + "-0" + Trim(txt_codigo)
        End If
        If VAR_CODF > 9999 And VAR_CODF < 100000 Then
           rs_datos!unidad_codigo_ant = VAR_TIPO + "-" + Trim(txt_codigo)
        End If
        'If VAR_CODF > 99999 Then
        '   rs_datos!unidad_codigo_ant = VAR_TIPO + "-" + Trim(txt_codigo)
        'End If
     End If
     If VAR_SW = "MOD" Then
        sino = Ado_datos.Recordset!compra_codigo
        var_cod = rs_datos!compra_codigo   'Codigo Llave de la Tabla
        VAR_COMPRA = rs_datos!solicitud_codigo  'Codigo Llave de la Tabla
        If CodBien = "" Then
            Select Case Aux
                Case "30"                           'Aux = "30" 'CCD
                    rs_datos!trans_codigo_EGR = "47"
                    VAR_DOCF = "R-160"              'R-160   Formulario Cargo de Cuenta
                    CodBien = "FONDO03"
                Case "29"                           'Aux = "29" 'VIAJES
                    rs_datos!trans_codigo_EGR = "46"
                    VAR_DOCF = "R-162"              'R-162   Formularios para Viajes
                    CodBien = "FONDO02"
                Case "28"                           'Aux = "28" 'CAJA CHICA
                    rs_datos!trans_codigo_EGR = "45"
                    VAR_DOCF = "R-161"              'R-161   Formulario Caja Chica
                    CodBien = "FONDO01"
                Case Else
                    rs_datos!trans_codigo_EGR = "47"
                    VAR_DOCF = "R-160"
                    CodBien = "FONDO03"
            End Select
        End If
     End If
     rs_datos!compra_fecha = DTPfecha1.Value
     rs_datos!ges_gestion = Year(DTPfecha1.Value)
     If dtc_codigo10.Text = "" Or dtc_codigo10.Text = "0" Then
        dtc_codigo10.Text = "1"
     End If
     If dtc_codigo10.Text = "2" Then
        rs_datos!venta_tipo = "B"
     Else
        rs_datos!venta_tipo = "F"
     End If
     rs_datos!solicitud_tipo = Aux
     rs_datos!solicitud_tipo_egr = 1
     rs_datos!clasif_codigo = "ADM"
     rs_datos!poa_codigo = "4.2.3"
     rs_datos!proceso_codigo = "FIN"
     rs_datos!subproceso_codigo = "FIN-03"
     rs_datos!etapa_codigo = "FIN-03-02"
     rs_datos!compra_DESCRIPCION = Txt_descripcion.Text
     If txt_obs.Text = "" Then
        rs_datos!compra_observaciones = dtc_desc3.Text   'txt_obs.Text
     Else
        rs_datos!compra_observaciones = txt_obs.Text
     End If
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     rs_datos!beneficiario_codigo_alm = IIf(IsNull(dtc_codigo4.Text), 0, dtc_codigo4.Text)
     rs_datos!beneficiario_codigo = IIf(IsNull(dtc_codigo4.Text), 0, dtc_codigo4.Text)
     
     rs_datos!nro_nota_remision = "0"
     rs_datos!compra_cantidad_total = "1"
     rs_datos!compra_monto_bs = IIf(lbl_total_bs.Text = "", 0, CDbl(lbl_total_bs.Text))
     rs_datos!compra_monto_DOL = Round(IIf(lbl_total_dol.Text = "", 0, CDbl(lbl_total_dol.Text)), 2)
     rs_datos!tipo_moneda = "BOB"

     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_registro_aprueba = Date
     rs_datos!estado_codigo_tec = "REG"
     rs_datos!doc_numero_alm = "1"
     rs_datos!adjudica_codigo = "0"
     rs_datos!doc_codigo_alm = "R-126"
     'hora_registro

'     Select Case Glaux
'             Case "PROVI"
'            'rs_datos!estado_codigo_eqp = "REG"
'             Case "TRANS"
'             'rs_datos!estado_codigo_tra = "REG"
'             Case "ADUAN"
'             '  rs_datos!estado_codigo_nac = "REG"
'             Case "DESCA"
'             '  rs_datos!estado_codigo_des = "REG"
'             Case "CONTR"
'             '  rs_datos!estado_codigo = "REG"
'             Case Else
'                rs_datos!estado_codigo_eqp = "REG"
'                rs_datos!estado_codigo = "REG"
'    End Select

     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!estado_codigo_eqp = "REG"
     rs_datos!estado_codigo_tra = "REG"
     rs_datos!estado_codigo_nac = "REG"
     rs_datos!estado_codigo_des = "REG"
     rs_datos!estado_codigo = "REG"
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     'Ado_datos.Recordset.Update
     rs_datos.Update    'Batch 'adAffectAll
     'db.CommitTrans
     var_cod = rs_datos!compra_codigo   'Codigo Llave de la Tabla
'     ' GRABA DETALLE
'     If Ado_detalle1.Recordset.RecordCount = 0 Then
'        db.Execute "insert into ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,             grupo_codigo, subgrupo_codigo, " & _
'             " par_codigo, tipo_descuento, almacen_codigo, usr_usuario,    fecha_registro, hora_registro, unimed_codigo, estado_codigo, adjudica_monto_bs_87, compra_tdc, tipo_moneda) " & _
'             " VALUES ('" & glGestion & "', " & var_cod & ", '" & CodBien & "', '1', " & CDbl(lbl_total_bs.Text) & ",      '0',            " & CDbl(lbl_total_bs.Text) & ", " & CDbl(lbl_total_dol.Text) & ",           '0',   " & CDbl(lbl_total_dol.Text) & ", '" & Txt_descripcion.Text & "', '20000', '26000',  " & _
'             " '26990',    '0',            '15',        '" & glusuario & "', '" & Date & "', '0',       'UNI',            'REG', " & CDbl(lbl_total_bs.Text) * 0.87 & ", '6.96', 'BOB'  )"
'     Else
'     End If
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If

     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
        rs_datos.Find "compra_codigo = " & var_cod & " ", , , 1
        'rs_datos.Find "solicitud_codigo = " & VAR_COMPRA & " ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
          lbl_total_bs.Text = "0"
          lbl_total_dol.Text = "0"
          Call ABRIR_TABLA_DET
          
        VAR_SW = ""
     Else
        VAR_SW = ""
        rs_datos.MoveLast
     End If
     
'     If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Call OptFilGral1_Click
'     Else
'        Call OptFilGral2_Click
'     End If
'     rs_datos.MoveLast
     mbDataChanged = False

     Fra_datos.Enabled = False
     Fra_datos.Visible = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     
        FraDet2.Visible = True
        FraDet1.Visible = True
    dtc_desc3.Locked = False

  End If
  
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
  'db.RollbackTrans
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo8.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo10.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validaci?n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
 If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        CR01.ReportFileName = App.Path & "\Reportes\Comex\rr_proceso_contratacion.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!solicitud_codigo
        CR01.StoredProcParam(1) = Glaux
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresi?n"
        CR01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atenci?n"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atenci?n"
  End If
End Sub

Private Sub BtnImprimir4_Click()
    CR04.Reset
    CR04.WindowState = crptMaximized
    CR04.WindowShowSearchBtn = True
    CR04.WindowShowRefreshBtn = True
    CR04.WindowShowPrintSetupBtn = True
    If GlBaseDatos = "ADMIN_EMPRESA" Then
       CR04.ReportFileName = App.Path & "\Reportes\COMEX\ar_compra_adjudica_FAC.rpt"
    Else
        CR04.ReportFileName = App.Path & "\Reportes\COMEX\ar_compra_adjudica_FAC_PRUEBA.rpt"
    End If
    'CR02.Formulas(5) = "Titulo = 'ORDEN DE PAGO'"
    'CR02.Formulas(6) = "Subtitulo = '" & lbl_titulo.Caption & "'"
    'CR02.Formulas(0) = "mes_asistencia ='" & UCase(MonthName(txt_mes.Text)) & "'"
    CR04.StoredProcParam(0) = Ado_datos.Recordset!compra_codigo                 'VAR_COMPRA
    'CR02.StoredProcParam(1) = Ado_detalle2.Recordset!adjudica_codigo
    'CR02.StoredProcParam(2) = Ado_datos.Recordset!ges_gestion
  iResult = CR04.PrintReport
  If iResult <> 0 Then
    MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub BtnModDetalle1_Click()
'On Error GoTo UpdateErr
 
    If Ado_datos.Recordset.RecordCount > 0 Then
             If Ado_datos.Recordset!estado_codigo_eqp <> "REG" Then
                MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
                Exit Sub
             End If
    If Ado_detalle1.Recordset.RecordCount > 0 Then
      If rs_datos.RecordCount > 0 And Ado_detalle1.Recordset!estado_codigo = "REG" Then
        If Ado_detalle2.Recordset.RecordCount > 0 Then
            Set rs_aux8 = New ADODB.Recordset
               If rs_aux8.State = 1 Then rs_aux8.Close
               'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
               rs_aux8.Open "select * from ao_compra_adjudica_bienes where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
               'rs_det1A.Sort = "compra_codigo_det"
            If rs_aux8.RecordCount > 0 Then
                sino = MsgBox("este bien ya esta con un proveedor, para modificar quite el bien del proveedor ", vbInformation, "AVISO")
                Exit Sub
            End If
        End If
      
      
        marca1 = Ado_detalle1.Recordset.Bookmark
        swnuevo = 2
        VAR_SW = "MOD"
        fraOpciones.Visible = False
        fraOpcionesDet.Visible = False
        FraNavega.Enabled = False
        FraDet2.Enabled = False
        FrmABMDet2.Visible = False
        FraDet1.Enabled = False
            
''        Select Case dtc_codigo2.Text
''            Case "1"    'SOLO COMPRAS BB y SS
''            Case "2"    'SOLO VENTA DE BIENES
''            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
''            Case "L"    'IMPORTACION DIRECTA CLIENTE
''                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
''                frm_solicitud_bienes.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
''                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
''
''                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
''                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
''
''                frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
''                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
''                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
''                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
''
''                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
''                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
''                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
''                frm_solicitud_bienes.Txt_campo19.Text = Me.Ado_detalle1.Recordset("bien_cantidad_por_empaque")
''
''                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
''                frm_solicitud_bienes.Txt_campo15.Text = "10" 'Me.Ado_detalle1.Recordset("fosa_dimension_frente")
''
''                frm_solicitud_bienes.lbl_det.Caption = "43340"
''                frm_solicitud_bienes.Show vbModal
''            Case "V"    'FACTURACION LOCAL
''                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
''               frm_solicitud_bienes.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
''                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
''
''                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
''                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
''
''                frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
''                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
''                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
''                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
''
''                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
''                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
''                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
''
''                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
''    '            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
''    '               frm_solicitud_bienes.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
''                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
''
''                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
''                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
''
''                frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
''                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
''                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
''                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
''
''                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
''                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
''                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
''
''                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
''    '            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
''    '            frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
''                frm_solicitud_bienes.lbl_det.Caption = "43340"
''                frm_solicitud_bienes.Show vbModal
''
''        End Select
''               'frm_solicitud_bienes_gral.dtc_codigo2 = Me.Ado_detalle1.Recordset!unimed_codigo
                If Me.Ado_detalle1.Recordset("almacen_codigo") <> "NULL" And parametro <> "COMEX" Then
                    frm_solicitud_bienes_gral.dtc_desc_alm.BoundText = Me.Ado_detalle1.Recordset("almacen_codigo")
                End If
                frm_solicitud_bienes_gral.Txt_campo1.Caption = dtc_codigo1.Text   'Unidad
                frm_solicitud_bienes_gral.dtc_desc1.BoundText = Me.Ado_detalle1.Recordset("bien_codigo")
                
                frm_solicitud_bienes_gral.dtc_desc1.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.dtc_aux1.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.dtc_aux2.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.dtc_aux3.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.Txt_campo2.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.Txt_campo3.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.Txt_campo4.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.Txt_campo18.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.dtc_codigo2.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                frm_solicitud_bienes_gral.Txt_campo14.BoundText = frm_solicitud_bienes_gral.dtc_codigo1.BoundText
                
                'frm_solicitud_bienes_gral.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("bien_codigo")
                frm_solicitud_bienes_gral.Txt_campo10.Text = Format(Me.Ado_detalle1.Recordset("compra_precio_unitario_bs"), "###,###,##0.00")
                frm_solicitud_bienes_gral.Txt_campo11.Text = Format(Me.Ado_detalle1.Recordset("compra_precio_total_bs"), "###,###,##0.00")
                frm_solicitud_bienes_gral.Text2.Text = Format(IIf(IsNull(Me.Ado_detalle1.Recordset("compra_precio_total_dol")), 0, Me.Ado_detalle1.Recordset("compra_precio_total_dol")), "###,###,##0.00")
                
                frm_solicitud_bienes_gral.Txt_campo16.Text = Me.Ado_detalle1.Recordset("compra_cantidad") 'dtc_codigo2
                'frm_solicitud_bienes_gral.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("compra_cantidad")
                frm_solicitud_bienes_gral.MOD_NEW.Caption = "MOD"
                'frm_solicitud_bienes_gral.dtc_desc2.BoundText = IIf(IsNull(Me.Ado_detalle1.Recordset("unimed_codigo")), EQP, Me.Ado_detalle1.Recordset("unimed_codigo"))
                frm_solicitud_bienes_gral.txt_gestion = Me.Ado_detalle1.Recordset("ges_gestion")
                If Ado_detalle1.Recordset!compra_tdc < 2 Or IsNull(Ado_detalle1.Recordset!compra_tdc) Then
                    frm_solicitud_bienes_gral.Txt_tdc.Text = "6.96"
                Else
                    frm_solicitud_bienes_gral.Txt_tdc.Text = Ado_detalle1.Recordset!compra_tdc
                End If
                'frm_solicitud_bienes_gral.lbl_edif.Caption = dtc_codigo3.Text
              
                    
'                frm_solicitud_bienes_gral.Txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
'                frm_solicitud_bienes_gral.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
'                frm_solicitud_bienes_gral.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
'                frm_solicitud_bienes_gral.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes_gral.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
'                frm_solicitud_bienes_gral.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes_gral.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
'
'                frm_solicitud_bienes_gral.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
'    '           frm_solicitud_bienes_gral.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
'    '           frm_solicitud_bienes_gral.dtc_desc2.BoundText = frm_solicitud_bienes_gral.dtc_codigo2.BoundText
'                frm_solicitud_bienes_gral.lbl_det.Caption = "43340"
               
                frm_solicitud_bienes_gral.Show vbModal
        
'        swnuevo = 0
        fraOpciones.Visible = True
        fraOpcionesDet.Visible = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Visible = True
        FraDet1.Enabled = True
        BtnSalir.Visible = True
    
       ' Ado_detalle1.Recordset.Move marca1 - 1
   
      Else
        MsgBox "No se puede MODIFICAR, porque ya est? APROBADO o ANULADO, Verifique por favor!! ", vbExclamation
      End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validaci?n de Registro"
  End If
 DETALLE1 = Ado_detalle1.Recordset!compra_codigo_det
       
VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
 Call ABRIR_TABLA_DET
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     End If
     
         
  If (dg_det1.SelBookmarks.Count <> 0) Then
        dg_det1.SelBookmarks.Remove 0
     End If
     If Ado_detalle1.Recordset.RecordCount > 0 Then
        rs_det1.Find "compra_codigo_det = " & DETALLE1 & "   ", , , 1
        dg_det1.SelBookmarks.Add (rs_det1.Bookmark)
End If
End If
End Sub

Private Sub BtnModDetalle2_Click()
On Error GoTo UpdateErr
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo <> "REG" Then          'estado_codigo_eqp
        MsgBox "No se puede modificar este registro, porque este ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
        Exit Sub
    End If
    BtnAprobar1.Visible = True
'    BtnAprobar3.Visible = True
    VAR_SW = "MOD"
    sw_nuevo = "MOD"
    marca1 = Ado_datos.Recordset.Bookmark
'   If rs_det2.RecordCount > 0 Then
'   Exit Sub
'   End If
 If Ado_detalle2.Recordset.RecordCount = 0 Then
 
    Exit Sub
 End If
 sino = Ado_detalle2.Recordset!compra_codigo

  If rs_datos.RecordCount > 0 And Ado_detalle2.Recordset!estado_codigo = "REG" Then
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        swnuevo = 2
        fraOpciones.Visible = False
        fraOpcionesDet.Visible = False
        FraNavega.Enabled = False
        FraDet2.Visible = False
        FrmABMDet2.Visible = False
        '    'Call ABRIR_TABLA_DET
        'ges_gestion,     adjudica_fecha, proceso_codigo, subproceso_codigo, etapa_codigo,
        'clasif_codigo, doc_codigo, doc_numero,  adjudica_descripcion, adjudica_cantidad_total,  tipo_moneda,
        '    fecha_recibe_almacen, almacen_codigo, poa_codigo, estado_codigo,
         'usr_codigo , fecha_registro, hora_registro, usr_codigo_aprueba, fecha_aprueba
            fw_adjudica_fondos.txtSW.Text = swnuevo                 'Me.Ado_datos.Recordset!venta_tipo
            fw_adjudica_fondos.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
            fw_adjudica_fondos.Txt_campo1.Text = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
            fw_adjudica_fondos.Txt_descripcion.Caption = Me.dtc_desc1.Text
            fw_adjudica_fondos.txtCodigo1.Caption = Me.Ado_detalle2.Recordset("compra_codigo")
            'fw_adjudica_fondos.Txt_estado.Caption = "REG"
            
            fw_adjudica_fondos.lbl_adjudica.Caption = Me.Ado_detalle2.Recordset("adjudica_codigo")
            fw_adjudica_fondos.dtc_codigo5.Text = Me.Ado_detalle2.Recordset("beneficiario_codigo")
            fw_adjudica_fondos.dtc_desc5.BoundText = fw_adjudica_fondos.dtc_codigo5.BoundText
            fw_adjudica_fondos.dtc_aux4.BoundText = fw_adjudica_fondos.dtc_codigo5.BoundText
            fw_adjudica_fondos.dtc_aux5.BoundText = fw_adjudica_fondos.dtc_codigo5.BoundText

            fw_adjudica_fondos.txt_Nota.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("nro_nota_remision")), "", Me.Ado_detalle2.Recordset("nro_nota_remision"))
            fw_adjudica_fondos.txt_total_bs.Text = Format(IIf(IsNull(Me.Ado_detalle2.Recordset("adjudica_monto_bs")), 0, Me.Ado_detalle2.Recordset("adjudica_monto_bs")), "###,###,##0.00")
            fw_adjudica_fondos.txt_total_dol.Text = Format(IIf(IsNull(Me.Ado_detalle2.Recordset!adjudica_monto_dol), 0, Me.Ado_detalle2.Recordset!adjudica_monto_dol), "###,###,##0.00")
            fw_adjudica_fondos.txtFecha.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_inicio_contrato")), Date, Me.Ado_detalle2.Recordset("fecha_inicio_contrato"))
            fw_adjudica_fondos.txtFecha2.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_fin_contrato")), Date, Me.Ado_detalle2.Recordset("fecha_fin_contrato"))
            fw_adjudica_fondos.txtFecha3.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_envio_proveedor")), Date, Me.Ado_detalle2.Recordset("fecha_envio_proveedor"))
            
            fw_adjudica_fondos.cmb_mes_ini = IIf(IsNull(Me.Ado_detalle2.Recordset!mes_inicio_crono), UCase(MonthName(Month(Date))), Me.Ado_detalle2.Recordset!mes_inicio_crono)
            fw_adjudica_fondos.txtCantCuota.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!cantidad_cuotas_pag), "1", Me.Ado_detalle2.Recordset!cantidad_cuotas_pag)
            fw_adjudica_fondos.cmd_unimed2 = IIf(IsNull(Me.Ado_detalle2.Recordset!unimed_codigo_pag), "MES", Me.Ado_detalle2.Recordset!unimed_codigo_pag)
            
            fw_adjudica_fondos.txtSW.Text = swnuevo                 'Me.Ado_datos.Recordset!venta_tipo
            If Me.Ado_detalle2.Recordset("almacen_codigo") <> "NULL" Then
            fw_adjudica_fondos.dtc_desc_alm.BoundText = Me.Ado_detalle2.Recordset("almacen_codigo")
            End If
            fw_adjudica_fondos.txt_pais.Text = VAR_PAIS
            fw_adjudica_fondos.txtfecha_compra.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!fecha_compra), Date, Me.Ado_detalle2.Recordset!fecha_compra)
            fw_adjudica_fondos.txt_autorizacion.Text = IIf(IsNull(Ado_detalle2.Recordset("nro_autorizacion")), "", Ado_detalle2.Recordset("nro_autorizacion"))
            fw_adjudica_fondos.txt_cod_control.Text = IIf(IsNull(Ado_detalle2.Recordset("codigo_control")), "", Ado_detalle2.Recordset("codigo_control"))
            
            fw_adjudica_fondos.txt_nro_dui.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("nro_dui")), 0, Me.Ado_detalle2.Recordset("nro_dui"))
             fw_adjudica_fondos.txt_importe_no_fiscal.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("importe_no_credito_fisc")), 0, Me.Ado_detalle2.Recordset("importe_no_credito_fisc"))
            fw_adjudica_fondos.txt_descuentos.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("descuento")), 0, Me.Ado_detalle2.Recordset("descuento"))
            If fw_adjudica_fondos.txt_tipo_cambio = "0" Or fw_adjudica_fondos.txt_tipo_cambio = "" Then
            fw_adjudica_fondos.txt_tipo_cambio = GlTipoCambioOficial
            Else
            fw_adjudica_fondos.txt_tipo_cambio = Ado_detalle2.Recordset("tipo_cambio")
            End If
            If Ado_detalle2.Recordset("tipo_moneda") = "USD" Then
            fw_adjudica_fondos.opt_usd.Value = True
            End If
             If Ado_detalle2.Recordset("tipo_moneda") = "BOB" Then
            fw_adjudica_fondos.opt_bs.Value = True
            End If
            If fw_compras_fondos.Ado_detalle2.Recordset("factura") = "NO" Then
            fw_adjudica_fondos.opt_no.Value = True
            Else
            fw_adjudica_fondos.opt_si.Value = True
            End If
            fw_adjudica_fondos.txtFecha.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!fecha_inicio_contrato), Me.Ado_detalle2.Recordset!fecha_compra, Me.Ado_detalle2.Recordset!fecha_inicio_contrato)
             fw_adjudica_fondos.txtFecha2.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!fecha_fin_contrato), Me.Ado_detalle2.Recordset!fecha_compra, Me.Ado_detalle2.Recordset!fecha_fin_contrato)
             fw_adjudica_fondos.txtFecha3.Value = IIf(IsNull(Me.Ado_detalle2.Recordset!fecha_envio_proveedor), Me.Ado_detalle2.Recordset!fecha_compra, Me.Ado_detalle2.Recordset!fecha_envio_proveedor)
             
'          If parametro = "COMEX" Then
'            fw_adjudica_fondos.txtFecha.Visible = True
'            fw_adjudica_fondos.txtFecha2.Visible = True
'            fw_adjudica_fondos.txtFecha3.Visible = True
'            fw_adjudica_fondos.txtFecha.Value = Date
'            fw_adjudica_fondos.txtFecha2.Value = Date
'            fw_adjudica_fondos.txtFecha3.Value = Date
'            fw_adjudica_fondos.txt_nro_dui.Enabled = True
'            fw_adjudica_fondos.lblbien(2).Visible = True
'            fw_adjudica_fondos.lblbien(3).Visible = True
'            fw_adjudica_fondos.lblbien(4).Visible = True
'            Else
'            fw_adjudica_fondos.txtFecha.Visible = False
'            fw_adjudica_fondos.txtFecha2.Visible = False
'            fw_adjudica_fondos.txtFecha3.Visible = False
'            fw_adjudica_fondos.txt_nro_dui.Enabled = False
'            fw_adjudica_fondos.lblbien(2).Visible = False
'            fw_adjudica_fondos.lblbien(3).Visible = False
'            fw_adjudica_fondos.lblbien(4).Visible = False
'            End If
            fw_adjudica_fondos.Show vbModal
'        swnuevo = 0
        fraOpciones.Visible = True
        fraOpcionesDet.Visible = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Visible = True
        BtnSalir.Visible = True
     Else
        MsgBox "No se puede Modificar un registro inexistente, vuelva a intentar!! ", vbExclamation
        Exit Sub
     End If
  Else
    MsgBox "No se puede Modificar el registro, porque este ya est? Aprobado (APR) 0 Aulado (ANL)!! ", vbExclamation
  End If
DETALLE1 = Ado_detalle1.Recordset!compra_codigo_det
DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
'
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
        rs_datos.Find "solicitud_codigo = " & VAR_COD2 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     
     
       If (dg_det1.SelBookmarks.Count <> 0) Then
        dg_det1.SelBookmarks.Remove 0
     End If
     If Ado_detalle1.Recordset.RecordCount > 0 Then
        rs_det1.Find "compra_codigo_det = " & DETALLE1 & "   ", , , 1
        dg_det1.SelBookmarks.Add (rs_det1.Bookmark)
     Else
        rs_det1.MoveLast
     End If
     
     
     If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     Else
        rs_det2.MoveLast
     End If

     

End If
Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  
  If Ado_datos.Recordset.RecordCount > 0 Then
    '  lblStatus.Caption = "Modificar registro"
   ' If Ado_datos.Recordset!estado_codigo = "REG" Then
    If Ado_datos.Recordset!estado_codigo_eqp = "REG" Then
        Fra_datos.Visible = True
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        dtc_desc3.Locked = True
        FraDet2.Visible = False
        FraDet1.Visible = False
        dtc_codigo1.Text = Ado_datos.Recordset!unidad_codigo
        dtc_desc1.BoundText = dtc_codigo1.BoundText
        dtc_aux1.BoundText = dtc_codigo1.BoundText
        dtc_codigo3.Text = Ado_datos.Recordset!edif_codigo
        dtc_desc3.BoundText = dtc_codigo3.BoundText
        dtc_codigo4.Text = Ado_datos.Recordset!beneficiario_codigo_alm
        dtc_desc_ben.BoundText = dtc_codigo4.BoundText
        dtc_codigo11.Text = Ado_datos.Recordset!beneficiario_codigo_resp
        dtc_desc11.BoundText = dtc_codigo11.BoundText
        dtc_codigo2.Text = Ado_datos.Recordset!venta_tipo
        dtc_desc2.BoundText = dtc_codigo2.BoundText
        VAR_SW = "MOD"

        dtc_desc11.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
        FraGrabarCancelar.Visible = True
        BtnCancelar.Visible = True
    Else
        MsgBox "No se puede MODIFICAR, el registro ya est? Aprobado (APR) o Anulado (ANL)!! ", vbExclamation
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atenci?n!"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub BtnVer_Click()
  On Error GoTo QError
  If rs_datos!estado_codigo = "APR" Then
    Dim ARCH_FOTO As String
    Dim SW0 As String
    Select Case Left(Trim(Ado_datos.Recordset("edif_codigo")), 1)
        Case "1"    'CHQ
            VAR_DPTO = "CHQ"
        Case "2"    'LPZ
            VAR_DPTO = "LPZ"
        Case "3"    'CBB
            VAR_DPTO = "CBB"
        Case "4"    'SCZ
            VAR_DPTO = "SCZ"
        Case "5"    'PTS
            VAR_DPTO = "PTS"
        Case "6"    'ORU
            VAR_DPTO = "ORU"
        Case "7"    'TJA
            VAR_DPTO = "TJA"
        Case "8"    'BEN
            VAR_DPTO = "BEN"
        Case "9"    'PDO
            VAR_DPTO = "PDO"
    End Select
    If Ado_datos.Recordset!archivo_respaldo_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "DED2"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      'negocia_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atenci?n")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        'e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    '    If SW0 = 1 Then
    '    '    If GlServidor = "SRVPRO" Then
    '    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    '    Else
    '            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo)
    '            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    '    End If
    '        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    '        CodBien = Ado_datos.Recordset!edif_codigo
    '        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
    '            MsgBox "Se cargo la Imagen Correctamente !!"
    '        Else
    '            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    '        End If
    '    Else
    '        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    '        Image2 = Img_Foto
    '    End If
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validaci?n de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atenci?n"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub BtnDesAprobar3_Click()
On Error GoTo UpdateErr
 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        If Ado_detalle1A.Recordset.RecordCount > 0 Then
            If Ado_detalle2.Recordset!estado_codigo = "APR" Then
                sino = MsgBox("No Se Puede Quitar Este ITEM, Ya Esta APROBADO(APR) como Ingreso a Almac?n...", vbCritical, "ERROR")
                Exit Sub
            End If
            If (Ado_datos.Recordset!edif_codigo = "20101-3") Or (Ado_datos.Recordset!edif_codigo = "30101-3") Or (Ado_datos.Recordset!edif_codigo = "70101-3") Or (Ado_datos.Recordset!edif_codigo = "10101-3") Then
                db.Execute "update ao_almacen_TOTALES SET stock_ingreso = stock_ingreso - " & Ado_detalle1A.Recordset!adjudica_cantidad & " WHERE WHERE almacen_codigo = " & Ado_detalle1A.Recordset!almacen & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "'  "
                db.Execute "DELETE ao_almacen_ingresos WHERE compra_codigo = " & Ado_detalle1A.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "' "
                db.Execute "update ao_compra_detalle set estado_codigo = 'REG' WHERE compra_codigo = " & Ado_detalle1A.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "' "
                db.Execute "DELETE ao_compra_adjudica_bienes WHERE compra_codigo = " & Ado_detalle1A.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "' "
            Else
                db.Execute "update ao_compra_detalle set estado_codigo = 'REG' WHERE compra_codigo = " & Ado_detalle1A.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "' "
                db.Execute "DELETE ao_compra_adjudica_bienes WHERE compra_codigo = " & Ado_detalle1A.Recordset!compra_codigo & " AND bien_codigo = '" & Ado_detalle1A.Recordset!bien_codigo & "' "
            End If
            DETALLE2 = Ado_detalle2.Recordset!adjudica_codigo
                
            Call ABRIR_TABLA_DET
            'ao_almacen_ingresos
            If (dg_det2.SelBookmarks.Count <> 0) Then
               dg_det2.SelBookmarks.Remove 0
            End If
            If Ado_detalle2.Recordset.RecordCount > 0 Then
               rs_det2.Find "adjudica_codigo = " & DETALLE2 & "   ", , , 1
               dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
            Else
               rs_det2.MoveLast
            End If
        End If
    End If
 End If
 Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub Command2_Click()
  Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    rs_det2.Open "select * from ao_compra_adjudica", db, adOpenKeyset, adLockOptimistic, adCmdText '"select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    sino = rs_det2.RecordCount
    rs_det2.MoveFirst
    
    While Not rs_det2.EOF
    
'   If rs_det2!bien_codigo = "479" Then
'   rs_det2!literal_neto = Literal(rs_det2!importe_cred_fisc - rs_det2!credito_fiscal_13)
'   Else
If rs_det2!adjudica_monto_bs_87 <> "NULL" Then
    rs_det2!literal_neto = Literal(rs_det2!adjudica_monto_bs_87)
End If
   
    rs_det2.MoveNext
    Wend
End Sub


Private Sub BtnImprimir3_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_detalle2.Recordset.RecordCount = 0 Then
        MsgBox "No se puede IMPRIMIR, debe elegir al menos una Factura del Proveedor ", vbExclamation
        Exit Sub
    End If
    If Ado_detalle2.Recordset!estado_codigo = "ANL" Then       '<>
       
       COUNTER = 0
       sino = MsgBox("El registro NO debe estar ANULADO(ANL) Para Poder Imprimir", vbInformation, "SOFIA")
       Exit Sub
    End If
    CR02.Reset
    CR02.WindowState = crptMaximized
    CR02.WindowShowSearchBtn = True
    CR02.WindowShowRefreshBtn = True
    CR02.WindowShowPrintSetupBtn = True
    'If Ado_detalle1.Recordset.RecordCount > 15 Then
    '   CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes_pag1.rpt"
    'Else
    '   CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes.rpt"
    'End If
    CR02.ReportFileName = App.Path & "\Reportes\compras\ar_descargos_fondos.rpt"
    
    'CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes.rpt"
    'If parametro = "UALMI" Or parametro = "ALMIB" Or parametro = "ALMIC" Or parametro = "ALMIS" Then
    'CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes.rpt"
    'End If
    '
    'If parametro = "UALMR" Or parametro = "ALMRB" Or parametro = "ALMRC" Or parametro = "ALMRS" Then
    'CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes_rep.rpt"
    'End If
    '
    'If parametro = "UALMR" Or parametro = "ALMRB" Or parametro = "ALMRC" Or parametro = "ALMRS" Then
    'CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes_rep.rpt"
    'End If
    'If parametro = "UALMH" Or parametro = "ALMHB" Or parametro = "ALMHC" Or parametro = "ALMHS" Then
    'CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_ingreso_almacenes_her.rpt"
    'End If
    
    CR02.Formulas(5) = "Titulo = 'FONDOS A RENDIR' "
    CR02.Formulas(6) = "Subtitulo = '" & lbl_titulo.Caption & "'"
    'CR02.Formulas(0) = "mes_asistencia ='" & UCase(MonthName(txt_mes.Text)) & "'"
    CR02.StoredProcParam(0) = Ado_datos.Recordset!compra_codigo
    CR02.StoredProcParam(1) = Ado_detalle2.Recordset!adjudica_codigo
    CR02.StoredProcParam(2) = Ado_datos.Recordset!ges_gestion
  
  
  iResult = CR02.PrintReport
  If iResult <> 0 Then
    MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End If
End Sub


Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    'dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc_ben.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc_ben_Click(Area As Integer)
   dtc_codigo4.BoundText = dtc_desc_ben.BoundText
End Sub

Private Sub dtc_desc_ben_LostFocus()
    Txt_descripcion.Text = "DESEMBOLSO DE " + dtc_desc3.Text + ", PARA: " + dtc_desc_ben.Text
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'End Sub

'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
'End Sub

'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_codigo9_LostFocus()
''  If VAR_SW = "ADD" Then
''    Set rs_aux2 = New ADODB.Recordset
''    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
''    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux2.RecordCount > 0 Then
''        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''        txt_campo1.Caption = rs_aux2!correl_doc
''        rs_aux2.Update
''    End If
''  End If
'  txt_aux9.Text = dtc_desc9.Text
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

'Private Sub pnivel5(codigo5 As String)
'   'Dim strConsultaF As String
'   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub pnivel1(codigo1 As String)
''   Dim strConsultaF As String
''   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'
'   Set dtc_codigo10.RowSource = Nothing
''   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo10.ReFill
'   dtc_codigo10.BoundText = Empty
'
'   Set dtc_desc10.RowSource = Nothing
'   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc10.ReFill
'   dtc_desc10.BoundText = Empty
End Sub

'Private Sub pnivel11(codigo1 As String)
'   Dim strConsultaF As String
'   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
'End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
    'Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text + " Cite: " + Txt_campo2.Caption
End Sub

Private Sub dtc_desc11_LostFocus()
If VAR_SW = "ADD" Then
    'Txt_descripcion.Text = lbl_titulo + " - Para: " + dtc_desc3.Text
    Txt_descripcion.Text = "DESEMBOLSO DE " + dtc_desc3.Text + ", PARA: " + dtc_desc_ben.Text
Else
    If Txt_descripcion.Text = "" Then
    
    'If (dtc_codigo3.Text = "20101-4") Then
    '    Txt_descripcion.Text = lbl_titulo + ". Cite Tramite: " + Txt_campo2      'dtc_desc3.Text
    'Else
    '    Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text
    'End If
     Txt_descripcion.Text = "DESEMBOLSO DE " + dtc_desc3.Text + ", PARA: " + dtc_desc_ben.Text
        'Txt_descripcion.Text = lbl_titulo + " - Para: " + dtc_desc3.Text
'        Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
    End If
End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc3_LostFocus()
'    dtc_codigo4.Text = dtc_aux3.Text
'    Txt_descripcion.Text = lbl_titulo + " - " + dtc_desc3.Text
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
''    Call pnivel6(dtc_codigo6.BoundText)
''    dtc_desc7.Enabled = True
'End Sub

'Private Sub pnivel6(codigo6 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_proceso_nivel3 where subproceso_codigo = '" & codigo6 & "'"
'
'   Set dtc_codigo7.RowSource = Nothing
'   Set dtc_codigo7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute("EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_codigo7.ReFill
'   dtc_codigo7.BoundText = Empty
'
'   Set dtc_desc7.RowSource = Nothing
'   Set dtc_desc7.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo7.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel3 '" & codigo6 & "' ")
'   dtc_desc7.ReFill
'   dtc_desc7.BoundText = Empty
'End Sub

'Private Sub dtc_desc7_Click(Area As Integer)
'    dtc_codigo7.BoundText = dtc_desc7.BoundText
'End Sub

'Private Sub dtc_desc8_Click(Area As Integer)
'    dtc_codigo8.BoundText = dtc_desc8.BoundText
'    Call pnivel8(dtc_codigo8.BoundText)
'    'dtc_desc9.Enabled = True
'    dtc_codigo9.Enabled = True
'End Sub

'Private Sub pnivel8(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
'
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

'Private Sub dtc_desc9_Click(Area As Integer)
'    dtc_codigo9.BoundText = dtc_codigo9.BoundText
'End Sub

Private Sub Form_Load()
On Error GoTo UpdateErr
    swnuevo = 0
    VAR_SW = ""
    
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_BENEF = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_BENEF = "0"
        VAR_DA = "1.3"
    End If
    VAR_FONDO = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            VAR_DPTO = "3"
            VAR_UORIGEN = "DADMB"
            VAR_TIPO_ALM = "A"
            
        Case "1.7"    'Santa Cruz
            VAR_DPTO = "7"
            VAR_UORIGEN = "DADMS"
            VAR_TIPO_ALM = "A"
            
        Case "1.4", "1.3", "1.2"    'La Paz
            VAR_DPTO = "2"
            VAR_UORIGEN = "DCONT"
            VAR_TIPO_ALM = "A"
            
        Case "1.9"    ' Chuquisaca
            VAR_DPTO = "1"
            VAR_UORIGEN = "DADMC"
            VAR_TIPO_ALM = "A"
            
        Case Else    ' TODO
            VAR_DPTO = "2"
            VAR_UORIGEN = "DCONT"
            VAR_TIPO_ALM = "A"
     End Select
    VAR_DPTO_AUX = VAR_DPTO
    parametro = VAR_UORIGEN
    If parametro = "DCONT" Then
        BtnImprimir3.Visible = False
        BtnAprobar1.Visible = False
'        BtnAprobar3.Visible = False

        DTPfecha1.Enabled = False
        'dtc_desc3.Locked = True
        'dtc_desc3.backColor = &HC0C0C0
        Text1.Visible = True
        'dtc_desc_ben.Locked = True
        'dtc_desc_ben.backColor = &HC0C0C0
        Text5.Visible = True
'        BtnA?adir.Visible = False
        BtnAprobar.Visible = True
    Else
        BtnImprimir3.Visible = True
        BtnAprobar1.Visible = True
'        BtnAprobar3.Visible = True
        BtnAnlDetalle1.Visible = True
        DTPfecha1.Enabled = True
        BtnAddDetalle1.Visible = True
        'dtc_desc3.Locked = False
        'dtc_desc3.backColor = &HFFFFFF
        Text1.Visible = False
        'dtc_desc_ben.Locked = False
        'dtc_desc_ben.backColor = &HFFFFFF
        Text5.Visible = False
'        BtnA?adir.Visible = True
        BtnAprobar.Visible = False
    End If

    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'Call ABRIR_TABLA_DET
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    'lbl_titulo2.Caption = lbl_titulo.Caption
    'If Glaux = "PROVI" Then
        FraDet1.Caption = "DETALLE DE BIENES"
    'Else
    '    FraDet1.Caption = "EQUIPOS A IMPORTAR"
    'End If
    If rs_datos.RecordCount > 0 Then
        rs_datos.MoveFirst
    End If
    Set rsNada = New ADODB.Recordset
    If rsNada.State = 1 Then rsNada.Close
    'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rsNada.Open "select * from gc_beneficiario where beneficiario_codigo =  'XX-1000000'", db, adOpenKeyset, adLockOptimistic, adCmdText
    sino = rsNada.RecordCount
Exit Sub
UpdateErr:
  MsgBox Err.Description
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "SELECT * FROM gc_unidad_ejecutora where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'gc_tipo_solicitud
'    Set rs_datos11 = New ADODB.Recordset
'    If rs_datos11.State = 1 Then rs_datos11.Close
'    rs_datos11.Open "Select * from ac_tipo_compra_venta", db, adOpenStatic
'    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'
    'ac_tipo_compra_venta
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from ac_tipo_compra_venta ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
   
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from fo_proyectos_ejecucion order by pro_codigo_det_descripcion", db, adOpenStatic
    rs_datos3.Open "select * from av_edif_responsable WHERE edif_codigo_corto = '" & Aux & "' ORDER BY edif_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
   ' sino = Ado_datos3.Recordset.RecordCount
     
     
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "select * from gc_beneficiario where estado_codigo <> 'ANL' AND tipoben_codigo = 1 ORDER BY beneficiario_denominacion", db, adOpenStatic
    'rs_datos4.Sort = "beneficiario_denominacion"
    Set Ado_datos4.Recordset = rs_datos4
    
    dtc_codigo4.BoundText = dtc_desc_ben.BoundText
 '  sino = Ado_datos4.Recordset.RecordCount
'    Select Case Glaux
'        Case "PROVI"    'PROVISION DE EQUIPOS
'            rs_datos2.Open "Select * from gc_tipo_solicitud where solicitud_tipo = '3' order by solicitud_tipo", db, adOpenStatic
'        Case "TRANS"    'TRANSPORTE
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'TRANSP' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Case "ADUAN"    'DESADUANIZACION
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'ADUANA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'        Case "DESCA"    'DESCARGUIO Y OTROS
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "  and par_codigo = '22300' and  bien_codigo_anterior = 'DESCAR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    End Select

    'pc_poa_actividad
    
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_empresas ", db, adOpenStatic
    'rs_datos10.Open "select * from pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

'Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepci?n as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
'    queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub ABRIR_TABLA_DET()
'    Set rsNada = New ADODB.Recordset
'    If rsNada.State = 1 Then rsNada.Close
'    'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    rsNada.Open "select * from gc_beneficiario where beneficiario_codigo =  'XX-1000000'", db, adOpenKeyset, adLockOptimistic, adCmdText
'    sino = rsNada.RecordCount
    sino = "0"

    'BIENES (Equipos a Comprar)
'    Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
'            rs_det1.Sort = "compra_codigo_det"
'            Set Ado_detalle1.Recordset = rs_det1
'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
'                dg_det1.Visible = True
'                BtnAddDetalle3.Visible = True
'                Set dg_det1.DataSource = Ado_detalle1.Recordset
'            Else
'                dg_det1.Visible = False
'                dg_det1A.Visible = False
'                Command1.Visible = False
'                BtnAddDetalle3.Visible = False
'                 'Set Ado_detalle1.Recordset = rsNada
'                Set dg_det1.DataSource = rsNada
'                 'dg_det1.ClearFields
'
'            End If

'            If Ado_detalle1.Recordset.RecordCount > 0 Then
'            BtnAddDetalle1.Visible = False
'            Else
'            BtnAddDetalle1.Visible = True
'            End If

 If parametro = "COMEX" Then
    Select Case Glaux
        Case "PROVI"    'PROVISION DE EQUIPOS
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_det1.Open "select * from av_compra_detalle_tipo where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            Set Ado_detalle1.Recordset = rs_det1
            Ado_detalle1.Recordset.Sort = "bien_descripcion"
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                Ado_detalle1.Recordset.MoveFirst
                Set dg_det1.DataSource = Ado_detalle1.Recordset
            Else
                dg_det1.Visible = False
                Set dg_det1.DataSource = rsNada
            End If
                 
            
        Case "TRANS"    'TRANSPORTE
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND bien_codigo_anterior = 'TRANS' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " AND bien_codigo_anterior = 'TRANS' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            Set Ado_detalle1.Recordset = rs_det1
            Set dg_det1.DataSource = Ado_detalle1.Recordset
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                Set dg_det1.DataSource = Ado_detalle1.Recordset
            Else
                dg_det1.Visible = False
                Set dg_det1.DataSource = rsNada
            End If
        Case "ADUAN"    'DESADUANIZACION
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and  bien_codigo_anterior = 'ADUAN'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and  bien_codigo_anterior = 'ADUAN'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            Set Ado_detalle1.Recordset = rs_det1
            Set dg_det1.DataSource = Ado_detalle1.Recordset
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                Set dg_det1.DataSource = Ado_detalle1.Recordset
            Else
                dg_det1.Visible = False
                Set dg_det1.DataSource = rsNada
            End If
        Case "DESCA"    'DESCARGUIO Y OTROS
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and  bien_codigo_anterior = 'DESCA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and  bien_codigo_anterior = 'DESCA'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            Set Ado_detalle1.Recordset = rs_det1
            Set dg_det1.DataSource = Ado_detalle1.Recordset
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                Set dg_det1.DataSource = Ado_detalle1.Recordset
            Else
                dg_det1.Visible = False
                Set dg_det1.DataSource = rsNada
            End If
        Case Else
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
             Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_det1.Sort = "compra_codigo_det"
            rs_det1.Sort = "bien_descripcion"
            Set Ado_detalle1.Recordset = rs_det1
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                'BtnAddDetalle3.Visible = True
                Set dg_det1.DataSource = Ado_detalle1.Recordset
            Else
                dg_det1.Visible = False
                'dg_det1A.Visible = False
                Command1.Visible = False
                'BtnAddDetalle3.Visible = False
                 'Set Ado_detalle1.Recordset = rsNada
                Set dg_det1.DataSource = rsNada
                 'dg_det1.ClearFields
               
            End If
    End Select
  Else
        Set rs_aux11 = New ADODB.Recordset
        If rs_aux11.State = 1 Then rs_aux11.Close
        Set rs_det1 = New ADODB.Recordset
            If rs_det1.State = 1 Then rs_det1.Close
            'rs_det1.Open "select * from ao_compra_detalle where compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1.Open "select * from av_compra_detalle where compra_codigo = " & VAR_COMPRA & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux11.Open "select * from av_compra_detalle where compra_codigo = " & VAR_COMPRA & "", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_det1.Sort = "bien_descripcion"   '"compra_codigo_det"
            
            Set Ado_detalle1.Recordset = rs_det1
            If Ado_detalle1.Recordset.RecordCount > 0 Then
                'VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
                dg_det1.Visible = True
                'BtnAddDetalle3.Visible = True
                Set dg_det1.DataSource = Ado_detalle1.Recordset
                'Set rs_aux11 = rs_det1
            Else
                dg_det1.Visible = False
               ' dg_det1A.Visible = False
                'Command1.Visible = False
                'BtnAddDetalle3.Visible = False
                 'Set Ado_detalle1.Recordset = rsNada
                Set dg_det1.DataSource = rsNada
                 'dg_det1.ClearFields
            End If
            
 End If
 SUMbs = 0
 SUMdol = 0
 If Ado_detalle1.Recordset.RecordCount > 0 Then
 'If rs_aux11.RecordCount > 0 Then
    'rs_det1.MoveFirst
    rs_aux11.MoveFirst
    While Not rs_aux11.EOF
    'While Not rs_det1.EOF
        'SUMbs = Format(Round(SUMbs + IIf(IsNull(rs_det1("compra_precio_total_bs")), 0, rs_det1("compra_precio_total_bs")), 2), "###,###,##0.00")
        'SUMdol = Format(Round(SUMdol + IIf(IsNull(rs_det1("compra_precio_total_dol")), 0, rs_det1("compra_precio_total_dol")), 2), "###,###,##0.00")
        SUMbs = Format(Round(SUMbs + IIf(IsNull(rs_aux11("compra_precio_total_bs")), 0, rs_aux11("compra_precio_total_bs")), 2), "###,###,##0.00")
        SUMdol = Format(Round(SUMdol + IIf(IsNull(rs_aux11("compra_precio_total_dol")), 0, rs_aux11("compra_precio_total_dol")), 2), "###,###,##0.00")
        'rs_det1.MoveNext
        rs_aux11.MoveNext
    Wend
    'JQA
    'Ado_detalle1.Recordset.MoveFirst
    lbl_total_bs.Text = IIf(IsNull(SUMbs), 0, SUMbs)
    lbl_total_dol.Text = IIf(IsNull(SUMdol), 0, SUMdol)
 End If
 '    'rs_det1.Open "select * from ao_compra_detalle where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
 If Ado_detalle1.Recordset.RecordCount > 0 Then
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    rs_det2.Open "select * from ao_compra_adjudica where compra_codigo = " & VAR_COMPRA & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and
    '"select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'AND compra_codigo_det = " & Ado_detalle1.Recordset!compra_codigo_det & "
    rs_det2.Sort = "adjudica_codigo"
    Set Ado_detalle2.Recordset = rs_det2
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        dg_det2.Visible = True
        dg_det1A.Visible = True
'        FrmABMDet2.Visible = True
        Ado_detalle2.Recordset.MoveFirst
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det2.Visible = False
        dg_det1A.Visible = False
        Set dg_det2.DataSource = rsNada
        ' dg_det1A.Visible = False
        'Set Ado_detalle2.Recordset = rsNada
    End If
 Else
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    rs_det2.Open "select * from ao_compra_adjudica where unidad_codigo = '" & "025 " & "' and solicitud_codigo = " & "025" & "", db, adOpenKeyset, adLockOptimistic, adCmdText '"select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det2.Sort = "adjudica_codigo"
    Set Ado_detalle2.Recordset = rs_det2
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        dg_det2.Visible = True
        dg_det1A.Visible = True
'        FrmABMDet2.Visible = True
        Ado_detalle2.Recordset.MoveFirst
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det2.Visible = False
        dg_det1A.Visible = False
        Set dg_det2.DataSource = rsNada
        'Set Ado_detalle2.Recordset = rsNada
    End If


        'dg_det2.ClearFields
 End If
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'If parametro <> "COMEX" Then
    BtnAprobar1.Visible = True
'    BtnAprobar3.Visible = True
'End If
  'Esto mostrar? la posici?n de registro actual para este Recordset
  If Ado_datos.Recordset.BOF = False Then
    If Ado_datos.Recordset.RecordCount > 0 Then
        VAR_COMPRA = Ado_datos.Recordset!compra_codigo
      If IsNull(DTPfecha1.Value) = True Then
        DTPfecha1.Value = Ado_datos.Recordset!compra_fecha
      End If
    End If
    
  'dtc_codigo11.Text = rs_datos!beneficiario_codigo_resp
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificaci?n del Cliente                Fin -->   'esto es de Caption
    If VAR_SW <> "ADD" Then         'Or VAR_SW <> "MOD"
        VAR_COD4 = parametro
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        Call ABRIR_TABLA_DET
        Call ABRIR_TABLA_AUX2
    Else
'        If VAR_SW <> "MOD" Then
'            VAR_COD4 = parametro
'            VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'            Call ABRIR_TABLA_DET
'            Call ABRIR_TABLA_AUX2
'        'Set rs_det1 = New ADODB.Recordset
'        Else
            Set dg_det2.DataSource = rsNada
'        End If
       ' Set Ado_detalle2.Recordset = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    'FraDet1.Caption = "BIT?CORA DE: " + dtc_desc1.Text
'    txt_aux9.Text = dtc_desc9.Text
'    If Ado_datos.Recordset!estado_codigo_eqp = "APR" Then
'            'FrmABMDet2.Visible = False
'    Else
'            'FrmABMDet2.Visible = True
'    End If
  Else
'  If Ado_datos.Recordset.RecordCount > 0 Then
    Set dg_det1.DataSource = rsNada
    'Set Ado_detalle1.Recordset = rsNada
    Set dg_det2.DataSource = rsNada
    'Set Ado_detalle2.Recordset = rsNada
  End If
End Sub

Private Sub Ado_detalle3_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    VAR_COD4 = parametro
    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aqu? se coloca el c?digo de validaci?n
  'Se llama a este evento cuando ocurre la siguiente acci?n
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub BtnA?adir_Click()
'VAR_SW = "NEW"
    Fra_datos.Visible = True
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    FraDet2.Visible = False
    FraDet1.Visible = False
  On Error GoTo AddErr
    VAR_SW = "ADD"
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
If Ado_datos.Recordset.RecordCount > 0 Then
     Ado_datos.Recordset.MoveLast
 End If
    Ado_datos.Recordset.AddNew
    dtc_desc11.SetFocus
    DTPfecha1.Value = Date
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_desc2.Locked = True
    Text1.Visible = False
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
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'dtc_desc3.BoundText = "20101-4"
    'dtc_codigo3.Text = "20101-4"
    
'    dtc_codigo5.Text = "COM"
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_codigo6.Text = "COM-01"
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'    dtc_codigo7.Text = "COM-01-02"
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    BtnVer.Visible = False
'    dtc_codigo9.Enabled = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto s?lo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
'    rs.Open GlSqlAux, db, adOpenStatic
'    ExisteReg = rs!Cuantos > 0
End Function

Private Sub lbl_total_bs_LostFocus()
    If lbl_total_bs.Text = "" Or lbl_total_bs.Text = "0" Then
        lbl_total_bs.Text = "1"
    Else
        lbl_total_dol.Text = CDbl(lbl_total_bs.Text) / 6.96
    End If
End Sub

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'If parametro = "COMEX" Then
        Select Case Aux
            Case "30"   'Aux = "30" 'CCD
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp = 'REG'  AND solicitud_tipo = '" & Aux & "' "                   'AND unidad_codigo_adm = '" & parametro & "'
            Case "29"   'Aux = "29" 'VIAJES
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp = 'REG'  AND solicitud_tipo = '" & Aux & "' "
            Case "28"   'Aux = "28" 'CAJA CHICA
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp = 'REG'  AND solicitud_tipo = '" & Aux & "' "
             
            Case Else
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp = 'REG'  AND solicitud_tipo = '" & Aux & "' "
        End Select
    'Else
    '    queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp = 'REG' AND unidad_codigo_adm = '" & parametro & "' "
    'End If
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_numero_alm"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'If parametro = "COMEX" Then
        Select Case Aux
            Case "30"   'Aux = "30" 'CCD
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp <> 'ANL'  AND solicitud_tipo = '" & Aux & "' "                   'AND unidad_codigo_adm = '" & parametro & "'
            Case "29"   'Aux = "29" 'VIAJES
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp <> 'ANL'  AND solicitud_tipo = '" & Aux & "' "
            Case "28"   'Aux = "28" 'CAJA CHICA
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp <> 'ANL'  AND solicitud_tipo = '" & Aux & "' "
            Case Else
                queryinicial = "Select * from ao_compra_cabecera where estado_codigo_eqp <> 'ANL'  AND solicitud_tipo = '" & Aux & "' "
        End Select
    'Else
    '    queryinicial = "Select * from ao_compra_cabecera where  unidad_codigo_adm = '" & parametro & "' "
    'End If
    
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_numero_alm"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_ao_pagos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Administrativos - COMEX - Proceso de Pagos"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   15270
   Icon            =   "Frm_ao_pagos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.2232e6
   ScaleMode       =   0  'User
   ScaleWidth      =   1.00607e7
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Picture         =   "Frm_ao_pagos.frx":0A02
      ScaleHeight     =   1755
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   184
      Top             =   8160
      Width           =   1935
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ver..."
         Height          =   645
         Left            =   600
         Picture         =   "Frm_ao_pagos.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Ver Detalle del Bien ..."
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1515
      Left            =   120
      Picture         =   "Frm_ao_pagos.frx":6CFEC
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   180
      Top             =   6525
      Width           =   1935
      Begin VB.CommandButton BtnImprimir4 
         BackColor       =   &H80000018&
         Caption         =   "Cronogr."
         Height          =   640
         Left            =   120
         Picture         =   "Frm_ao_pagos.frx":D901E
         Style           =   1  'Graphical
         TabIndex        =   183
         ToolTipText     =   "Imprime Cronograma de Cobranzas ..."
         Top             =   750
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Ver..."
         Height          =   645
         Left            =   600
         Picture         =   "Frm_ao_pagos.frx":DA7A0
         Style           =   1  'Graphical
         TabIndex        =   182
         ToolTipText     =   "Ver Datos de la Venta ..."
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H80000018&
         Caption         =   "O. Pago"
         Height          =   645
         Left            =   960
         Picture         =   "Frm_ao_pagos.frx":DAD58
         Style           =   1  'Graphical
         TabIndex        =   181
         ToolTipText     =   "Imprime Orden de Pago"
         Top             =   750
         Width           =   765
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "Frm_ao_pagos.frx":DB315
      ScaleHeight     =   960
      ScaleWidth      =   14940
      TabIndex        =   169
      Top             =   120
      Width           =   15000
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "Frm_ao_pagos.frx":147347
         Style           =   1  'Graphical
         TabIndex        =   178
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "Frm_ao_pagos.frx":14796B
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "Frm_ao_pagos.frx":147F4B
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "Frm_ao_pagos.frx":148C15
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "Frm_ao_pagos.frx":148E1F
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "Frm_ao_pagos.frx":1493D7
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   360
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "Frm_ao_pagos.frx":1495E1
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "Frm_ao_pagos.frx":149A23
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00808000&
         Caption         =   "Kardex"
         Height          =   720
         Left            =   4320
         Picture         =   "Frm_ao_pagos.frx":149C2D
         Style           =   1  'Graphical
         TabIndex        =   170
         ToolTipText     =   "Imprime Recibo / Kardex"
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
         TabIndex        =   179
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "Frm_ao_pagos.frx":14B3AF
      ScaleHeight     =   915
      ScaleWidth      =   14940
      TabIndex        =   165
      Top             =   120
      Width           =   15000
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "Frm_ao_pagos.frx":1B73E1
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "Frm_ao_pagos.frx":1B75EB
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Cancelar"
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
         TabIndex        =   168
         Top             =   300
         Width           =   1305
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   5880
      TabIndex        =   9
      Top             =   1245
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9049
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "REGISTRO DE PAGOS"
      TabPicture(0)   =   "Frm_ao_pagos.frx":1B77F5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCobros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "BtnImprimir3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "DATOS DE LA COMPRA"
      TabPicture(1)   =   "Frm_ao_pagos.frx":1B7811
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmCabecera"
      Tab(1).Control(1)=   "BtnSalir2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DETALLE BIENES Y SERVICIOS"
      TabPicture(2)   =   "Frm_ao_pagos.frx":1B782D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmEdita"
      Tab(2).Control(1)=   "BtnSalir3"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton BtnImprimir3 
         BackColor       =   &H00C0C000&
         Caption         =   "Factura"
         Height          =   640
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Imprime Factura"
         Top             =   2880
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Volver..."
         Height          =   645
         Left            =   -66720
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Ver Datos de la Venta ..."
         Top             =   4200
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir2 
         BackColor       =   &H80000018&
         Caption         =   "Volver..."
         Height          =   645
         Left            =   -66720
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Ver Datos de la Venta ..."
         Top             =   3960
         Width           =   765
      End
      Begin VB.Frame FrmEdita 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "E"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4680
         Left            =   -74950
         TabIndex        =   109
         Top             =   360
         Width           =   9135
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   8430
            TabIndex        =   128
            Top             =   370
            Width           =   255
         End
         Begin VB.TextBox Txt_modelo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "modelo_codigo"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   2040
            TabIndex        =   127
            Text            =   "0"
            Top             =   1860
            Width           =   1815
         End
         Begin VB.TextBox Txt_modelo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "modelo_codigo_x"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   4680
            TabIndex        =   126
            Text            =   "0"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Txt_modelo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "modelo_codigo_h"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   125
            Text            =   "0"
            Top             =   2100
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Txt_modelo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "modelo_codigo1"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   360
            TabIndex        =   124
            Text            =   "0"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.OptionButton OpMod3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3"
            Height          =   285
            Left            =   6600
            TabIndex        =   123
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton OpMod2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            Height          =   285
            Left            =   4320
            TabIndex        =   122
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton OpMod1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   285
            Left            =   2055
            TabIndex        =   121
            Top             =   2100
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   6420
            TabIndex        =   120
            Top             =   1210
            Width           =   280
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   3540
            TabIndex        =   119
            Top             =   2540
            Width           =   245
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   8430
            TabIndex        =   118
            Top             =   1210
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   290
            Left            =   8430
            TabIndex        =   117
            Top             =   2535
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   8430
            TabIndex        =   116
            Top             =   1875
            Width           =   255
         End
         Begin VB.TextBox txt_descripcion_venta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            CausesValidation=   0   'False
            DataField       =   "concepto_venta"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   345
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Top             =   3975
            Width           =   7665
         End
         Begin VB.TextBox TxtNroVenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "venta_codigo"
            DataSource      =   "ado_datos14"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   114
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox TxtCantidad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "venta_det_cantidad"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   360
            TabIndex        =   113
            Text            =   "0"
            Top             =   3255
            Width           =   975
         End
         Begin VB.TextBox TxtDescuento 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "venta_descuento_bs"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3840
            TabIndex        =   112
            Text            =   "0"
            Top             =   3255
            Width           =   1455
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "venta_precio_total_bs"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   111
            Text            =   "0"
            Top             =   3255
            Width           =   1575
         End
         Begin VB.TextBox TxtPrecioU 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            DataField       =   "venta_precio_unitario_bs"
            DataSource      =   "ado_datos14"
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
            Height          =   285
            Left            =   1800
            TabIndex        =   110
            Text            =   "0"
            Top             =   3255
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dtc_preciocompra15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3600
            TabIndex        =   129
            Top             =   2760
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "bien_precio_compra"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_subgrupo15 
            CausesValidation=   0   'False
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3720
            TabIndex        =   130
            Top             =   1560
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "subgrupo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dtc_grupo15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5280
            TabIndex        =   131
            Top             =   1560
            Visible         =   0   'False
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "grupo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dtc_precioventafinal15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6045
            TabIndex        =   132
            Top             =   2760
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "bien_precio_venta_final"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6760
            TabIndex        =   133
            Top             =   1200
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   360
            TabIndex        =   134
            Top             =   1200
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc12 
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3600
            TabIndex        =   135
            Top             =   120
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "tipoben_descripcion"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_aux12 
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5280
            TabIndex        =   136
            Top             =   120
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "tipoben_descuento"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5260
            TabIndex        =   137
            Top             =   360
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_unimed15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7300
            TabIndex        =   138
            Top             =   1860
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "unimed_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dtc_stocktotal15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7320
            TabIndex        =   139
            Top             =   2520
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "bien_stock_actual"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dtc_codigo12 
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   2880
            TabIndex        =   140
            Top             =   120
            Visible         =   0   'False
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "tipoben_codigo"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7800
            TabIndex        =   141
            Top             =   120
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "almacen_codigo"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_Stock13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   2520
            TabIndex        =   142
            Top             =   2520
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "stock_actual"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_partida15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   2040
            TabIndex        =   143
            Top             =   1560
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "par_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_precioventabase15 
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   1080
            TabIndex        =   144
            Top             =   2760
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "bien_precio_venta_base"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Almacen de Origen:"
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
            TabIndex        =   160
            Top             =   375
            Width           =   1770
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Total Actual"
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
            Left            =   4980
            TabIndex        =   159
            Top             =   2520
            Width           =   1635
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad de Medida:"
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
            Left            =   5400
            TabIndex        =   158
            Top             =   1920
            Width           =   1725
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Pago:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   240
            Left            =   360
            TabIndex        =   157
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción y Características Complementarias"
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
            Left            =   360
            TabIndex        =   156
            Top             =   3675
            Width           =   4245
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Height          =   240
            Left            =   360
            TabIndex        =   155
            Top             =   3000
            Width           =   810
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Bien"
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
            Left            =   360
            TabIndex        =   154
            Top             =   930
            Width           =   1860
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Bien"
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
            Left            =   6720
            TabIndex        =   153
            Top             =   930
            Width           =   1110
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total"
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
            Left            =   6000
            TabIndex        =   152
            Top             =   3000
            Width           =   1305
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   5400
            TabIndex        =   151
            Top             =   3240
            Width           =   285
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1440
            TabIndex        =   150
            Top             =   3285
            Width           =   240
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Unitario"
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
            Left            =   1800
            TabIndex        =   149
            Top             =   3000
            Width           =   1560
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3480
            TabIndex        =   148
            Top             =   3240
            Width           =   225
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
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
            Left            =   3915
            TabIndex        =   147
            Top             =   3000
            Width           =   1350
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo del Bien:"
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
            Left            =   360
            TabIndex        =   146
            Top             =   1870
            Width           =   1515
         End
         Begin VB.Label Label23 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Almacen Origen"
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
            Height          =   315
            Left            =   360
            TabIndex        =   145
            Top             =   2520
            Width           =   2145
         End
      End
      Begin VB.Frame FrmCabecera 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4710
         Left            =   -74950
         TabIndex        =   57
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   320
            Width           =   1125
         End
         Begin VB.Frame Fra_Total 
            BackColor       =   &H00000000&
            Caption         =   "----------------------------------------------- Datos Totalizados del Tramite en Bs."
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
            Height          =   975
            Left            =   60
            TabIndex        =   78
            Top             =   3420
            Width           =   7935
            Begin VB.TextBox TxtBstotal 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               DataField       =   "venta_saldo_p_cobrar_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos16"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FFFF&
               Height          =   285
               Left            =   6045
               TabIndex        =   83
               Text            =   "0"
               Top             =   580
               Width           =   1545
            End
            Begin VB.TextBox TxtMontoBs 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               DataField       =   "venta_monto_total_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos16"
               ForeColor       =   &H0080FFFF&
               Height          =   285
               Left            =   1920
               TabIndex        =   82
               Text            =   "0"
               Top             =   580
               Width           =   1545
            End
            Begin VB.TextBox txtCantTotal 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               DataField       =   "venta_cantidad_total"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos16"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FFFF&
               Height          =   285
               Left            =   480
               TabIndex        =   81
               Text            =   "0"
               Top             =   580
               Width           =   855
            End
            Begin VB.TextBox TxtCobrado 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               DataField       =   "venta_monto_cobrado_bs"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "Ado_datos16"
               ForeColor       =   &H0080FFFF&
               Height          =   285
               Left            =   3960
               TabIndex        =   80
               Text            =   "0"
               Top             =   580
               Width           =   1545
            End
            Begin VB.TextBox txtTDC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               DataField       =   "venta_tipo_cambio"
               DataSource      =   "Ado_datos16"
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   8160
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   79
               Top             =   180
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   285
               Left            =   5565
               TabIndex        =   89
               Top             =   585
               Width           =   405
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFF80&
               Height          =   285
               Left            =   3495
               TabIndex        =   88
               Top             =   585
               Width           =   405
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Total Contrato"
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
               Height          =   285
               Left            =   1920
               TabIndex        =   87
               Top             =   285
               Width           =   1335
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo p/Pagar"
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
               Height          =   285
               Left            =   6060
               TabIndex        =   86
               Top             =   285
               Width           =   1380
            End
            Begin VB.Label Label21 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Cantidad Total -->"
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
               Height          =   285
               Left            =   240
               TabIndex        =   85
               Top             =   285
               Width           =   1335
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFF80&
               X1              =   1755
               X2              =   1755
               Y1              =   1080
               Y2              =   120
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Total Pagado"
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
               Height          =   285
               Left            =   3915
               TabIndex        =   84
               Top             =   285
               Width           =   1455
            End
         End
         Begin VB.Frame Fra_datos 
            BackColor       =   &H00000000&
            Caption         =   "-- Fecha de Pago ------- Tipo de Tramite----------------------------- Plazo días Calend.----- Código Registro"
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
            Height          =   1755
            Left            =   60
            TabIndex        =   62
            Top             =   1640
            Width           =   9015
            Begin VB.TextBox TxtConcepto 
               Appearance      =   0  'Flat
               BackColor       =   &H80000013&
               DataField       =   "venta_descripcion"
               DataSource      =   "Ado_datos16"
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   66
               Top             =   1200
               Width           =   5655
            End
            Begin VB.TextBox TxtPlazo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000013&
               DataField       =   "venta_plazo_dias_calendario"
               DataSource      =   "Ado_datos16"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   5520
               TabIndex        =   65
               Text            =   "0"
               Top             =   270
               Width           =   1335
            End
            Begin VB.TextBox Text5 
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   4920
               TabIndex        =   64
               Top             =   285
               Width           =   320
            End
            Begin VB.TextBox Text14 
               BackColor       =   &H80000013&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   280
               Left            =   6570
               TabIndex        =   63
               Top             =   730
               Width           =   270
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos16"
               Height          =   315
               Left            =   2160
               TabIndex        =   67
               Top             =   720
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483629
               ForeColor       =   16777215
               ListField       =   "descripcion"
               BoundColumn     =   "codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos16"
               Height          =   315
               Left            =   2160
               TabIndex        =   68
               Top             =   270
               Width           =   3090
               _ExtentX        =   5450
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483629
               ForeColor       =   16777215
               ListField       =   "venta_tipo_descripcion"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos16"
               Height          =   315
               Left            =   4560
               TabIndex        =   69
               Top             =   135
               Visible         =   0   'False
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "venta_tipo"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2160
               TabIndex        =   70
               Top             =   720
               Visible         =   0   'False
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "codigo"
               BoundColumn     =   "codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_aux4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos16"
               Height          =   315
               Left            =   5760
               TabIndex        =   71
               Top             =   480
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "codigo2"
               BoundColumn     =   "codigo"
               Text            =   "DataCombo1"
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
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
               Left            =   240
               TabIndex        =   77
               Top             =   1185
               Width           =   1020
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Responsable Pago :"
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
               TabIndex        =   76
               Top             =   735
               Width           =   1860
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Nro. Documento"
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
               Index           =   13
               Left            =   7440
               TabIndex        =   75
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label txt_campo1 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               Caption         =   "0"
               DataField       =   "doc_numero"
               DataSource      =   "Ado_datos16"
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
               Left            =   7440
               TabIndex        =   74
               Top             =   1200
               Width           =   1365
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFF80&
               X1              =   7080
               X2              =   7080
               Y1              =   120
               Y2              =   1815
            End
            Begin VB.Label txt_codigo1 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               Caption         =   "0"
               DataField       =   "doc_codigo"
               DataSource      =   "Ado_datos16"
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
               Left            =   7320
               TabIndex        =   73
               Top             =   360
               Width           =   1605
            End
            Begin VB.Label DTPfechasol 
               Alignment       =   2  'Center
               BackColor       =   &H80000013&
               Caption         =   "36NO"
               DataField       =   "venta_fecha"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MMM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               DataSource      =   "Ado_datos16"
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
               TabIndex        =   72
               Top             =   270
               Width           =   1600
            End
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   61
            Top             =   1230
            Width           =   330
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   60
            Top             =   790
            Width           =   330
         End
         Begin VB.TextBox Text12 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   320
            Left            =   9045
            TabIndex        =   59
            Top             =   780
            Width           =   375
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   58
            Top             =   365
            Width           =   330
         End
         Begin MSDataListLib.DataCombo Dtc_deudor2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   8520
            TabIndex        =   91
            Top             =   780
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   255
            ForeColor       =   0
            ListField       =   "beneficiario_deudor"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   5580
            TabIndex        =   92
            Top             =   780
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1200
            TabIndex        =   93
            Top             =   780
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   5460
            TabIndex        =   94
            Top             =   120
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
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   3315
            TabIndex        =   95
            Top             =   345
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Dtc_aux2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   4320
            TabIndex        =   96
            Top             =   600
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   -2147483624
            ListField       =   "codigo2"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   7260
            TabIndex        =   97
            Top             =   1200
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo5"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5580
            TabIndex        =   98
            Top             =   1215
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   1200
            TabIndex        =   99
            Top             =   1215
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
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
            Left            =   180
            TabIndex        =   108
            Top             =   795
            Width           =   1005
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   285
            Left            =   180
            TabIndex        =   107
            Top             =   80
            Width           =   1125
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cite de Tramite"
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
            Index           =   6
            Left            =   7605
            TabIndex        =   106
            Top             =   75
            Width           =   1380
         End
         Begin VB.Label lbl_campo1 
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
            Left            =   3345
            TabIndex        =   105
            Top             =   75
            Width           =   1680
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cod.Trámite"
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
            Left            =   1680
            TabIndex        =   104
            Top             =   75
            Width           =   1110
         End
         Begin VB.Label Txt_campo2 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "36NO"
            DataField       =   "unidad_codigo_ant"
            DataSource      =   "Ado_datos16"
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
            Left            =   7500
            TabIndex        =   103
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label txt_codigo 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "solicitud_codigo"
            DataSource      =   "Ado_datos16"
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
            Left            =   1660
            TabIndex        =   102
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Deudor ?"
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
            Left            =   7470
            TabIndex        =   101
            Top             =   795
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto:"
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
            Left            =   180
            TabIndex        =   100
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.Frame FrmCobros 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4710
         Left            =   60
         TabIndex        =   15
         Top             =   380
         Width           =   9135
         Begin VB.CommandButton BtnVer2 
            BackColor       =   &H00C0C000&
            Caption         =   "Nuevo"
            Height          =   640
            Left            =   7875
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Registra Nuevo Beneficiario"
            Top             =   1515
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8745
            TabIndex        =   56
            Top             =   3210
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_desc2A 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1515
            TabIndex        =   4
            Top             =   2115
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin VB.TextBox TxtObs 
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos"
            Height          =   465
            Left            =   1515
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2565
            Width           =   6555
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   1905
            TabIndex        =   0
            Text            =   "0"
            Top             =   1140
            Width           =   1455
         End
         Begin VB.TextBox TxtDscto 
            Alignment       =   2  'Center
            DataField       =   "cobranza_descuento_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   1905
            TabIndex        =   1
            Text            =   "0"
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8625
            TabIndex        =   18
            Top             =   2135
            Width           =   255
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8625
            TabIndex        =   16
            Top             =   1655
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   17
            Top             =   1635
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2A 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7080
            TabIndex        =   19
            Top             =   2115
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc4A 
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1515
            TabIndex        =   3
            Top             =   1635
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro 
            DataField       =   "cobranza_fecha_cobro"
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   7275
            TabIndex        =   2
            Top             =   1140
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   247791619
            CurrentDate     =   41678
            MaxDate         =   109939
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo dtc_cta 
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1515
            TabIndex        =   54
            Top             =   3195
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_ctades 
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4200
            TabIndex        =   55
            Top             =   3195
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   -2147483629
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.Bancaria:"
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
            Left            =   225
            TabIndex        =   53
            Top             =   3220
            Width           =   1200
         End
         Begin VB.Label TxtMontoDol 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_deuda_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3520
            TabIndex        =   52
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label TxtDsctoTot 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_programada_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   240
            TabIndex        =   51
            Top             =   1140
            Width           =   1455
         End
         Begin VB.Label TxtCobrador 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "descripcion"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1515
            TabIndex        =   50
            Top             =   1680
            Width           =   4245
         End
         Begin VB.Label DTPFechaProg 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_fecha_prog"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   5325
            TabIndex        =   49
            Top             =   1140
            Width           =   1665
         End
         Begin VB.Label TxtNroVentaC 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4920
            TabIndex        =   48
            Top             =   240
            Width           =   1245
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   9135
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   5140
            X2              =   5140
            Y1              =   720
            Y2              =   1680
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Programada"
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
            Left            =   5160
            TabIndex        =   47
            Top             =   855
            Width           =   1875
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "doc_codigo_fac"
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
            Left            =   1755
            TabIndex        =   24
            Top             =   4245
            Width           =   1245
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
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
            Left            =   7920
            TabIndex        =   46
            Top             =   240
            Width           =   980
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de cuota:"
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
            Left            =   6600
            TabIndex        =   45
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Trámite:"
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
            Left            =   3675
            TabIndex        =   44
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Pago:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   240
            Left            =   225
            TabIndex        =   40
            Top             =   255
            Width           =   1065
         End
         Begin VB.Label lbl_obs 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
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
            Left            =   225
            TabIndex        =   39
            Top             =   2620
            Width           =   1200
         End
         Begin VB.Label lbl_monto 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagado en Bs."
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
            Left            =   1905
            TabIndex        =   38
            Top             =   855
            Width           =   1470
         End
         Begin VB.Label Label46 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Programado Bs."
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
            Left            =   240
            TabIndex        =   37
            Top             =   855
            Width           =   1545
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagado en Dol."
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
            Left            =   3440
            TabIndex        =   36
            Top             =   855
            Width           =   1635
         End
         Begin VB.Label Lbl_Cobrador 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Pagador CGI:"
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
            Left            =   225
            TabIndex        =   35
            Top             =   1665
            Width           =   1215
         End
         Begin VB.Label lbl_fechas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Pago"
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
            Left            =   7290
            TabIndex        =   34
            Top             =   855
            Width           =   1515
         End
         Begin VB.Label lbl_factura 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Ch/Trf."
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
            Left            =   4035
            TabIndex        =   33
            Top             =   4260
            Width           =   975
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   1755
            TabIndex        =   32
            Top             =   1140
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3195
            TabIndex        =   31
            Top             =   1140
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "No.Documento"
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
            Left            =   7200
            TabIndex        =   30
            Top             =   3960
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lbl_doc1 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1755
            TabIndex        =   29
            Top             =   3760
            Width           =   1245
         End
         Begin VB.Label lbl_docnro 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   5085
            TabIndex        =   28
            Top             =   3760
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Nro. de Respaldo:"
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
            Index           =   1
            Left            =   3285
            TabIndex        =   27
            Top             =   3760
            Width           =   1650
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cod. Respaldo:"
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
            Index           =   2
            Left            =   225
            TabIndex        =   26
            Top             =   3760
            Width           =   1410
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Cód.Registro"
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
            Index           =   4
            Left            =   225
            TabIndex        =   25
            Top             =   4260
            Width           =   1185
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   9135
            Y1              =   3640
            Y2              =   3640
         End
         Begin VB.Label Lbl_nombre_fac 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
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
            Left            =   225
            TabIndex        =   23
            Top             =   2130
            Width           =   1230
         End
         Begin VB.Label Txt_cod_cobro 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_codigo"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1760
            TabIndex        =   22
            Top             =   240
            Width           =   1485
         End
         Begin VB.Label TxtCmpbte 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_nro_factura"
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
            Left            =   5085
            TabIndex        =   21
            Top             =   4245
            Width           =   1365
         End
         Begin VB.Label TxtAutorizacion 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "cobranza_nro_autorizacion"
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
            Left            =   7080
            TabIndex        =   20
            Top             =   4245
            Visible         =   0   'False
            Width           =   1845
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTA"
      ForeColor       =   &H00FFFFC0&
      Height          =   5160
      Left            =   135
      TabIndex        =   12
      Top             =   1200
      Width           =   5745
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   4380
         Left            =   80
         TabIndex        =   42
         Top             =   240
         Width           =   5600
         _ExtentX        =   9869
         _ExtentY        =   7726
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Fecha.Prog."
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
            DataField       =   "cobranza_codigo"
            Caption         =   "No.Pago"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Cobrador"
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
            DataField       =   "cobranza_fecha_cobro"
            Caption         =   "Fecha.Pago"
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
         BeginProperty Column04 
            DataField       =   "cobranza_total_bs"
            Caption         =   "Pagado Bs."
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
         BeginProperty Column05 
            DataField       =   "cobranza_total_dol"
            Caption         =   "Cobrado en Dol."
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.Nota.Cobro"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "Nro. Factura"
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto/Observaciones"
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
         BeginProperty Column10 
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cliente"
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
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   2445.166
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
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
         TabIndex        =   14
         Top             =   4755
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
         TabIndex        =   13
         Top             =   4755
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   80
         Top             =   4680
         Width           =   5600
         _ExtentX        =   9869
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
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00000000&
      Caption         =   "DATOS DE LA COMPRA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   1575
      Left            =   2160
      TabIndex        =   11
      Top             =   6465
      Width           =   12975
      Begin MSDataGridLib.DataGrid dg_datos16 
         Height          =   1170
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   2064
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Compra"
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
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Proveedor"
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
            DataField       =   "venta_fecha"
            Caption         =   "Fecha.Compra"
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
            DataField       =   "edif_codigo"
            Caption         =   "Cod.Proyecto"
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
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.Ejecutora"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Nro.Solicitud/Neg."
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
         BeginProperty Column07 
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3509.858
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1649.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE BIENES Y SERVICIOS"
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
      Height          =   1965
      Left            =   2160
      TabIndex        =   10
      Top             =   8085
      Width           =   12975
      Begin MSDataGridLib.DataGrid DtGLista 
         Height          =   1620
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   2858
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
            DataField       =   "venta_codigo"
            Caption         =   "Nro.Venta"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.BB.SS"
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
            DataField       =   "concepto_venta"
            Caption         =   "Descripcion y Características del BB.SS."
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
            DataField       =   "venta_det_cantidad"
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
         BeginProperty Column04 
            DataField       =   "venta_precio_unitario_bs"
            Caption         =   "Prec.Unitario"
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
         BeginProperty Column05 
            DataField       =   "venta_descuento_bs"
            Caption         =   "Descuento"
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
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total"
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
         BeginProperty Column07 
            DataField       =   "modelo_codigo"
            Caption         =   "Unidad.Medida"
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
            DataField       =   "almacen_codigo"
            Caption         =   "Almacen"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4334.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   615.118
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   120
      Top             =   11280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6840
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc ado_datos14 
      Height          =   330
      Left            =   0
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "ado_datos14"
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
   Begin MSAdodcLib.Adodc ado_datos17 
      Height          =   330
      Left            =   9120
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_datos17"
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
      Left            =   -120
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc Ado_datos16 
      Height          =   330
      Left            =   2280
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Ado_datos16"
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
   Begin MSAdodcLib.Adodc ado_datos15 
      Height          =   330
      Left            =   6840
      Top             =   10440
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
      Caption         =   "ado_datos15"
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
   Begin MSAdodcLib.Adodc AdoDsctos 
      Height          =   330
      Left            =   11400
      Top             =   10080
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
      Caption         =   "AdoDsctos"
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2280
      Top             =   10440
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
      Caption         =   "Ado_Datos12"
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
   Begin MSAdodcLib.Adodc Ado_datos13 
      Height          =   330
      Left            =   4560
      Top             =   10440
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
      Caption         =   "Ado_datos13"
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   13680
      Top             =   10080
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
      Caption         =   "AdoAux"
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
      Left            =   4560
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   10080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc ado_datos4A 
      Height          =   330
      Left            =   9120
      Top             =   10080
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "ado_datos4A"
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
   Begin Crystal.CrystalReport CryR01 
      Left            =   720
      Top             =   11280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   4560
      Top             =   10800
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Ado_datos20"
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label lblUni_codigo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_ao_pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Public frenteservicio As String
''Ventas
'Dim rs_datos As New ADODB.Recordset     'VENTAS
'Dim rs_datos1 As New ADODB.Recordset
'Dim rs_datos2 As New ADODB.Recordset
'Dim rs_datos3 As New ADODB.Recordset
'Dim rs_datos4 As New ADODB.Recordset
'Dim rs_datos11 As New ADODB.Recordset
'Dim rs_datos12 As New ADODB.Recordset
'Dim rs_datos13 As New ADODB.Recordset
'Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
'Dim rs_datos15 As New ADODB.Recordset
'Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
'Dim rs_datos17 As New ADODB.Recordset
'Dim rs_datos18 As New ADODB.Recordset
'
'Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
'Dim rs_datos20 As New ADODB.Recordset   'Cta Bancaria
'
'Dim rs_Ventas_lista As New ADODB.Recordset
'Dim rs_aux1 As New ADODB.Recordset
'Dim rs_aux2 As New ADODB.Recordset
'Dim rs_aux3 As New ADODB.Recordset
'Dim rs_aux4 As New ADODB.Recordset
'Dim rstdestino As New ADODB.Recordset
'Dim rstcorrel_ing As New ADODB.Recordset
''CLASIFICADORES
'Dim rstdetsalalm As New ADODB.Recordset
'Dim RS_BENEF As New ADODB.Recordset
'Dim rs_TipoCambio As New ADODB.Recordset
'Dim rs_almacen2 As New ADODB.Recordset
'Dim rstacumdet As New ADODB.Recordset
'Dim rsAuxDetalle As New ADODB.Recordset
'
''==== busquedas ====
'Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim PosibleApliqueFiltro As Boolean
'Dim msgSalir As String
'Dim queryinicial As String
'Dim queryinicial2 As String
''Almacenes
'Dim descri_bien As String
'Dim Cant_Alm, VAR_CANT As Integer
'Dim correlativo1 As Integer
''VARIABLES
'Dim marca1 As Variant
'Dim swgrabar, swnuevo, deta2 As Integer
'Dim nroventa, correlv As Integer
'Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
'Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
'Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2 As Double
'Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
'Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
'Dim VAR_COD1, var_cod2, VAR_COD3 As String
'Dim VAR_CODANT, Var_Comp As Integer
'
'Private Sub CmdDetalle_Click()
'    FrmCobranza.Visible = True
'End Sub
'
'Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
'            Set DataGrid2.DataSource = Nothing
'            Set DataGrid2.DataSource = rstdetsalalm
'            DataGrid2.ReBind
'        End Select
'End Sub
'
'Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
'        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
'            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
'            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
'        Else
'            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
'            txtcorrdet.Text = " "
'            dtccodpar.Text = " "
'            dtcdescripar.Text = " "
'            txtsolpeso.Text = 0
'        End If
'    End If
'End Sub
'
'Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  Dim descri_bien As String
'  Dim Cant_Alm As Integer
'  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
'     If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then
'        If (Ado_datos.Recordset("estado_codigo") = "REG") Then
'            BtnAprobar.Visible = True
''                BtnDesAprobar.Visible = False
'            BtnModificar.Visible = True
'            BtnEliminar.Visible = True
'            BtnVer.Visible = False
''            If IsNull(Ado_datos.Recordset("venta_tipo")) Then
''                FrmABMDet.Visible = False
''                FrmABMDet2.Visible = False
''                FrmCobranza.Visible = False
''            Else
''                FrmABMDet.Visible = True
''                FrmABMDet2.Visible = True
''                FrmCobranza.Visible = True
''            End If
'            If (Ado_datos.Recordset("cobranza_fecha_prog") <= Date) Then
'                TxtDsctoTot.BackColor = &HFF&             'ROJO
'                DTPFechaProg.BackColor = &HFF&             'ROJO
'            Else
'                If (Ado_datos.Recordset("cobranza_fecha_prog") > Date) And (Ado_datos.Recordset("cobranza_fecha_prog") <= Date + 15) Then
'                    TxtDsctoTot.BackColor = &H80FF&           'NARANJA
'                    DTPFechaProg.BackColor = &H80FF&           'NARANJA
'                Else
'                    TxtDsctoTot.BackColor = &H80000013      'Fondo Oscuro
'                    DTPFechaProg.BackColor = &H80000013      'Fondo Oscuro
'                End If
'            End If
'
'        Else
'            BtnAprobar.Visible = False
''                BtnDesAprobar.Visible = True
'            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
'            BtnVer.Visible = True
''            FrmABMDet.Visible = False
''            FrmABMDet2.Visible = True
''            FrmCobranza.Visible = True
'            TxtDsctoTot.BackColor = &H80000013      'Fondo Oscuro
'            DTPFechaProg.BackColor = &H80000013      'Fondo Oscuro
'        End If
''            If Ado_datos.Recordset("estado_codigo") = "APR" Then
''                BtnAprobar.Enabled = False
'''                BtnDesAprobar.Enabled = False
''                FrmABMDet.Visible = False
''                BtnModDetalle.Visible = False
''                BtnAnlDetalle.Visible = False
''            Else
''                BtnAprobar.Enabled = True
''                FrmABMDet.Visible = True
''                BtnModDetalle.Visible = True
''                BtnAnlDetalle.Visible = True
''            End If
''            If (Ado_datos.Recordset("venta_tipo") = "C") And Ado_datos.Recordset("estado_codigo") = "APR" Then
''                FrmABMDet2.Visible = True
''                FrmCobranza.Visible = True
''            Else
''                FrmABMDet2.Visible = False
''                FrmCobranza.Visible = False
''            End If
''        If (Ado_datos.Recordset("venta_tipo") = "C") Then
''            TxtPlazo.Visible = True
''            BtnAddDetalle2.Visible = True
''        Else
''            TxtPlazo.Visible = False
''            If Ado_datos.Recordset("venta_tipo") = "E" Then
''                BtnAddDetalle2.Visible = False
''            End If
''        End If
'
''        If Dtc_deudor2.Text = "SI" Then
''            Dtc_deudor2.BackColor = &HFF&
''        Else
''            Dtc_deudor2.BackColor = &H80000010
''        End If
'        'If Ado_datos.Recordset("beneficiario_codigo") <> "" And Ado_datos.Recordset("beneficiario_codigo") <> "VD" Then
'        If Ado_datos.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'            'RS_BENEF.Recordset.Requery
'            If RS_BENEF.RecordCount > 0 Then
'                If RS_BENEF!beneficiario_deudor = "SI" Then
'                    Dtc_deudor2.BackColor = &HFF&
'                Else
'                    Dtc_deudor2.BackColor = &H80000010
'                End If
'            End If
'
'        End If
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
'        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'        Set ado_datos14.Recordset = rs_datos14
'        ado_datos14.Recordset.Requery
'        If ado_datos14.Recordset.RecordCount > 0 Then
'            deta2 = 1
'            'TxtMontoBs.Text = Ado_datos.Recordset!monto_total_bS
'            'TxtMontoUs.Text = Ado_datos.Recordset!deuda_cobrada
'            'Text2.Text = Ado_datos.Recordset!saldo_p_cobrar
'            Call AbreAlmacen
''            If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Then
''                FrmABMDet2.Visible = True
''                FrmCobranza.Visible = True
''
''            Else
''                FrmABMDet2.Visible = False
''                FrmCobranza.Visible = False
''            End If
'        Else
'            deta2 = 0
''            'TxtMontoBs.Text = 0
''            'TxtMontoUs.Text = 0
''            'Text2.Text = 0
''            FrmABMDet2.Visible = False
''            FrmCobranza.Visible = False
'        End If
'
'        Set rs_datos16 = New ADODB.Recordset
'        If rs_datos16.State = 1 Then rs_datos16.Close
'        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos16.Recordset = rs_datos16
'        Ado_datos16.Recordset.Requery
'        If Ado_datos16.Recordset.RecordCount > 0 Then
'            FrmCobranza.Visible = True
'            'BtnImprimir2.Visible = True
'            'BtnImprimir3.Visible = True
'        Else
'            FrmCobranza.Visible = False
'            'BtnImprimir2.Visible = False
'            'BtnImprimir3.Visible = False
'        End If
'
'        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
'        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'
'        TxtCobrador = Trim(dtc_desc4A.Text)
'     End If
'     FrmDetalle.Visible = True
'     FrmCobranza.Visible = True
'  Else
'    BtnAprobar.Visible = False
''                BtnDesAprobar.Visible = True
'    BtnModificar.Visible = False
'    BtnEliminar.Visible = False
'    BtnVer.Visible = False
'    FrmDetalle.Visible = False
'    FrmCobranza.Visible = False
'    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
'  End If
'End Sub
'
'Private Sub AbreAlmacen()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh
'
'End Sub
'
'Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
' If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
'    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
'        BtnModDetalle.Visible = True
'        BtnImprimir4.Visible = True
''        If (Ado_datos16.Recordset("estado_codigo") = "REG") Then
''            'If (Ado_datos.Recordset("estado_codigo") = "APR") Then
''            '    BtnAprobar2.Visible = True
''            'Else
''            '    BtnAprobar2.Visible = False
''            'End If
''            BtnImprimir2.Visible = True
''            BtnImprimir3.Visible = True
''            BtnAnlDetalle2.Visible = True
''            BtnModDetalle2.Visible = True
''            'If (Ado_datos16.Recordset("doc_numero") > 0) Then
''            '    BtnImprimir3.Visible = True
''            'Else
''            '    BtnImprimir3.Visible = False
''            'End If
''        End If
''        If (Ado_datos16.Recordset("estado_codigo") = "APR") Then
''            BtnImprimir2.Visible = True
''            BtnImprimir3.Visible = True
''            BtnAnlDetalle2.Visible = False
'''            BtnModDetalle2.Visible = False
''        End If
''        If (Ado_datos16.Recordset("estado_codigo") = "ANL") Then
''            BtnImprimir2.Visible = False
''            BtnImprimir3.Visible = False
''            BtnAnlDetalle2.Visible = False
'''            BtnModDetalle2.Visible = False
''            BtnImprimir3.Visible = False
''        End If
'    Else
'        'BtnAprobar2.Visible = False
'        'BtnImprimir2.Visible = False
'        BtnImprimir4.Visible = False
'        'BtnAnlDetalle2.Visible = False
'        BtnModDetalle.Visible = False
'    End If
' Else
'    'BtnAprobar2.Visible = False
'    'BtnImprimir2.Visible = False
'    BtnImprimir4.Visible = False
'    'BtnAnlDetalle2.Visible = False
'    BtnModDetalle.Visible = False
' End If
'End Sub
'
'Private Sub BtnAñadir_Click()
'marca1 = Ado_datos.Recordset.Bookmark
'  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
'  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
'    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
'    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
'        swnuevo = 1
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        FrmCobros.Visible = True
'        FrmCobros.Enabled = True
'        fraOpciones.Enabled = False
'        FraNavega.Enabled = False
'        FrmDetalle.Visible = False
'        FrmCobranza.Visible = False
'        FrmABMDet.Visible = False
'        FrmABMDet2.Visible = False
'        TxtCobrador.Visible = False
'        Ado_datos16.Recordset.AddNew
'        dtc_codigo2A.Text = dtc_codigo2.Text
'        dtc_desc2A.Text = dtc_desc2.Text
'        TxtMonto.SetFocus
'        DTPFechaProg.Visible = True
'        DTPFechaCobro.Visible = True
'        Lbl_nombre_fac.Caption = "Cliente :"
'        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
'        Txt_parche.Visible = True
'        'Ado_datos.Recordset.Move marca1 - 1
'    Else
'        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
'    End If
'  Else
'    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'Private Sub BtnAprobar_Click()
' If Ado_datos.Recordset.RecordCount > 0 Then
'     If IsNull(Ado_datos.Recordset("cobranza_observaciones")) Or (Ado_datos.Recordset("cobranza_deuda_bs") = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'        Exit Sub
'     Else
'        If Ado_datos.Recordset("estado_codigo") = "REG" Then
'           sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'           If sino = vbYes Then
'               'If Ado_datos.Recordset("venta_tipo") = "C" Or Ado_datos.Recordset("venta_tipo") = "V" Then
'               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
'               'End If
'               gestion0 = Ado_datos.Recordset("ges_gestion")
'               correlv = Ado_datos.Recordset("venta_codigo")
'               nroventa = Ado_datos.Recordset("venta_codigo")
'               ' APRUEBA ao_ventas_cabecera
'               db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                    Ado_datos.Recordset!doc_numero = rs_aux2!correl_doc
'                    'Txt_campo1.Caption = rs_aux2!correl_doc
'                    rs_aux2.Update
'                End If
'                ' GRABA Nombre de Archivo en ao_ventas_cabecera
'
'                'VAR_ARCH = RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " "
'                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " "
'
'
'               'marca1 = Ado_datos.Recordset.Bookmark
'               'Ado_datos.Recordset.Requery
'        '       Ado_datos.Refresh
'               'Ado_datos.Recordset.Move marca1 - 1
'
'               '  Set rstacumdet = New ADODB.Recordset
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'                '  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranzas where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
'                '
'                '  Set rstdestino = New ADODB.Recordset
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & gestion0 & "' and venta_codigo = " & nroventa, db, adOpenKeyset, adLockOptimistic
'                '  If rstdestino.RecordCount > 0 Then
'                '    rstdestino!deuda_cobrada = rstacumdet!Cobrobs
'                '    rstdestino!saldo_p_cobrar = (rstdestino!monto_total_Bs - rstdestino!monto_cobrado - rstdestino!deuda_cobrada)
'                '    rstdestino.Update
'                '  End If
'                '  If rstdestino.State = 1 Then rstdestino.Close
'                '  If rstacumdet.State = 1 Then rstacumdet.Close
'
'               Call Contabiliza_venta
'               Call OptFilGral1_Click
'           End If
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'Private Sub BtnBuscar_Click()
''JQA
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'      PosibleApliqueFiltro = False
'      Dim rsNada As ADODB.Recordset
'      Dim GrSqlAux As String
'      Set ClBuscaGrid = New ClBuscaEnGridExterno
'      Set ClBuscaGrid.Conexión = db
'      ClBuscaGrid.EsTdbGrid = False
'      Set ClBuscaGrid.GridTrabajo = dg_datos
'      ClBuscaGrid.QueryUtilizado = queryinicial
'      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'      ClBuscaGrid.CamposVisibles = "110"
'      ClBuscaGrid.Ejecutar
'      PosibleApliqueFiltro = True
'  Else
'    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub
'
'Private Sub BtnCancelar_Click()
'  'Ado_datos.Refresh
'  fraOpciones.Visible = True
'  FraGrabarCancelar.Visible = False
'  marca1 = Ado_datos.Recordset.Bookmark
'  If Ado_datos.Recordset("estado_codigo") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
'  FraNavega.Enabled = True
'  FrmCobros.Enabled = False
'  'Fra_datos.Enabled = True
'  FrmDetalle.Visible = True
'  FrmCobranza.Visible = True
'  'Fra_Total.Visible = True
'  dg_datos.Visible = True
'  FrmABMDet.Visible = True
'  FrmABMDet2.Visible = True
'
'  SSTab1.Tab = 0
'  SSTab1.TabEnabled(0) = True
'  SSTab1.TabEnabled(1) = False
'  SSTab1.TabEnabled(2) = False
'  'Ado_datos.Recordset.Move marca1 - 1
'  BtnImprimir2.Visible = True
'  BtnImprimir3.Visible = True
'
'  swnuevo = 0
'
'End Sub
'
'Private Sub BtnEliminar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset("estado_codigo") = "REG" Then
'      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
'      If sino = vbYes Then
'          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
'          'Dim rstdestino As New ADODB.Recordset
'          'Set rstdestino = New ADODB.Recordset
'          'If rstdestino.State = 1 Then rstdestino.Close
'          'rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  ", db, adOpenDynamic, adLockOptimistic
'          'If Not rstdestino.BOF Then rstdestino.MoveFirst
'          'If Not rstdestino.BOF And Not rstdestino.EOF Then
'          '    rstdestino("estado_codigo") = "E"
'          '    rstdestino.Update
'          'End If
'          'If rstdestino.State = 1 Then rstdestino.Close
'          marca1 = Ado_datos.Recordset.Bookmark
'          'Ado_datos.Recordset.Requery
'          'Ado_datos.Refresh
'          Call OptFilGral1_Click
'          Ado_datos.Recordset.Move marca1 - 1
'      End If
'    Else
'      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
'    End If
'  Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'
'Private Sub BtnGrabar_Click()
'  If dtc_codigo4A = "" Then
'    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
'  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
'    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
'  If TxtObs = "" Then
'    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
'  'If swnuevo = 2 Then
'  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
''  If DTPFechaProg.Visible = False Then
''    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
''       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''      Exit Sub
''    End If
''  End If
'  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  'valida = 1
'  'If valida = 1 And dtc_codigo4A <> "" Then
''    Set rstdestino = New ADODB.Recordset
''    If rstdestino.State = 1 Then rstdestino.Close
'    db.BeginTrans
'    If swnuevo = 1 Then
''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
''      Set Ado_datos16.Recordset = rstdestino
''      Ado_datos16.Recordset.AddNew
'      Ado_datos.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
'      Ado_datos.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
'      'Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg                                'Fecha Programada a Cobrar
'    End If
'      Ado_datos.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
'      Ado_datos.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
'      'Ado_datos.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'      Ado_datos.Recordset!cobranza_deuda_bs = CDbl(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos.Recordset!cobranza_deuda_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
'      'If TxtDscto.Text = "" Or TxtDscto.Text = "0" Or TxtDscto.Text = "0.00" Then
'        Ado_datos.Recordset!cobranza_descuento_bs = 0                                 'Descuento Bs
'        Ado_datos.Recordset!cobranza_descuento_dol = 0                                    'Descuento Dol
'      'Else
'      '  Ado_datos.Recordset!cobranza_descuento_bs = CDbl(TxtDscto.Text)                                 'Descuento Bs
'      '  Ado_datos.Recordset!cobranza_descuento_dol = CDbl(TxtDscto.Text) / GlTipoCambioMercado        'Descuento Dol
'      'End If
'      Ado_datos.Recordset!cobranza_total_bs = Ado_datos.Recordset!cobranza_deuda_bs - Ado_datos.Recordset!cobranza_descuento_bs               'Monto Total Bs
'      Ado_datos.Recordset!cobranza_total_dol = Ado_datos.Recordset!cobranza_deuda_dol - Ado_datos.Recordset!cobranza_descuento_dol               'Monto Total Dol
'      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'      If Ado_datos.Recordset!cobranza_total_bs <> 0 Then
'            Ado_datos.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
'      End If
'      'Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
'      'Call acumulaMont(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!correl_venta, Ado_datos.Recordset!venta_codigo)
'      Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
'
'      Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text
'      Ado_datos.Recordset!proceso_codigo = "COM"
'      Ado_datos.Recordset!subproceso_codigo = "COM-02"
'      Ado_datos.Recordset!etapa_codigo = "COM-02-02"
'      Ado_datos.Recordset!clasif_codigo = "ADM"
'      Ado_datos.Recordset!doc_codigo = IIf(lbl_doc1 = "", "R-105", lbl_doc1)
'      Ado_datos.Recordset!doc_numero = IIf(lbl_docnro = "", "0", lbl_docnro)
'      Ado_datos.Recordset!doc_codigo_fac = "R-101"
'      If Ado_datos.Recordset!factura_impresa = "N" Then
'         TxtCmpbte.Caption = "0"
'         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
'      Else
'         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
'      End If
'      Ado_datos.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion = "", "0", Trim(TxtAutorizacion))
'      Ado_datos.Recordset!poa_codigo = "3.1.2"
'      Ado_datos.Recordset!Cta_Codigo = IIf(dtc_cta.Text = "", "0", dtc_cta.Text)
'      'If DTPFechaProg.Visible = False Then
'        Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value         'Fecha de Cobranza
'      '  'Ado_datos.Recordset!estado_codigo = "APR"
'      'Else
'      '  Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaProg.Value         'Fecha de Cobranza
'      '  Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg.Value           'Fecha Programada de Cobranza
'
'      'End If
'      Ado_datos.Recordset!estado_codigo = "REG"
'      Ado_datos.Recordset!usr_codigo = GlUsuario
'      Ado_datos.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      Ado_datos.Recordset.Update
'    db.CommitTrans
'    'Ado_datos.Recordset!doc_numero = Ado_datos.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"
'  If swnuevo = 1 Then
'    'Call abre_solicitud_lista
'    'rc_Cobranza.Requery
'    'Ado_datos.Refresh
'    'Ado_datos.Recordset.MoveLast
'  End If
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'    FraNavega.Enabled = True
'    fraOpciones.Visible = True
'    FraGrabarCancelar.Visible = False
'    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
'    FrmCobros.Enabled = False
'    TxtCobrador.Visible = True
'    FrmABMDet.Visible = True
'    FrmABMDet2.Visible = True
'    BtnImprimir2.Visible = True
'    BtnImprimir3.Visible = True
'
'    swnuevo = 0
'
'  'Else
'  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
'  'End If
'
'End Sub
'
'Private Sub BtnImprimir_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant, i%, y%
'    Dim co As New ADODB.Command
'
''    Dim rs As New ADODB.Recordset
''    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
''            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
''    i = 1
''    y = 1
'    CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_nota_de_venta.rpt"
'    CryV01.WindowShowRefreshBtn = True
'    CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'    CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'    iResult = CryV01.PrintReport
'    If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub
'
'Private Sub BtnImprimir3_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If (Ado_datos.Recordset!factura_impresa = "N") And (Ado_datos.Recordset!cobranza_deuda_bs <> "0.00") Then
'        '===== ini GENERA EL CODIGO DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FACTURA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            VAR_COD1 = CDbl(rs_aux1!numero_correlativo) + 1
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!numero_correlativo = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'            Else
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'        db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'
'        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          var_cod2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
'        db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & var_cod2 & " Where ao_ventas_cobranza.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'        'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"
'        db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'        'IMPRIMIR FACTURA
''        Dim iResult As Variant  ', i%, y%
''        CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-103_recibo_cobranza.rpt"
''        CryR01.WindowShowRefreshBtn = True
''        CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''        CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''        CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
''
''        CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
''        CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
''        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
''        iResult = CryR01.PrintReport
''        If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'        TxtCmpbte.Caption = VAR_COD1
'        If Ado_datos.Recordset("estado_codigo") = "REG" Then
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'     Else
'        MsgBox "La Factura Nro. " + Ado_datos.Recordset!cobranza_nro_factura + " ya fue Impresa", , "Atención"
'     End If
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub
'
'Private Sub BtnImprimir4_Click()
'  If Ado_datos16.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-105_kardex.rpt"
'    CryR01.WindowShowRefreshBtn = True
'    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo
'    'Literal por el Total de la Compra
'    var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
'    CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'    'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_prog_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub
'
'Private Sub BtnModificar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset!estado_codigo = "REG" And (Ado_datos16.Recordset!venta_tipo = "E" Or Ado_datos16.Recordset!venta_tipo = "V" Or Ado_datos16.Recordset!venta_tipo = "C") Then
'      FraNavega.Enabled = False
'      fraOpciones.Visible = False
'      FraGrabarCancelar.Visible = True
'      FrmDetalle.Visible = False
'      FrmCobranza.Visible = False
'      'swgrabar = 0
'      swnuevo = 2
'      TxtCobrador.Visible = False
'      'TxtMonto.SetFocus
'      'TxtNroVenta.Enabled = False
'      'marca1 = ado_datos14.Recordset.BookMark
'      'txt_descripcion_venta.Enabled = True
'      'TxtNroVenta.Text = txt_venta.Text
'      'lbltipoVenta.Caption = dtc_desc11.Text
'      'lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
'      SSTab1.Tab = 0
'      SSTab1.TabEnabled(0) = True
'      SSTab1.TabEnabled(1) = False
'      SSTab1.TabEnabled(2) = False
'      FrmCobros.Visible = True
'      FrmCobros.Enabled = True
'      FrmABMDet.Visible = False
'      FrmABMDet2.Visible = False
'
'      BtnImprimir2.Visible = False
'      BtnImprimir3.Visible = False
'      If Ado_datos.Recordset!factura_impresa = "N" Then
'      '    sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
'      '    If sino = vbYes Then
'              'DTPFechaProg.Visible = True
'              DTPFechaCobro.Visible = True
'              Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'              lbl_fechas.Caption = "Fecha de Cobranza"
'              TxtCmpbte.Caption = "0"
'      '        Txt_parche.Visible = False      '&H80000013&
'      '        'dtc_desc2A.BackColor = &H80000013
'      '    Else
'      '        DTPFechaProg.Visible = True
'      '        DTPFechaCobro.Visible = False
'      '        Lbl_nombre_fac.Caption = "Cliente :"
'      '        lbl_fechas.Caption = "Fecha Programada de Cobranza"
'      '        Txt_parche.Visible = True       '&H80000005&
'      '        'dtc_desc2A.BackColor = &H80000005
'      '    End If
'      Else
'      '    DTPFechaProg.Visible = True
'      '    DTPFechaCobro.Visible = False
'      '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
'      End If
'      TxtMonto.Text = CDbl(TxtDsctoTot)
'      TxtMonto.SetFocus
'    Else
'      MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'    End If
'  Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'Private Sub BtnSalir_Click()
'    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
'    If sino = vbYes Then
''        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
''        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
''        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
''        If rs_datos14.State = 1 Then rs_datos14.Close
''        If rs_Ventas.State = 1 Then rs_Ventas.Close
'        Unload Me
'    End If
'End Sub
'
''Private Sub Cmd_Cliente_Click()
''    glPersNew = "P"
''    frmBeneficiario.Show 'vbModal
''End Sub
'
'Private Sub CmdCancelaCobro_Click()
'End Sub
'
''Private Sub CmdCancelaDet_Click()
''  'TxtNroVenta.Enabled = True
''  FrmEdita.Enabled = False
''  swgrabar = 0
''  'Call cerea
''  swnuevo = 0
''  'cmdElige.Enabled = False
''  marca1 = Ado_datos.Recordset.Bookmark
''  If Ado_datos.Recordset("estado_codigo") = "APR" Then
''    Call OptFilGral2_Click
''  Else
''    Call OptFilGral1_Click
''  End If
''    SSTab1.Tab = 0
''    SSTab1.TabEnabled(0) = True
''    SSTab1.TabEnabled(1) = False
''    SSTab1.TabEnabled(2) = False
''    FraNavega.Enabled = True
''    FrmDetalle.Enabled = True
''    'FrmDetalle.Visible = True
''    FrmCobranza.Visible = True
''    FrmABMDet.Visible = True
''    FrmABMDet2.Visible = True
''  'ado_datos14.Refresh
''  'Ado_datos.Recordset.Move marca1 - 1
''End Sub
'
'Private Sub BtnModDetalle2_Click()
'  If ado_datos14.Recordset.RecordCount > 0 Then
'    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = False
'
'    FrmEdita.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No Existen Bienes Registrados, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'Private Sub BtnDesAprobar_Click()
''  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
''  If sino = vbYes Then
''    Dim rstdestino As New ADODB.Recordset
''    Set rstdestino = New ADODB.Recordset
''    If rstdestino.State = 1 Then rstdestino.Close
''    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
''    If Not rstdestino.BOF Then rstdestino.MoveFirst
''    If Not rstdestino.BOF And Not rstdestino.EOF Then
''      rstdestino("estado_codigo") = "REG"
''      rstdestino.Update
''    End If
''    If rstdestino.State = 1 Then rstdestino.Close
''    marca1 = Ado_datos.Recordset.Bookmark
''    Call OptFilGral1_Click
''    Ado_datos.Recordset.Move marca1 - 1
''  End If
'End Sub
'
''Private Sub CmdDetallePoa_Click()
''  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
''   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
''  Else
''    marca1 = Ado_datos.Recordset.BookMark
''    FrmPoasCapturaALB.Lblformulario = "F11"
''    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
''    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
''    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
''    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
''    FrmPoasCapturaALB.Show vbModal
''  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
''    '
''  Else
''    Ado_datos.Refresh
''    Ado_datos.Recordset.Move marca1 - 1
''  End If
''  End If
''End Sub
'
''Private Sub cmdElige_Click()
''  With ALFrmMateriales
''        .ALPrincipal
''        If .QResp Then
''            txtCodigo.Text = .QCodigo
''            txtDesc.Text = .QItem
''        End If
''    End With
''    Txtcant_alm = 0
''    Cant_Alm = 0
''    DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
''    Txtcant_alm = Cant_Alm
''    If Cant_Alm >= TxtCantPedi Then
''        optSi = True
''    Else
''        optNo = True
''    End If
''End Sub
'
'Private Sub Contabiliza_venta()
''    Call graba_proyecto
'    Call graba_ingreso
'  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
'  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
'  'If sino = vbYes Then
'    ' INI CORRECCION 18-JUN-2014
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    fte_codigo1 = VAR_FTE
'    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    Select Case VAR_CODTIPO
'        Case "DEI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'              'cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'              'Subcta_deb11 = rstdestino!Subcta_cred1
'              'Subcta_deb21 = rstdestino!Subcta_cred2
'
'              'cta_credito1 = rstdestino2!cta_deb
'              'Subcta_cred11 = rstdestino2!Subcta_deb1
'              'Subcta_cred21 = rstdestino2!Subcta_deb2
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "REC"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
'              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
'                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
'                Exit Sub
'              End If
'            End If
'            If rs_aux1.State = 1 Then rs_aux1.Close
'
'        Case "DYR"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "DES"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "ANI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "DVI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'            '' 02/07/2014 VERIFICAR
'            'If rstdestino.State = 1 Then rstdestino.Close
'            'rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'            'If rstdestino2.State = 1 Then rstdestino2.Close
'            'rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'            '  Exit Sub
'            'End If
'        Case Else
'            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'            If rstdestino.State = 1 Then rstdestino.Close
'            Exit Sub
'    End Select
'    'If rstdestino.State = 1 Then rstdestino.Close
'    '**** FIN VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
'
'    Dim cta_deb1 As String
'    Dim Subcta_deb11 As String
'    Dim Subcta_deb21 As String
'
'    Dim cta_credito1 As String
'    Dim Subcta_cred11 As String
'    Dim Subcta_cred21 As String
'
'    Dim cod_ant As Integer
'    Dim org_ant As String
'
'    'If DtCCta_codigo.Text <> "01" Then
'    '  If rstdestino.State = 1 Then rstdestino.Close
'    '  rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
'    '  If Not rstFc_cuenta_bancaria.EOF Then
'    '    fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
'    '  Else
'    '  End If
'    'Else
'    '    fte_codigo1 = Me.DtCFte_codigo.Text
'    'End If
'    'If VAR_CODTIPO = "DEI" Or VAR_CODTIPO = "DES" Then
'    '  fte_codigo1 = Me.DtCFte_codigo.Text
'    'End If
'
''    fte_codigo1 = VAR_FTE
''
''    Dim i As Integer
''    Dim j As Integer
''    Dim v_Tipo_Comp(1, 2)
''
''    v_Tipo_Comp(1, 1) = VAR_CODTIPO
'
''    If VAR_CODTIPO = "DYR" Then
''      'j = 2
''      'v_Tipo_Comp(1, 1) = "CAD"
''      'v_Tipo_Comp(1, 2) = "CAR"
''      j = 2
''      v_Tipo_Comp(1, 1) = "DYR"
''    Else
''      j = 1
''      v_Tipo_Comp(1, 1) = IIf(VAR_CODTIPO = "DEI", "DEI", IIf(VAR_CODTIPO = "REC", "REC", IIf(VAR_CODTIPO = "DES", "DES", IIf(VAR_CODTIPO = "ANI", "ANI", ""))))
''    End If
''
''    If VAR_CODTIPO = "DVI" Then
''      j = 1
''      v_Tipo_Comp(1, 1) = "DVI"
''    End If
'
''    For i = 1 To j
''      If rstdestino.State = 1 Then rstdestino.Close
''      If v_Tipo_Comp(1, i) = "DEI" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "REC" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DYR" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DES" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "ANI" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DVI" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DVI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "" Then
''        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
''        Exit Sub
''      End If
'
'    ' INI CORRECCION 18-JUN-2014
''      If v_Tipo_Comp(1, i) = "DVI" Then
''        ' 02/07/2014 VERIFICAR
''        If rs_aux2.State = 1 Then rs_aux2.Close
''        rs_aux2.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
''        If rstdestino2.State = 1 Then rstdestino2.Close
''        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''        If rs_aux2.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
''          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
''          Exit Sub
''        End If
''      End If
''
''      If rs_aux2.RecordCount < 1 Then
''        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
''        Exit Sub
''      End If
''    Next
'
'    'If rstdestino.State = 1 Then rstdestino.Close
'
'    v_Tipo_Comp(1, 1) = VAR_CODTIPO
'
'    db.BeginTrans
''    Frmmensaje.Visible = True
''    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
'    '========================================
'    '==== verifica si ya fue contabilizado
'      yacontabilizo = 0
'      Set rs_aux2 = New ADODB.Recordset
'      If rs_aux2.State = 1 Then rs_aux2.Close
'      rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '" & VAR_CODANT & "' and org_codigo = '" & VAR_ORG & "' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
'      If rs_aux2.RecordCount > 0 Then
'        ' revisar para validar mejor si YA contabilizo !!
'        'yacontabilizo = 1
'        yacontabilizo = 0
'      Else
'        yacontabilizo = 0
'      End If
'      If yacontabilizo = 1 Then
'        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
'        Var_Comp = rs_aux2!Cod_Comp
'      Else
'        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
'        Set rstCodComp = New ADODB.Recordset
'        rstCodComp.CursorLocation = adUseClient
'        If rstCodComp.State = 1 Then rstCodComp.Close
'        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
'        If rstCodComp.RecordCount > 0 Then
'          Var_Comp = CDbl(rstCodComp!numero_correlativo)
'          Var_Comp = Var_Comp + 1
'          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
'          rstCodComp.Update
'        End If
'        If rstCodComp.State = 1 Then rstCodComp.Close
'        '===== fin TERMINA GENERACION DE COMPROBANTE =====
'
'      '==== ini registro co_comprobante_m
'
'        rs_aux2.AddNew
'        rs_aux2("cod_comp") = Var_Comp
'      End If
'    '========================================
'    'anterior
'    '      If rstdestino.State = 1 Then rstdestino.Close
'    '      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
'    '      If rstdestino.RecordCount > 0 Then
'    '      End If
'    '      rstdestino.AddNew
'
'    '      rstdestino("cod_comp") = Var_Comp
'    'anterior
'      rs_aux2("Tipo_Comp") = VAR_CODTIPO        'v_Tipo_Comp(1, i)
'      rs_aux2("cod_trans") = VAR_CODANT
'      rs_aux2("org_codigo") = VAR_ORG
'      rs_aux2("ges_gestion") = Year(Date)
'      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
'      If yacontabilizo = 0 Then
'        rs_aux2("Fecha_transacion") = Date
'      End If
'      rs_aux2("beneficiario_codigo") = VAR_BENEF
'      rs_aux2("glosa") = "CONTABILIZA INGRESO DE: " + VAR_GLOSA
'      rs_aux2("unidad_codigo") = Ado_datos16.Recordset("unidad_codigo")
'      rs_aux2("solicitud_codigo") = Ado_datos16.Recordset("solicitud_codigo")
'      rs_aux2("tipo_moneda") = VAR_MONEDA
'      rs_aux2("unidad_codigo_ant") = VAR_CITE
'
'      rs_aux2("proceso_codigo") = "FIN"
'      rs_aux2("subproceso_codigo") = "FIN-02"
'      Select Case VAR_CODTIPO
'        Case "DEI"
'            rs_aux2("etapa_codigo") = "FIN-02-01"
'        Case "REC"
'            rs_aux2("etapa_codigo") = "FIN-02-02"
'        Case "DYR"
'            rs_aux2("etapa_codigo") = "FIN-02-01"
'        Case "DES"
'            rs_aux2("etapa_codigo") = "FIN-02-01"
'        Case "ANI"
'            rs_aux2("etapa_codigo") = "FIN-02-02"
'        Case "DVI"
'            rs_aux2("etapa_codigo") = "FIN-02-02"
'      End Select
'
'      rs_aux2("clasif_codigo") = "ADM"
'      rs_aux2("doc_codigo") = "R-128"
'      rs_aux2("doc_numero") = Var_Comp
'      rs_aux2("pro_codigo_det") = VAR_PROY2
'
'      rs_aux2("estado_codigo") = "APR"
'
'      If yacontabilizo = 0 Then
'        rs_aux2("usr_codigo") = GlUsuario
'        rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
'        rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
'      End If
'      rs_aux2.Update
'      '==== fin registro co_comprobantre_m
'
'    Dim d_cta_nombre_1 As String
'    Dim d_aux1_1 As String
'    Dim d_aux2_1 As String
'    Dim d_aux3_1 As String
'    Dim h_cta_nombre_1 As String
'    Dim h_aux1_1 As String
'    Dim h_aux2_1 As String
'    Dim h_aux3_1 As String
'    'If rstdestino.State = 1 Then rstdestino.Close
'
'    For i = 1 To j
''    ' nuevo ini
''      If v_Tipo_Comp(1, i) = "DEI" Then     'Devengado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DYR' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DES' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'ANI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
'
''      If v_Tipo_Comp(1, i) = "DVI" Then
''        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
''        If rstdestino.State = 1 Then rstdestino.Close
''        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEI' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
''        If rstdestino2.State = 1 Then rstdestino2.Close
''        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'REC' and rec_rub_i <= " & (VAR_PARTIDA) & " and rec_rub_f >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
''          cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
''          Subcta_deb11 = rstdestino!Subcta_cred1
''          Subcta_deb21 = rstdestino!Subcta_cred2
''
''          cta_credito1 = rstdestino2!cta_deb
''          Subcta_cred11 = rstdestino2!Subcta_deb1
''          Subcta_cred21 = rstdestino2!Subcta_deb2
''        Else
''          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'''          Exit Sub
''        End If
''      End If
''
''      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
''        cta_deb1 = rstdestino("cta_deb")
''        Subcta_deb11 = rstdestino("Subcta_deb1")
''        Subcta_deb21 = rstdestino("Subcta_deb2")
''        cta_credito1 = rstdestino("cta_cred")
''        Subcta_cred11 = rstdestino("Subcta_cred1")
''        Subcta_cred21 = rstdestino("Subcta_cred2")
''      Else
''        'MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''        'Exit Sub
''
''      End If
'
'      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'
'        cta_credito1 = rstdestino("cta_cred")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
'      Else
'        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'        Subcta_deb11 = rstdestino!Subcta_cred1
'        Subcta_deb21 = rstdestino!Subcta_cred2
'
'        cta_credito1 = rstdestino!cta_deb
'        Subcta_cred11 = rstdestino!Subcta_deb1
'        Subcta_cred21 = rstdestino!Subcta_deb2
'      End If
'
'      If rs_aux1.State = 1 Then rs_aux1.Close
'      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
'      If rs_aux1.RecordCount > 0 Then
'        d_cta_nombre_1 = rs_aux1("NombreCta")
'        d_aux1_1 = rs_aux1("aux1")
'        d_aux2_1 = rs_aux1("aux2")
'        d_aux3_1 = rs_aux1("aux3")
'      End If
'      If rs_aux1.State = 1 Then rs_aux1.Close
'      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
'      If rs_aux1.RecordCount > 0 Then
'        h_cta_nombre_1 = rs_aux1("NombreCta")
'        h_aux1_1 = rs_aux1("aux1")
'        h_aux2_1 = rs_aux1("aux2")
'        h_aux3_1 = rs_aux1("aux3")
'      End If
'    ' nuevo fin
'
'      '===== ini registra CO_diaRIO =========
'      Set rstdestino2 = New ADODB.Recordset
'      If rstdestino2.State = 1 Then rstdestino2.Close
'      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
'      'If rstdestino2.RecordCount > 0 Then
'      '  MsgBox "Ya Existe el asiento, se reemplazará con los nuevos datos..."
'      'Else
'        rstdestino2.AddNew
'        rstdestino2("Cod_Comp") = Var_Comp
'      'End If
'        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
'      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
'      'rstdestino2("Cod_Comp_C") = Var_Comp
'      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
'      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
'        rstdestino2("D_Cuenta") = cta_deb1
'        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        ' para Aux1
''        Select Case d_aux1_1
''                Case "01"
''                    VAR_COD1 = VAR_BENEF
''                Case "02"
''                    VAR_COD1 = VAR_CTA
''                Case "03"
''                    VAR_COD1 = VAR_PROY2
''                Case "04"
''                    VAR_COD1 = Ado_datos.Recordset("unidad_codigo")
''                Case "05"
''                    VAR_COD1 = ""
''                Case "06"
''                    VAR_COD1 = ""
''                Case "07"
''                    VAR_COD1 = ""
''                Case "08"
''                    VAR_COD1 = ""
''                Case "09"
''                    VAR_COD1 = VAR_ORG
''                Case "10"
''                    VAR_COD1 = ""
''                Case "11"
''                    VAR_COD1 = ""
''                Case "12"
''                    VAR_COD1 = ""
''        End Select
'        ' ini PARA EL FUTURO ******** REVISAR
''        Set rs_aux4 = New ADODB.Recordset
''        If rs_aux4.State = 1 Then rs_aux4.Close
''        SQL_FOR = "select * from cc_tipo_auxiliar where aux = '" & d_aux1_1 & "' "
''        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''        If rs_aux4.RecordCount > 0 Then
''            Set rs_aux1 = New ADODB.Recordset
''            If rs_aux1.State = 1 Then rs_aux1.Close
''            SQL_FOR = "select * from " + rs_aux4!NombreTabla + " where " + rs_aux4!nombre_codigo + " = " + VAR_COD1
''            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''            If rs_aux1.RecordCount > 0 Then
''        Else
''        End If
'        ' fin PARA EL FUTURO ******** REVISAR
'        Select Case d_aux1_1
'            Case "01"
'                rstdestino2("D_Cta_Aux1") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux1") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux1") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux1") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux1") = ""
'        End Select
'
'        Select Case d_aux2_1
'            Case "01"
'                rstdestino2("D_Cta_Aux2") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux2") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux2") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux2") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux2") = ""
'        End Select
'
'        Select Case d_aux3_1
'            Case "01"
'                rstdestino2("D_Cta_Aux3") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux3") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux3") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux3") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux3") = ""
'        End Select
''        If d_aux1_1 = "01" Then
''          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''        End If
''        If d_aux1_1 = "02" Then
''          rstdestino2("D_Cta_Aux1") = VAR_CTA
''        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
'        'AQUI MONEDA 02/07/01
'        'rstdestino2("D_Cambio") = GlTipoCambioMercado
'        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
'        rstdestino2("H_Cuenta") = cta_credito1
'        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_cred11
'        rstdestino2("H_SubCta2") = Subcta_cred21
'        rstdestino2("H_Aux1") = h_aux1_1
'        rstdestino2("H_Aux2") = h_aux2_1
'        rstdestino2("H_Aux3") = h_aux3_1
'        'rstdestino2("H_Cta_Aux1") = ""
'        Select Case h_aux1_1
'            Case "01"
'                rstdestino2("H_Cta_Aux1") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux1") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux1") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux1") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux1") = ""
'        End Select
'
'        Select Case h_aux2_1
'            Case "01"
'                rstdestino2("H_Cta_Aux2") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux2") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux2") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux2") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux2") = ""
'        End Select
'
'        Select Case h_aux3_1
'            Case "01"
'                rstdestino2("H_Cta_Aux3") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux3") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux3") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux3") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux3") = ""
'        End Select
'
''        If h_aux1_1 = "01" Then
''          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''          'DtCCta_descripcion_larga
''        End If
''        If h_aux1_1 = "02" Then
''          rstdestino2("H_Cta_Aux1") = VAR_CTA
''        End If
''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
'      End If
'
'      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
'      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
'        'desafecta un devengado
'        rstdestino2("D_Cuenta") = cta_credito1
'        rstdestino2("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_cred11
'        rstdestino2("D_SubCta2") = Subcta_cred21
'        rstdestino2("D_Aux1") = h_aux1_1
'        rstdestino2("D_Aux2") = h_aux2_1
'        rstdestino2("D_Aux3") = h_aux3_1
''        rstdestino2("D_Cta_Aux1") = "VESCT"
'        Select Case h_aux1_1
'            Case "01"
'                rstdestino2("D_Cta_Aux1") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux1") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux1") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux1") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux1") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux1") = ""
'        End Select
'
'        Select Case h_aux2_1
'            Case "01"
'                rstdestino2("D_Cta_Aux2") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux2") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux2") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux2") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux2") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux2") = ""
'        End Select
'
'        Select Case h_aux3_1
'            Case "01"
'                rstdestino2("D_Cta_Aux3") = VAR_BENEF
'            Case "02"
'                rstdestino2("D_Cta_Aux3") = VAR_CTA
'            Case "03"
'                rstdestino2("D_Cta_Aux3") = VAR_PROY2
'            Case "04"
'                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "07"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux3") = VAR_ORG
'            Case "10"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux3") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux3") = ""
'        End Select
''        If h_aux1_1 = "01" Then
''          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''        End If
''        If h_aux1_1 = "02" Then
''          rstdestino2("D_Cta_Aux1") = VAR_CTA
''        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado
'
'        rstdestino2("H_Cuenta") = cta_deb1
'        rstdestino2("H_Nombre") = d_cta_nombre_1  ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_deb11
'        rstdestino2("H_SubCta2") = Subcta_deb21
'        rstdestino2("H_Aux1") = d_aux1_1
'        rstdestino2("H_Aux2") = d_aux2_1
'        rstdestino2("H_Aux3") = d_aux3_1
''        rstdestino2("H_Cta_Aux1") = "VESCT"
'        Select Case d_aux1_1
'            Case "01"
'                rstdestino2("H_Cta_Aux1") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux1") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux1") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux1") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux1") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux1") = ""
'        End Select
'
'        Select Case d_aux2_1
'            Case "01"
'                rstdestino2("H_Cta_Aux2") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux2") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux2") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux2") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux2") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux2") = ""
'        End Select
'
'        Select Case d_aux3_1
'            Case "01"
'                rstdestino2("H_Cta_Aux3") = VAR_BENEF
'            Case "02"
'                rstdestino2("H_Cta_Aux3") = VAR_CTA
'            Case "03"
'                rstdestino2("H_Cta_Aux3") = VAR_PROY2
'            Case "04"
'                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
'            Case "05"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "07"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux3") = VAR_ORG
'            Case "10"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux3") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux3") = ""
'        End Select
''        If d_aux1_1 = "01" Then
''          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''          'DtCCta_descripcion_larga
''        End If
''        If d_aux1_1 = "02" Then
''          rstdestino2("H_Cta_Aux1") = VAR_CTA
''        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado
'      End If
'
''      '==== INI DVI ====
''      If (VAR_CODTIPO = "DVI") Then
''        rstdestino2("D_Cuenta") = cta_deb1
'''        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
''        rstdestino2("D_Subcta1") = Subcta_deb11
''        rstdestino2("D_SubCta2") = Subcta_deb21
''        rstdestino2("D_Aux1") = d_aux1_1
''        rstdestino2("D_Aux2") = d_aux2_1
''        rstdestino2("D_Aux3") = d_aux3_1
''        If d_aux1_1 = "01" Then
''          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''        End If
''        If d_aux1_1 = "02" Then
''          rstdestino2("D_Cta_Aux1") = VAR_CTA
''        End If
'''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
''        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
''        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
''        rstdestino2("D_Cambio") = GlTipoCambioMercado
''        rstdestino2("H_Cuenta") = cta_credito1
'''        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
''        rstdestino2("H_SubCta1") = Subcta_cred11
''        rstdestino2("H_SubCta2") = Subcta_cred21
''        rstdestino2("H_Aux1") = h_aux1_1
''        rstdestino2("H_Aux2") = h_aux2_1
''        rstdestino2("H_Aux3") = h_aux3_1
''        'rstdestino2("H_Cta_Aux1") = "VESCT"
''        If h_aux1_1 = "01" Then
''          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''          'DtCCta_descripcion_larga
''        End If
''        If h_aux1_1 = "02" Then
''          rstdestino2("H_Cta_Aux1") = VAR_CTA
''        End If
'''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
''        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
''        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
''        rstdestino2("H_Cambio") = GlTipoCambioMercado
''      End If
''      '==== FIN DVI ====
'
'      If yacontabilizo = 0 Then
'        rstdestino2("Usr_codigo") = GlUsuario
'        rstdestino2("Fecha_registro") = Date
'        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
'      End If
'
'      rstdestino2.Update
'      If rstdestino2.State = 1 Then rstdestino2.Close
'      '======= fin registra co_diario ==========
'      rstdestino.MoveNext
'    Next i
'    '======= inI Actualiza campos de estatus de ingresos ==========
''    If rstdestino.State = 1 Then rstdestino.Close
''    rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '" & correlativo1 & "' and org_codigo = '" & VAR_ORG & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
''    rstdestino.MoveFirst
''    If Not (rstdestino.EOF) Then
''      rstdestino("estado_aprobacion") = "S"
''        If VAR_CODTIPO = "DEI" Then
''          rstdestino("estado_devengado") = "S"
''        End If
''        If VAR_CODTIPO = "REC" Then
''          rstdestino("estado_recaudado") = "S"
''        End If
''        If VAR_CODTIPO = "DYR" Then
''          rstdestino("estado_devengado") = "S"
''          rstdestino("estado_recaudado") = "S"
''        End If
''
''        If VAR_CODTIPO = "DES" Then
''          rstdestino("estado_desafectado") = "S"
''        End If
''        If VAR_CODTIPO = "ANI" Then
''          rstdestino("estado_anulado") = "S"
''        End If
''        If VAR_CODTIPO = "DVI" Then
''          rstdestino!estado_desafectado = "S"
''          rstdestino!estado_anulado = "S"
''        End If
''       rstdestino.Update
''       If rstdestino.State = 1 Then rstdestino.Close
''    End If
'    '======= fin Actualiza campos de estatus de ingresos ==========
'    ' AAAAAAAAAQQQQQQQQQQQUUUUUUUUUUUIIIIIIIIIII
'    cod_ant = 0
'    org_ant = ""
'    '======= ini Actualiza el monto recaudado  ==========
'    If (VAR_CODTIPO = "REC") Then
'      '      If rstdestino.State = 1 Then rstdestino.Close
'      '      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'      '        cod_ant = rstdestino("ingreso_codigo_anterior")
'      '        org_ant = rstdestino("org_codigo")
'      '      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + VAR_DOL2
'          rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") + VAR_BS2
'          rstdestino.Update
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'
'    If (VAR_CODTIPO = "DES") Then
''      If rstdestino.State = 1 Then rstdestino.Close
''      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
''      Print VAR_CODANT
''      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
''        cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
''        org_ant = rstdestino("org_codigo")
''      End If
'
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        If rstdestino("codigo_tipo") = "DEI" Then 'And VAR_CODTIPO = "DES"
''          rstdestino!estado_desafectado = "S" 02/07/01
'          rstdestino!estado_codigo = "DES"
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'        Else
'          rstdestino("estado_codigo") = "DES"
''          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
'          cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
'          org_ant = rstdestino("org_codigo")
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'          'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
'          rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
'            rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") - VAR_BS2
'          End If
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'        End If
'      End If
'    End If
'
'    If (VAR_CODTIPO = "ANI") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        If rstdestino("codigo_tipo") = "REC" Then
''          rstdestino("estado_desafectado") = ""
'          rstdestino("estado_codigo") = "ANI"
''          rstdestino("estado_devengado") = "S" 02/07/01
''          rstdestino("estado_anulado") = ""
''          rstdestino("codigo_tipo") = "DEI" 02/07/01
'          rstdestino("monto_recaudado_dolares") = 0
'        End If
'      End If
'      rstdestino.Update
''      Print rstdestino!ingreso_codigo_anterior
''      Print rstdestino!monto_recaudado
'      cod_ant = 0
'      org_ant = ""
'
'      'Call f_actual_rec(rstdestino!org_codigo, rstdestino!ingreso_codigo_anterior)
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'    If (VAR_CODTIPO = "DVI") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        rstdestino!estado_codigo = "DVI"
'      End If
'      rstdestino.Update
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'    '======= fin Actualiza el monto recaudado  ==========
'
'    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
'    If VAR_CODTIPO = "REC" Or VAR_CODTIPO = "DYR" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
'      If Not rstdestino.EOF Then
'        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
'        rstdestino.Update
'      End If
'    End If
'    If VAR_CODTIPO = "ANI" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
'      If Not rstdestino.EOF Then
'        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
'        rstdestino.Update
'      End If
'    End If
'    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
'    'LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
'    'Frmmensaje.Visible = False
'    db.CommitTrans
'  'End If
'  'marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"
'
'End Sub
'
''Private Sub f_actual_rec(org, codant)
''  Dim acumDl As Double
''  Dim rsrecalc As New ADODB.Recordset
''  Set rsrecalc = New ADODB.Recordset
''  If rsrecalc.State = 1 Then rsrecalc.Close
''  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
''  If rsrecalc.RecordCount > 0 Then
''    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
''  Else
''    acumDl = 0
''  End If
''  If rsrecalc.State = 1 Then rsrecalc.Close
''  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
''  If rsrecalc.RecordCount > 0 Then
''    rsrecalc!monto_recaudado_dolares = acumDl
''  End If
''  rsrecalc.Update
''  If rsrecalc.State = 1 Then rsrecalc.Close
''
''End Sub
'
'Private Sub graba_proyecto()
''    Select Case Ado_datos.Recordset!unidad_codigo
''        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
''            VAR_PROY = 12
''        Case "UCOM"
''            VAR_PROY = 17
''        Case "DVTA"
''            VAR_PROY = 18
''
''    End Select
''
''    Set rs_aux1 = New ADODB.Recordset
''    If rs_aux1.State = 1 Then rs_aux1.Close
''    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
''    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux1.RecordCount > 0 Then
''        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
''    Else
''        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
''           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & GlUsuario & "', '" & Date & "')"
''    End If
'End Sub
'
'Private Sub graba_ingreso()
'    '======= Ini grabado de datos
'   'swgraba = 0
'   'Call valida
'
''   If swgraba = 1 Then
''      FraOpciones2.Visible = False
''      fraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
'
'      'If v_añadir = 1 Then
'        'EFECTIVO o a CREDITO
'         'db.BeginTrans
'         Call add_correl
'         Set rstdestino = New ADODB.Recordset
'         rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & Ado_datos16.Recordset!unidad_codigo & "' and solicitud_codigo= " & Ado_datos16.Recordset!solicitud_codigo & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
'         If rstdestino.RecordCount > 0 Then
'            VAR_CODANT = rstdestino!ingreso_codigo
'         End If
'         rstdestino.AddNew
'         rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
'         rstdestino("ingreso_codigo") = correlativo1
'         'VAR_CODANT = correlativo1
'         'CAMBIAR org_codigo
'         rstdestino("org_codigo") = "111"
'         VAR_ORG = "111"
'         'CAMBIAR org_codigo
'         'CAMBIAR COD ingreso_codigo_anterior
'         rstdestino("ingreso_codigo_anterior") = VAR_CODANT
'         'CAMBIAR COD ingreso_codigo_anterior
'         'CAMBIAR DEI O REC
'         rstdestino("Codigo_tipo") = "REC"
'         VAR_CODTIPO = "REC"
'         'CAMBIAR DEI O REC
'         rstdestino("proceso_codigo") = "FIN"
'         rstdestino("subproceso_codigo") = "FIN-01"
'         rstdestino("etapa_codigo") = "FIN-01-02"
'         rstdestino("clasif_codigo") = "ADM"
'         rstdestino("doc_codigo") = "R-110"
'         rstdestino("doc_numero") = correlativo1
'         rstdestino("unidad_codigo") = Ado_datos16.Recordset("unidad_codigo")
'         rstdestino("solicitud_codigo") = Ado_datos16.Recordset("solicitud_codigo")
'         rstdestino("solicitud_tipo") = "3"
'
'         rstdestino("beneficiario_codigo") = Ado_datos.Recordset("beneficiario_codigo")
'         VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
'         rstdestino("fecha_ingreso") = Date
'         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
'         rstdestino("tipo_moneda") = "BOB"
'         VAR_MONEDA = "BOB"
'         rstdestino("ingreso_concepto") = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
'         VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
'         If Ado_datos16.Recordset("venta_tipo") = "E" Then
'            rstdestino("tipo_comp") = "DYR"
'         Else
'            rstdestino("tipo_comp") = "REC"
'         End If
'         'CAMBIAR FTE
'         rstdestino("fte_codigo") = "10"
'         VAR_FTE = "10"
'         'CAMBIAR FTE
'         'CAMBIAR RUBROS
'         rstdestino("rubro_codigo") = "11200"
'         VAR_PARTIDA = "11200"
'         'CAMBIAR RUBROS
'         rstdestino("cheque_o_trf") = "T"
'         rstdestino("Bco_codigo") = "BCP"
'         'CAMBIAR CTA
'         rstdestino("cta_codigo") = IIf(dtc_cta = "", "1111111111", dtc_cta)
'         VAR_CTA = IIf(dtc_cta = "", "1111111111", dtc_cta)
'         'CAMBIAR CTA
'         rstdestino("numero_documento") = "0"
'         rstdestino("unidad_codigo_ant") = Ado_datos16.Recordset("unidad_codigo_ant")
'         VAR_CITE = Ado_datos16.Recordset("unidad_codigo_ant")
'         rstdestino("monto_dolares") = Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
'         VAR_DOL2 = Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
'         rstdestino("monto_bolivianos") = Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
'         VAR_BS2 = Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
'         rstdestino("monto_recaudado_dolares") = Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
'         rstdestino("monto_recaudado_bolivianos") = Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
'         rstdestino("convenio_codigo") = "NN"
'         rstdestino("pro_codigo_det") = Ado_datos16.Recordset("edif_codigo")
'         VAR_PROY2 = Ado_datos16.Recordset("edif_codigo")
'         rstdestino("estado_CODIGO") = "APR"
'         'rstdestino("estado_codigo_dr") = "DEI"
'
'         rstdestino("usr_CODIGO") = GlUsuario
'         rstdestino("fecha_registro") = Date
'         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'         rstdestino.Update
'         If rstdestino.State = 1 Then rstdestino.Close
'        'db.CommitTrans
'
''          If rstIngresos.State = 1 Then rstIngresos.Close
''          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
''          rstIngresos.Sort = "ingreso_codigo"
''          rstIngresos.Requery
'
''          rstIngresos.Requery
''          Set AdoIngresos.Recordset = rstIngresos
''          AdoIngresos.Refresh
''          AdoIngresos.Recordset.Find "ultimo = 'S'"
''          If Not (AdoIngresos.Recordset.EOF) Then
''            marca1 = AdoIngresos.Recordset.Bookmark
''            AdoIngresos.Recordset("ultimo") = "N"
''            AdoIngresos.Recordset.Update
''          End If
'
''          AdoIngresos.Recordset.Move marca1 - 1
'
''          marca1 = 0
'      'End If
''   Else
''      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
'''      FraOpciones2.Visible = False
'''      FraOpciones.Visible = True
'''      FraIngresosNav.Enabled = True
'''      FraIngresosDat.Enabled = False
'''      AdoIngresos.Refresh
''   End If
''   LblAccion = ""
''AAQQQQQUIIIIIIIIII    JQA
'
'End Sub
'
'Private Sub add_correl()
'  Set rstcorrel_ing = New ADODB.Recordset
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' ", db, adOpenDynamic, adLockOptimistic
'  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
'  If rstcorrel_ing.RecordCount = 0 Then
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = "111"   'Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
'  Else
'     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
'  End If
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'
'End Sub
'
''Private Sub CmdGrabaCobranza()
''    If swnuevo = 1 Then
'''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
'''      Set Ado_datos16.Recordset = rstdestino
'''      Ado_datos16.Recordset.AddNew
''      Ado_datos16.Recordset!correl_venta = Val(lblcorrelVenta.Caption)
''      Ado_datos16.Recordset!venta_codigo = Val(TxtNroVenta.Text)
''      Ado_datos16.Recordset!ges_gestion = Year(Date)    'Trim(LblGestion.Caption)
''    End If
''      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
''      Ado_datos16.Recordset!ci = dtc_codigo4A.Text                                                     'Codigo Cobrador
''      Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text + " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
''      Ado_datos16.Recordset!deuda_cobrada = Val(TxtMonto.Text)                                  'Monto Cobrado
''      Ado_datos16.Recordset!deuda_cobrada_dol = Val(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
''      Ado_datos16.Recordset!fecha_cobranza = DTPFechaCobro.Value                                'Fecha de Cobranza
''      'Call acumulaMont(Ado_datos16.Recordset!ges_gestion, Ado_datos16.Recordset!correl_venta, Ado_datos16.Recordset!venta_codigo)
''      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
''
''      Ado_datos16.Recordset!obs_cobranza = TxtObs
''      Ado_datos16.Recordset!nro_cmpbte = Trim(TxtCmpbte)
''      Ado_datos16.Recordset!usr_usuario = GlUsuario
''      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
''      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
''      Ado_datos16.Recordset.Update
''End Sub
'
''Private Sub CmdModDetalle_Click()
''  FraDetalle.Visible = True
''  FraDetalle.Enabled = True
''  txtnosolicitud1.Enabled = False
''  txtcorrdet.Enabled = False
''  dtccodpar.SetFocus
''  CmdGraDetalle.Enabled = True
''  CmdAddDetalle.Enabled = False
''  CmdModDetalle.Enabled = False
''  CmdSalDetalle.Enabled = False
''  CmdCanDetalle.Enabled = True
''  swgrabar = 2
''End Sub
'
''Private Sub CmdGraDetalle_Click()
''    If swgrabar = 1 Then
''        Dim rstdestino As New ADODB.Recordset
''        If rstdestino.State = 1 Then rstdestino.Close
''        rstdestino.Open "select * from ao_solicitud_detalle_correl where formulario = '" & "F11" & "' and correl_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
''        If Not (rstdestino.EOF) Then
''            rstdestino("correl_solicitud_detalle") = rstdestino("correl_solicitud_detalle") + 1
''        Else
''            rstdestino.AddNew
''            rstdestino("formulario") = "F11"
''            rstdestino("correl_solicitud") = Ado_datos.Recordset("codigo_solicitud")
''            rstdestino("correl_solicitud_detalle") = 1
''        End If
''        correldetalle = rstdestino("correl_solicitud_detalle")
''        rstdestino.Update
''        If rstdestino.State = 1 Then rstdestino.Close
''        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correlativo_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
''        rstdestino.AddNew
''        rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
''        rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
''        rstdestino("correlativo_detalle") = correldetalle
''        rstdestino("Par_codigo") = dtccodpar.Text
''        rstdestino("Importe_nacional") = txtsolpeso.Text
''        rstdestino("formulario") = "F11"
''        rstdestino.Update
''        If rstdestino.State = 1 Then rstdestino.Close
''        Set rs_datos14 = New ADODB.Recordset
''        If rs_datos14.State = 1 Then rs_datos14.Close
''        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
''        Set adoDetalleSolicitud.Recordset = rs_datos14
''        adoDetalleSolicitud.Refresh
''    End If
''    If swgrabar = 2 Then
''        If rstdestino.State = 1 Then rstdestino.Close
''        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoDetalleSolicitud.Recordset("ges_gestion") & "' and correlativo_solicitud = " & adoDetalleSolicitud.Recordset("correlativo_solicitud") & " and correlativo_detalle =" & adoDetalleSolicitud.Recordset("correlativo_detalle"), db, adOpenDynamic, adLockOptimistic
''        If Not (rstdestino.EOF) Then
''            rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
''            rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
''            rstdestino("correlativo_detalle") = correldetalle
''            rstdestino("Par_codigo") = dtccodpar.Text
''            rstdestino("Importe_nacional") = txtsolpeso.Text
''            rstdestino("formulario") = "F11"
''            rstdestino.Update
''        End If
''        If rstdestino.State = 1 Then rstdestino.Close
''        Set rs_datos14 = New ADODB.Recordset
''        If rs_datos14.State = 1 Then rs_datos14.Close
''        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
''        Set adoDetalleSolicitud.Recordset = rs_datos14
''        adoDetalleSolicitud.Refresh
''    End If
''    CmdGraDetalle.Enabled = False
''    CmdAddDetalle.Enabled = True
''    CmdModDetalle.Enabled = True
''    CmdSalDetalle.Enabled = True
''    CmdCanDetalle.Enabled = False
''    FraDetalle.Enabled = False
''    swgrabar = 0
''End Sub
'
'Private Sub CmdNOunidad_Click()
'    swunidad = 0
'    Frmunidad.Visible = False
'End Sub
'
'Private Sub CmdOKunidad_Click()
'    swunidad = 1
'        If swunidad = 1 Then
'            Dim rstpagos As New ADODB.Recordset
'            Set rstpagos = New ADODB.Recordset
'            If rstpagos.State = 1 Then rstpagos.Close
'            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
'            rstpagos.AddNew
'                rstpagos("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
'                rstpagos("codigo_pago") = "" 'genera jorge
'                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'                rstpagos("formulario") = Ado_datos.Recordset("formulario")
'                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
'                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
'                rstpagos("estado_compromiso") = "N"
'                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
'             rstpagos.Update
'        End If
'End Sub
'
'Private Sub CmdGrabaCobro_Click()
'End Sub
'
''Private Sub CmdGrabaDet_Click()
'''If dtc_desc12 = "" Then
'''    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'''    Exit Sub
'''  End If
''  If dtc_codigo15 = "" Then
''     MsgBox "Debe Elejir un Producto para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'''  If dtc_desc13 = "" Then
'''    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'''    Exit Sub
'''  End If
''    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
''    '    VAR_PARTIDA = "OK"
''    If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Then
''          'fraOpciones.Visible = True
''          'FraGrabarCancelar.Visible = False
''          'TxtNroVenta.Enabled = True
''          FrmEdita.Enabled = False
''        '  DtGListaN.Enabled = True
''          'cmdElige.Enabled = False
''        '  dtc_codigo15.Visible = False
''        '  dtc_desc15.Visible = False
''          'txt_descripcion_venta.Enabled = False
''        If swnuevo = 1 Then
''          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
''          ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
''          ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
''        End If
''          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
''          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
''          ado_datos14.Recordset!bien_codigo = Trim(dtc_codigo15.Text)                       'Codigo Bien (Equipo, Producto, etc)
''          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
''          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
''          ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
''          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
''          ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
''          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
''          If TxtCantidad.Text = "" Then
''            TxtCantidad.Text = "1"
''          End If
''          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
''          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
''          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
''          If CDbl(TxtDescuento) > 0 Then
''            ado_datos14.Recordset!venta_descuento_bs = CDbl(TxtDescuento.Text)      'Dcto por producto CON DESCUENTO
''            ado_datos14.Recordset!venta_descuento_dol = Val(TxtDescuento) / GlTipoCambioMercado
''          Else
''            'ado_datos14.Recordset!descuento_venta = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) * (CDbl(Dtc_aux12)) 'Dcto por producto DE LA TABLA
''            TxtDescuento.Text = "0"
''            ado_datos14.Recordset!venta_descuento_bs = 0
''            ado_datos14.Recordset!venta_descuento_dol = 0
''          End If
''          ado_datos14.Recordset!venta_precio_total_bs = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) - (CDbl(TxtDescuento)) 'Precio Total Producto
''          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
''          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
''          ado_datos14.Recordset!venta_precio_total_dol = (ado_datos14.Recordset!venta_precio_total_bs) / GlTipoCambioMercado
''          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
''          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
''          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
''          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
''          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
''          If OpMod1.Value = True Then
''            ado_datos14.Recordset!modelo_elegido = "S"
''          Else
''            ado_datos14.Recordset!modelo_elegido = "N"
''          End If
''          If OpMod2.Value = True Then
''            ado_datos14.Recordset!modelo_elegido_h = "S"
''          Else
''            ado_datos14.Recordset!modelo_elegido_h = "N"
''          End If
''          If OpMod2.Value = True Then
''            ado_datos14.Recordset!modelo_elegido_x = "S"
''          Else
''            ado_datos14.Recordset!modelo_elegido_x = "N"
''          End If
''          ado_datos14.Recordset!estado_codigo = "REG"
''          ado_datos14.Recordset!usr_codigo = GlUsuario
''          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
''          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
''          ado_datos14.Recordset.Update
''        'db.CommitTrans
''
''        'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
''        Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
''        SSTab1.Tab = 0
''        SSTab1.TabEnabled(0) = True
''        SSTab1.TabEnabled(1) = False
''        SSTab1.TabEnabled(2) = False
''        FraNavega.Enabled = True
''        FrmDetalle.Enabled = True
''        'FrmDetalle.Visible = True
''        FrmCobranza.Visible = True
''        FrmABMDet.Visible = True
''        FrmABMDet2.Visible = True
''        Call OptFilGral1_Click
''        If swnuevo = 1 Then
''          'Call abre_ventas_det
''          'rs_datos14.Requery
''          'ado_datos14.Refresh
''          'ado_datos14.Recordset.MoveLast
''
''        End If
''        swnuevo = 0
''    Else
''        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
''    End If
''  'Else
''  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
''  'End If
''End Sub
'
'Private Sub BtnImprimir2_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-103_recibo_cobranza.rpt"
'    CryR01.WindowShowRefreshBtn = True
'    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'End Sub
'
'Private Sub BtnModDetalle_Click()
'  If Ado_datos16.Recordset.RecordCount > 0 Then
'
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = True
''    BtnImprimir2.Visible = False
''    BtnImprimir3.Visible = False
'  Else
'    MsgBox "No existen datos de la Venta, Verifique por favor !! ", vbExclamation, "Atención!"
'  End If
'End Sub
'
'Private Sub BtnSalir2_Click()
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
'End Sub
'
'Private Sub BtnSalir3_Click()
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmEdita.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
'End Sub
'
'
'Private Sub dtc_aux2_Click(Area As Integer)
'    dtc_codigo2.BoundText = Dtc_aux2.BoundText
'    dtc_desc2.BoundText = Dtc_aux2.BoundText
'    Dtc_deudor2.BoundText = Dtc_aux2.BoundText
'End Sub
'
'Private Sub dtc_aux3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_aux3.BoundText
'    dtc_desc3.BoundText = dtc_aux3.BoundText
'End Sub
'
'Private Sub dtc_aux4_Click(Area As Integer)
'    dtc_codigo4.BoundText = dtc_aux4.BoundText
'    dtc_desc4.BoundText = dtc_aux4.BoundText
'End Sub
'
'Private Sub dtc_codigo1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'End Sub
'
'Private Sub dtc_codigo2_Click(Area As Integer)
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'    Dtc_aux2.BoundText = dtc_codigo2.BoundText
'    Dtc_deudor2.BoundText = dtc_codigo2.BoundText
'End Sub
'
'Private Sub dtc_codigo3_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'    dtc_aux3.BoundText = dtc_codigo3.BoundText
'End Sub
'
'Private Sub dtc_codigo4_Click(Area As Integer)
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'    dtc_aux4.BoundText = dtc_codigo4.BoundText
'End Sub
'
'Private Sub dtc_cta_Click(Area As Integer)
'    dtc_ctades.BoundText = dtc_cta.BoundText
'End Sub
'
'Private Sub dtc_ctades_Click(Area As Integer)
'    dtc_cta.BoundText = dtc_ctades.BoundText
'End Sub
'
'Private Sub dtc_desc1_Click(Area As Integer)
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
'End Sub
'
'Private Sub dtc_desc2_Click(Area As Integer)
'    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    Dtc_aux2.BoundText = dtc_desc2.BoundText
'    Dtc_deudor2.BoundText = dtc_desc2.BoundText
'End Sub
'
'Private Sub dtc_desc3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_desc3.BoundText
'    dtc_aux3.BoundText = dtc_desc3.BoundText
'End Sub
'
'Private Sub dtc_desc4_Click(Area As Integer)
'    dtc_codigo4.BoundText = dtc_desc4.BoundText
'    dtc_aux4.BoundText = dtc_desc4.BoundText
'End Sub
'
'Private Sub Dtc_deudor2_Click(Area As Integer)
'    dtc_codigo2.BoundText = Dtc_deudor2.BoundText
'    Dtc_aux2.BoundText = Dtc_deudor2.BoundText
'    dtc_desc2.BoundText = Dtc_deudor2.BoundText
'End Sub
'
'Private Sub dtc_codigo13_Click(Area As Integer)
'    dtc_desc13.BoundText = dtc_codigo13.BoundText
'    Dtc_Stock13.BoundText = dtc_codigo13.BoundText
'End Sub
'
'Private Sub dtc_desc13_Click(Area As Integer)
'    dtc_codigo13.BoundText = dtc_desc13.BoundText
'    Dtc_Stock13.BoundText = dtc_desc13.BoundText
'End Sub
'
'Private Sub dtc_codigo2A_Click(Area As Integer)
'    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
'End Sub
'
'Private Sub dtc_codigo4A_Click(Area As Integer)
'    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
'End Sub
'
'Private Sub DataCombo1_Click(Area As Integer)
'    DataCombo2.Text = DataCombo1.BoundText
'End Sub
'
'Private Sub DataCombo2_Click(Area As Integer)
'    DataCombo1.Text = DataCombo2.BoundText
'End Sub
'
'Private Sub cmdVerifica_existencia_Click()
'' verifica existencia  del almacen
'Cant_Alm = 0
'AlFrmExistencia_Almacen.Show
'
'DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
'Txtcant_alm = Cant_Alm
'If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub
'
'Private Sub dtc_codigo11_Click(Area As Integer)
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'End Sub
'
'Private Sub dtc_desc11_Click(Area As Integer)
'    dtc_codigo11.BoundText = dtc_desc11.BoundText
'End Sub
'
'Private Sub dtc_desc11_LostFocus()
'    If dtc_codigo11.Text = "C" Then
'        'TxtCobrado.Visible = False
'        'Label7.Visible = False
'        TxtConcepto.Text = "VENTA AL CREDITO - " + Txt_campo2.Caption
'        TxtPlazo.Visible = True
'    Else
'        If dtc_codigo11.Text = "C" Then
'            TxtConcepto.Text = "VENTA AL CONTADO - " + Txt_campo2.Caption
'            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
'        Else
'        'dtc_codigo2.Text = "VD"
'        'dtc_desc2.Text = "VENTA DIRECTA"
'        'TxtCobrado.Visible = True
'        'Label7.Visible = True
'            TxtConcepto.Text = "VENTA DIRECTA AL CLIENTE"
'            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
'        End If
'    End If
'End Sub
'
'Private Sub dtccodmanejo_Click(Area As Integer)
'    DtCCodigo.BoundText = dtccodmanejo.BoundText
'    DtCDescripcion.BoundText = dtccodmanejo.BoundText
'    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
'    dtccodpeso.BoundText = dtccodmanejo.BoundText
'End Sub
'
'Private Sub dtccodpeso_Click(Area As Integer)
'    DtCCodigo.BoundText = dtccodpeso.BoundText
'    DtCDescripcion.BoundText = dtccodpeso.BoundText
'    dtcunidadmedida.BoundText = dtccodpeso.BoundText
'    dtccodmanejo.BoundText = dtccodpeso.BoundText
'End Sub
'
'Private Sub dtc_codigo15_Click(Area As Integer)
'    dtc_desc15.BoundText = dtc_codigo15.BoundText
'    dtc_unimed15.BoundText = dtc_codigo15.BoundText
'    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
'    dtc_grupo15.BoundText = dtc_codigo15.BoundText
'    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
'    Dtc_partida15.BoundText = dtc_codigo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
'End Sub
'
'Private Sub dtccodpar_Click(Area As Integer)
'    dtcdescripar.Text = dtccodpar.BoundText
'End Sub
'
'Private Sub dtccodpoa_Click(Area As Integer)
'    dtcdespoa.Text = dtccodpoa.BoundText
'End Sub
'
'Private Sub dtccodpuesto_Click(Area As Integer)
'    dtcdenopuesto.Text = dtccodpuesto.BoundText
'End Sub
'
'Private Sub dtccodtipoid_Click(Area As Integer)
'    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
'End Sub
'
'Private Sub dtccoduni_Click(Area As Integer)
'    dtcdescripuni.Text = dtccoduni.BoundText
'End Sub
'
'Private Sub dtccorrcompromiso_Click(Area As Integer)
'    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
'End Sub
'
'Private Sub dtccorrsol_Click(Area As Integer)
' dtcfechasol.BoundText = dtccorrsol.BoundText
'End Sub
'
'Private Sub dtcdenominacionruc_Click(Area As Integer)
'    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
'End Sub
'
'Private Sub dtcdenopuesto_Click(Area As Integer)
'    dtccodpuesto.Text = dtcdenopuesto.BoundText
'End Sub
'
'Private Sub DtCDescripcion_Click(Area As Integer)
'    DtCCodigo.BoundText = DtCDescripcion.BoundText
'    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
'    dtccodmanejo.BoundText = DtCDescripcion.BoundText
'    dtccodpeso.BoundText = DtCDescripcion.BoundText
'End Sub
'
'Private Sub dtc_precioventabase15_Click(Area As Integer)
'    dtc_desc15.BoundText = dtc_precioventabase15.BoundText
'    dtc_unimed15.BoundText = dtc_precioventabase15.BoundText
'    dtc_stocktotal15.BoundText = dtc_precioventabase15.BoundText
'    dtc_grupo15.BoundText = dtc_precioventabase15.BoundText
'    dtc_subgrupo15.BoundText = dtc_precioventabase15.BoundText
'    Dtc_partida15.BoundText = dtc_precioventabase15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_precioventabase15.BoundText
'    dtc_codigo15.BoundText = dtc_precioventabase15.BoundText
'    dtc_preciocompra15.BoundText = dtc_precioventabase15.BoundText
'End Sub
'
'Private Sub dtc_subgrupo15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_subgrupo15.BoundText
'    dtc_desc15.BoundText = dtc_subgrupo15.BoundText
'    dtc_unimed15.BoundText = dtc_subgrupo15.BoundText
'    dtc_stocktotal15.BoundText = dtc_subgrupo15.BoundText
'    dtc_grupo15.BoundText = dtc_subgrupo15.BoundText
'    Dtc_partida15.BoundText = dtc_subgrupo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_subgrupo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_subgrupo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_subgrupo15.BoundText
'End Sub
'
'Private Sub dtc_partida15_Click(Area As Integer)
'    dtc_desc15.BoundText = Dtc_partida15.BoundText
'    dtc_unimed15.BoundText = Dtc_partida15.BoundText
'    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
'    dtc_grupo15.BoundText = Dtc_partida15.BoundText
'    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
'    dtc_codigo15.BoundText = Dtc_partida15.BoundText
'    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
'    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
'    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
'End Sub
'
'Private Sub dtc_desc15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_desc15.BoundText
'    dtc_unimed15.BoundText = dtc_desc15.BoundText
'    dtc_stocktotal15.BoundText = dtc_desc15.BoundText
'    dtc_grupo15.BoundText = dtc_desc15.BoundText
'    dtc_subgrupo15.BoundText = dtc_desc15.BoundText
'    Dtc_partida15.BoundText = dtc_desc15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_desc15.BoundText
'    dtc_precioventabase15.BoundText = dtc_desc15.BoundText
'    dtc_preciocompra15.BoundText = dtc_desc15.BoundText
'End Sub
'
'Private Sub dtcdescripar_Click(Area As Integer)
'    dtccodpar.Text = dtcdescripar.BoundText
'End Sub
'
'Private Sub dtcdescripuni_Click(Area As Integer)
'    dtccoduni.Text = dtcdescripuni.BoundText
'End Sub
'
'Private Sub dtcdescrtipoid_Click(Area As Integer)
'    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
'End Sub
'
'Private Sub dtcfechacompromiso_Click(Area As Integer)
'    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
'End Sub
'
'Private Sub dtcfechasol_Click(Area As Integer)
'    dtccorrsol.BoundText = dtcfechasol.BoundText
'End Sub
'
'Private Sub dtcnroruc_Click(Area As Integer)
'    dtcdenominacionruc.Text = dtcnroruc.BoundText
'End Sub
'
'Private Sub dtc_desc2_LostFocus()
'    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.BackColor = &HFF&
'    Else
'        Dtc_deudor2.BackColor = &H80000010
'    End If
'
'End Sub
'
'Private Sub dtc_desc4A_Click(Area As Integer)
'    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
'End Sub
'
'Private Sub dtctipodoc_Click(Area As Integer)
'    dtcdenodoc.Text = dtctipodoc.BoundText
'End Sub
'
'Private Sub dtcunidadmedida_Click(Area As Integer)
'    DtCCodigo.BoundText = dtcunidadmedida.BoundText
'    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
'    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
'    dtccodpeso.BoundText = dtcunidadmedida.BoundText
'End Sub
'
'Private Sub dtcdespoa_Click(Area As Integer)
'    dtccodpoa.Text = dtcdespoa.BoundText
'End Sub
'
''Private Sub Dtcmaternobe_Click(Area As Integer)
''    Dtcpaternobe.BoundText = Dtcmaternobe.BoundText
''    Dtcnombrebe.BoundText = Dtcmaternobe.BoundText
''    dtc_desc4.Text = Dtcpaternobe.BoundText
''End Sub
'
'Private Sub dtc_desc15_LostFocus()
'    txt_descripcion_venta.Text = dtc_desc15.Text
'    TxtDescuento.Text = "0"
'    TxtPrecioU.Text = dtc_precioventabase15.Text
'    Call AbreAlmacen
'End Sub
'
'Private Sub dtc_codigo12_Click(Area As Integer)
'    Dtc_aux12.BoundText = dtc_codigo12.BoundText
'    dtc_desc12.BoundText = dtc_codigo12.BoundText
'End Sub
'
'Private Sub dtc_grupo15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_grupo15.BoundText
'    dtc_desc15.BoundText = dtc_grupo15.BoundText
'    dtc_unimed15.BoundText = dtc_grupo15.BoundText
'    dtc_stocktotal15.BoundText = dtc_grupo15.BoundText
'    dtc_subgrupo15.BoundText = dtc_grupo15.BoundText
'    Dtc_partida15.BoundText = dtc_grupo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_grupo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_grupo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_grupo15.BoundText
'End Sub
'
'Private Sub dtc_stocktotal15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_stocktotal15.BoundText
'    dtc_desc15.BoundText = dtc_stocktotal15.BoundText
'    dtc_unimed15.BoundText = dtc_stocktotal15.BoundText
'    dtc_grupo15.BoundText = dtc_stocktotal15.BoundText
'    dtc_subgrupo15.BoundText = dtc_stocktotal15.BoundText
'    Dtc_partida15.BoundText = dtc_stocktotal15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_stocktotal15.BoundText
'    dtc_precioventabase15.BoundText = dtc_stocktotal15.BoundText
'    dtc_preciocompra15.BoundText = dtc_stocktotal15.BoundText
'End Sub
'
'Private Sub Dtc_aux12_Click(Area As Integer)
'    dtc_codigo12.BoundText = Dtc_aux12.BoundText
'    dtc_desc12.BoundText = Dtc_aux12.BoundText
'End Sub
'
'Private Sub dtc_precioventafinal15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_desc15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_unimed15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_grupo15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_subgrupo15.BoundText = dtc_precioventafinal15.BoundText
'    Dtc_partida15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_stocktotal15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_precioventabase15.BoundText = dtc_precioventafinal15.BoundText
'    dtc_preciocompra15.BoundText = dtc_precioventafinal15.BoundText
'End Sub
'
'Private Sub dtc_preciocompra15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_preciocompra15.BoundText
'    dtc_desc15.BoundText = dtc_preciocompra15.BoundText
'    dtc_unimed15.BoundText = dtc_preciocompra15.BoundText
'    dtc_stocktotal15.BoundText = dtc_preciocompra15.BoundText
'    dtc_grupo15.BoundText = dtc_preciocompra15.BoundText
'    dtc_subgrupo15.BoundText = dtc_preciocompra15.BoundText
'    Dtc_partida15.BoundText = dtc_preciocompra15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_preciocompra15.BoundText
'    dtc_precioventabase15.BoundText = dtc_preciocompra15.BoundText
'End Sub
'
'Private Sub dtc_stock13_Click(Area As Integer)
'    dtc_codigo13.BoundText = Dtc_Stock13.BoundText
'    dtc_desc13.BoundText = Dtc_Stock13.BoundText
'End Sub
'
'Private Sub dtc_desc12_Click(Area As Integer)
'    Dtc_aux12.BoundText = dtc_desc12.BoundText
'    dtc_codigo12.BoundText = dtc_desc12.BoundText
'End Sub
'
'Private Sub dtc_desc12_LostFocus()
''  If GlSistema = "A" Then       'Or GlSistema = "Z"
''    If dtc_codigo12.Text = "10" Then
''        TxtPrecioU.Text = dtc_precioventabase15.Text
''    Else
''        TxtPrecioU.Text = dtc_precioventafinal15.Text
''    End If
''  Else
''    'If lblventa_tipo.Caption = "E" Then
''    '    TxtPrecioU.Text = dtc_precioventafinal15.Text
''    'Else
''    '    TxtPrecioU.Text = dtc_precioventabase15.Text
''    'End If
''    If Val(dtc_codigo12.Text) > 19 Then
''        TxtPrecioU.Text = dtc_precioventafinal15.Text
''    Else
''        TxtPrecioU.Text = dtc_precioventabase15.Text
''    End If
''    If Val(dtc_codigo12.Text) = 100 Then
''        TxtPrecioU.Text = dtc_preciocompra15.Text
''    End If
''    If Val(dtc_codigo12.Text) = 200 Then
''        TxtPrecioU.Text = "0"
''    End If
''  End If
'
'End Sub
'
'Private Sub dtc_unimed15_Click(Area As Integer)
'    dtc_codigo15.BoundText = dtc_unimed15.BoundText
'    dtc_desc15.BoundText = dtc_unimed15.BoundText
'    dtc_stocktotal15.BoundText = dtc_unimed15.BoundText
'    dtc_grupo15.BoundText = dtc_unimed15.BoundText
'    dtc_subgrupo15.BoundText = dtc_unimed15.BoundText
'    Dtc_partida15.BoundText = dtc_unimed15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_unimed15.BoundText
'    dtc_precioventabase15.BoundText = dtc_unimed15.BoundText
'    dtc_preciocompra15.BoundText = dtc_unimed15.BoundText
'End Sub
'
'Private Sub dtc_desc2A_Click(Area As Integer)
'    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
'End Sub
'
''Private Sub DTPfechasol_Change()
''    txtGes_gestion = CStr(Year(DTPfechasol.Value))
''End Sub
'
'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
'    Ado_datos4.Refresh
'End Sub
'
'Private Sub Form_Load()
'    swnuevo = 0
'    VAR_SW = ""
'    'parametro = "estado_codigo" + " = " + "'REG'"
'    '
'    Call ABRIR_TABLAS_AUX
'    Call OptFilGral1_Click
'    'Call ABRIR_TABLA
'    'Call ABRIR_TABLA_AUX2
'    'Call ABRIR_TABLA_DET3
'    'txt_codigo.Enabled = True
'    mbDataChanged = False
'    FrmCabecera.Enabled = False
'    FrmCobros.Enabled = False
'    dg_datos.Enabled = True
'    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'    GlNombFor = "F04"
'    'LblUsuario.Caption = GlUsuario
'    marca1 = 1
'    deta2 = 0
'    BtnImprimir2.Visible = True
'    BtnImprimir3.Visible = True
'
''    FrmEdita.Enabled = False
''    Cmd_Cliente.Visible = False
'    swnuevo = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
'	Call SeguridadSet(Me)
End Sub
'
'Private Sub ABRIR_TABLAS_AUX()
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
'    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
'    Set Ado_datos1.Recordset = rs_datos1
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
'
'    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
'    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
'    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
'
'    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
'    If rs_datos3.State = 1 Then rs_datos3.Close
'    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
'    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
'    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'
'    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario - Vendedor
'    If rs_datos4.State = 1 Then rs_datos4.Close
'    rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
'    Set Ado_datos4.Recordset = rs_datos4
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
'
'    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
'    If rs_datos4A.State = 1 Then rs_datos4A.Close
'    rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
'    Set ado_datos4A.Recordset = rs_datos4A
'    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
'
''    Set rs_datos5 = New ADODB.Recordset
''    If rs_datos5.State = 1 Then rs_datos5.Close
''    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
''    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
''    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''
''    Set rs_datos6 = New ADODB.Recordset
''    If rs_datos6.State = 1 Then rs_datos6.Close
''    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
''    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
''    Set Ado_datos6.Recordset = rs_datos6
''    dtc_desc6.BoundText = dtc_codigo6.BoundText
''
''    Set rs_datos7 = New ADODB.Recordset
''    If rs_datos7.State = 1 Then rs_datos7.Close
''    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
''    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
''    Set Ado_datos7.Recordset = rs_datos7
''    dtc_desc7.BoundText = dtc_codigo7.BoundText
''
''    Set rs_datos8 = New ADODB.Recordset
''    If rs_datos8.State = 1 Then rs_datos8.Close
''    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
''    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
''    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
''
''    Set rs_datos9 = New ADODB.Recordset
''    If rs_datos9.State = 1 Then rs_datos9.Close
''    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
''    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
''    Set Ado_datos9.Recordset = rs_datos9
''    dtc_desc9.BoundText = dtc_codigo9.BoundText
''
''    Set rs_datos10 = New ADODB.Recordset
''    If rs_datos10.State = 1 Then rs_datos10.Close
''    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
''    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
''    Set Ado_datos10.Recordset = rs_datos10
''    dtc_desc10.BoundText = dtc_codigo10.BoundText
'
'    Set rs_datos11 = New ADODB.Recordset
'    If rs_datos11.State = 1 Then rs_datos11.Close
'    rs_datos11.Open "ac_tipo_compra_venta", db, adOpenStatic
'    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'
'    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh
'
'    'Solo para Equipos (*)
'    Set rs_datos15 = New ADODB.Recordset
'    If rs_datos15.State = 1 Then rs_datos15.Close
'    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
'    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
'    Set ado_datos15.Recordset = rs_datos15
'    ado_datos15.Refresh
'
'   'wwwwwwwwwwwwwwwwwwww
'    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
'    'Call ABREVENTAS
'
''    Set rs_Dsctos = New ADODB.Recordset
''    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
''    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
''    Set AdoDsctos.Recordset = rs_Dsctos
''    AdoDsctos.Refresh
'
'    Set rs_datos17 = New ADODB.Recordset
'    If rs_datos17.State = 1 Then rs_datos17.Close
'    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
'    Set ado_datos17.Recordset = rs_datos17
'    ado_datos17.Refresh
'
'    Set rs_datos20 = New ADODB.Recordset
'    If rs_datos20.State = 1 Then rs_datos20.Close
'    rs_datos20.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
'    Set Ado_datos20.Recordset = rs_datos20
'    dtc_ctades.BoundText = dtc_cta.BoundText
'
'End Sub
'
'Private Sub grabar()
''  'db.BeginTrans
''    If swgrabar = 1 Then
'''      Dim rstdestino As New ADODB.Recordset
'''      Set rstdestino = New ADODB.Recordset
'''      If rstdestino.State = 1 Then rstdestino.Close
'''      rstdestino.Open "select tipo_tramite, numero_correlativo from fc_correl WHERE tipo_tramite='ventas'", db, adOpenDynamic, adLockOptimistic
'''      If rstdestino.RecordCount <> 0 Then
'''        Ado_datos.Recordset("venta_codigo") = (CDbl(rstdestino!numero_correlativo) + 1)
'''        rstdestino!numero_correlativo = (CDbl(rstdestino!numero_correlativo) + 1)
'''        rstdestino.Update
'''      Else
'''        Ado_datos.Recordset("venta_codigo") = 1
'''      End If
'''      If rstdestino.State = 1 Then rstdestino.Close
'''      'Ado_datos.Recordset("venta_codigo") = Ado_datos.Recordset.RecordCount
'''      'rstdestino.AddNew
''    End If
''       Ado_datos.Recordset("ges_gestion") = GlGestion       'CStr(Year(DTPfechasol.Value))
''       Ado_datos.Recordset("unidad_codigo") = dtc_codigo1.Text
''       Ado_datos.Recordset("solicitud_codigo") = txt_codigo.Caption
''       Ado_datos.Recordset("edif_codigo") = dtc_codigo3.Text
''
''       Ado_datos.Recordset("venta_fecha") = DTPfechasol
''       Ado_datos.Recordset("venta_tipo") = dtc_codigo11.Text                'E=Efectivo, C=Credito
''       Ado_datos.Recordset("beneficiario_codigo") = dtc_codigo2.Text        'CLIENTE
''       Ado_datos.Recordset("beneficiario_codigo_resp") = dtc_codigo4.Text   'Vendedor
''       Ado_datos.Recordset("venta_descripcion") = TxtConcepto.Text
''       Ado_datos.Recordset("venta_plazo_dias_calendario") = IIf(TxtPlazo.Text = "", "0", TxtPlazo.Text)
''       Ado_datos.Recordset("venta_tipo_cambio") = GlTipoCambioMercado        'Val(txtTDC.Text)
''        'GlTipoCambioOficial As Currency        'GlTipoCambioMercado As Currency        'GlTipoCambioGestion As Currency
''       Ado_datos.Recordset("tipoben_codigo") = IIf(Dtc_aux2.Text = "", "2", Dtc_aux2.Text)      'Tipo de Beneficiario
''
''       Ado_datos.Recordset("proceso_codigo") = "COM"
''       Ado_datos.Recordset("subproceso_codigo") = "COM-02"
''       Ado_datos.Recordset("etapa_codigo") = "COM-02-01"
''       Ado_datos.Recordset("clasif_codigo") = "COM"
''       Ado_datos.Recordset("doc_codigo") = "R-223"
''       Ado_datos.Recordset("doc_numero") = "0"
''       Ado_datos.Recordset("poa_codigo") = "3.1.2"
''
'''       'If Ado_datos.Recordset("venta_tipo") = "E" Then
'''       '     Ado_datos.Recordset("monto_cobrado") = IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'''       '     Ado_datos.Recordset("deuda_cobrada") = IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'''       '  Else
'''            Ado_datos.Recordset("monto_cobrado") = "0"
'''            Ado_datos.Recordset("deuda_cobrada") = "0"
'''       'End If
'''       If swgrabar = 1 Then
'''         Ado_datos.Recordset("cantidad_total_vendida") = 0
'''         Ado_datos.Recordset("monto_total_bS") = 0  'IIf(TxtCobrado.Text = "", "0", TxtCobrado.Text)
'''         Ado_datos.Recordset("monto_total_Us") = 0
'''       End If
'''       Ado_datos.Recordset("saldo_p_cobrar") = Ado_datos.Recordset("monto_total_bS") - Ado_datos.Recordset("deuda_cobrada")
''
''       Ado_datos.Recordset("estado_codigo") = "REG"
''
''       Ado_datos.Recordset("usr_codigo") = GlUsuario
''       Ado_datos.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
''       Ado_datos.Recordset("hora_registro") = Format(Time, "hh/mm/ss")
''       'Ado_datos.Recordset("usuario_aprueba") = ""
''        'Ado_datos.Recordset("fecha_aprueba") = ""
''
''    Ado_datos.Recordset.Update
''
''    'Ado_datos.Recordset.Requery
''    'If rstdestino.State = 1 Then rstdestino.Close
''    'db.CommitTrans
''    If Ado_datos.Recordset.RecordCount > 0 Then
''       marca1 = Ado_datos.Recordset.Bookmark
''       If Ado_datos.Recordset("venta_tipo") = "E" Then
''           db.Execute "INSERT INTO ao_ventas_cobranzas (venta_codigo, ges_gestion, beneficiario_codigo, beneficiario_codigo_resp, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, literal, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, factura_impresa, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
''           "VALUES ('" & Ado_datos.Recordset!venta_codigo & "', '" & Ado_datos.Recordset!ges_gestion & "', '" & Ado_datos.Recordset!beneficiario_codigo & "', '" & Ado_datos.Recordset!beneficiario_codigo_resp & "', " & Ado_datos.Recordset!venta_monto_total_bs & ", '" & Ado_datos.Recordset!venta_monto_total_dol & "', '0', '0', " & Ado_datos.Recordset!venta_monto_total_bs & ", " & Ado_datos.Recordset!venta_monto_total_dol & ", '" & Date & "', '" & Date & "', 'CANCELADO', 'CERO', 'COM', 'COM-02', 'COM-02-02', 'ADM', 'R-103', '0', 'R-101', '0', '0', 'N', '3.1.2', 'REG', '" & GlUsuario & "', '" & Date & "', '09:00')"
''           '  cobranza_codigo       'Especif. de Identidad
''       End If
''       Call OptFilGral1_Click
''       'Ado_datos.Refresh
''       'Ado_datos.Recordset.Move marca1 - 1
''        If swgrabar = 1 Then
''            Ado_datos.Refresh
''            Ado_datos.Recordset.MoveLast
''        End If
''    End If
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''  If glPersNew = "P" Then
''    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
''    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
''    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
''    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
''    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
''  End If
''  If glPersNew = "L" Then
''    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
''    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
''    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
''    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
''  End If
''  If glPersNew = "PL" Then
''    frmeo_Larvas_mosquitos.Dtc_pers_id = rs_Personal!pers_doc_id
''    frmeo_Larvas_mosquitos.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
''    frmeo_Larvas_mosquitos.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
''    frmeo_Larvas_mosquitos.Dtc_Pers_nombre = rs_Personal!pers_nombres
''  End If
''  If glPersNew = "PMA" Then
''    frmeo_mosquito_adulto.Dtc_pers_id = rs_Personal!pers_doc_id
''    frmeo_mosquito_adulto.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
''    frmeo_mosquito_adulto.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
''    frmeo_mosquito_adulto.Dtc_Pers_nombre = rs_Personal!pers_nombres
''  End If
''  glPersNew = "N"
'
'End Sub
'
'Private Sub OpMod1_Click()
''    Txt_modelo.Text = Txt_modelo1.Text
''    Set rs_datos18 = New ADODB.Recordset
''    If rs_datos18.State = 1 Then rs_datos18.Close
''    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
''    If rs_datos18.RecordCount > 0 Then
''        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
''    End If
''    'Set ado_datos17.Recordset = rs_datos18
''    'ado_datos17.Refresh
'End Sub
'
'Private Sub OpMod2_Click()
''    Txt_modelo.Text = Txt_modelo2.Text
''    Set rs_datos18 = New ADODB.Recordset
''    If rs_datos18.State = 1 Then rs_datos18.Close
''    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
''    If rs_datos18.RecordCount > 0 Then
''        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_h
''    End If
'End Sub
'
'Private Sub OpMod3_Click()
''    Txt_modelo.Text = Txt_modelo3.Text
''    Set rs_datos18 = New ADODB.Recordset
''    If rs_datos18.State = 1 Then rs_datos18.Close
''    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
''    If rs_datos18.RecordCount > 0 Then
''        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_x
''    End If
'End Sub
'
'Private Sub OptFilGral1_Click()
'  '===== Proceso para filtrado general de datos(registros no aprobados)
''   Set rs_datos = New ADODB.Recordset
''   If rs_datos.State = 1 Then rs_datos.Close
''   queryinicial = "select * from ao_ventas_cabecera where estado_codigo = 'REG' "
''   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
''   Set Ado_datos.Recordset = rs_datos
''   Ado_datos.Recordset.Requery
''   If Ado_datos.Recordset.RecordCount > 0 Then
''      Ado_datos.Recordset.Move marca1 - 1
''      'Ado_datos.Recordset.MoveLast
''      Set dg_datos.DataSource = rs_datos
''   End If
'
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select * From av_ventas_cabecera WHERE estado_codigo = 'REG' "
'    queryinicial = "select * From ao_ventas_cobranza WHERE estado_codigo = 'REG' "
'    'queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub
'
'Private Sub OptFilGral2_Click()
'  '===== Proceso para filtrado general de datos (todos los registros )
''  Set rs_datos = New ADODB.Recordset
''  If rs_datos.State = 1 Then rs_datos.Close
''  queryinicial = "select * from ao_ventas_cabecera "
''   rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
''   Set Ado_datos.Recordset = rs_datos
''   Ado_datos.Recordset.Requery
''   If Ado_datos.Recordset.RecordCount > 0 Then
''      Ado_datos.Recordset.Move marca1 - 1
''      'Ado_datos.Recordset.MoveLast
''      Set dg_datos.DataSource = rs_datos
''   End If
'
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    queryinicial = "select * From ao_ventas_cobranza "
'    'queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub
'
''Private Sub Option1_Click()
''    Fra_Total.Visible = True
''End Sub
''
''Private Sub Option2_Click()
''    FrmCobranza.Visible = True
''End Sub
'
'Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
' If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
'       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares_contra.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
''solo numeros y , .
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
'       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
'    Else
'       Txtmonto_dolares.Text = 0
'    End If
'  End If
'
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
'      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos_contra.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
'      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
'    Else
'      TxtMonto_bolivianos.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
'    'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
''solo numeros y , .
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub txtterref_KeyPress(KeyAscii As Integer)
'    If KeyAscii < 58 And KeyAscii > 47 Then
'        KeyAscii = Asc(UCase(Chr(0)))
'    Else
'        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Else
'            KeyAscii = Asc(UCase(Chr(0)))
'            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
'        End If
'    End If
'End Sub
'
'Private Sub cerea()
'  txt_venta = " "
'  dtc_codigo4.Text = " "
'  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
''  dtcmaternosol.Text = " "
''  dtcnombresol.Text = " "
'  txtCantTotal = "0"
'  TxtMontoBs = "0"
'  TxtMontoUs = "0"
'  TxtConcepto = ""
'  dtc_codigo2 = ""
'  dtc_desc2 = ""
'  txtTDC.Text = GlTipoCambioOficial
'
''  DtCDenominacion_moneda = ""
''  TxtMonto_bolivianos = 0
''  Txtmonto_dolares = 0
''  TxtMonto_bolivianos_contra = 0
''  Txtmonto_dolares_contra = 0
''  DtCOrg_descripcion = ""
''  txtjustifica = ""
''  txt_venta = ""
''  txtterref = ""
'End Sub
''Private Sub fbuscaunidad()
''  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
''  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
''  If rstFc_unidad_ejecutora.RecordCount > 0 Then
''    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
''  Else
''    LblUni_descripcion_larga.Caption = ""
''  End If
''  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
''End Sub
'
'Sub creaVista()
'db.Execute "drop view vwF04"
'
'db.Execute "create view vwF04 as " & _
'            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
'            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
'            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
'            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
'            "from ao_solicitud_lista  ,     " & _
'                 "ao_solicitud       ,     " & _
'                 "fc_unidad_ejecutora,     " & _
'                 "rc_personal,             " & _
'                 "ac_bienes                " & _
'            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
'                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
'                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
'                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
'                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
'                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
'                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
'                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
'            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
'            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
'            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "
'
''            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
'
''''db.Execute "create view vwF05 as " & _
''''            "select  ao_solicitud_lista.* " & _
''''            "from ao_solicitud_lista"
'End Sub
'
'Sub CREAVISTAF11()
'db.Execute "drop view VWF11"
'db.Execute "create view VWF11 as " & _
'    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
'    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
'    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
'    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
'    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
'    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
'    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
'    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
'    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
'    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
'"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
'"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
'    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
'    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
'    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
'    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
'    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
'    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
'    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
'    "ao_Solicitud.ci = RC_Personal.ci"

'End Sub
'
'Private Sub acumulaMont(ges, nro)
'  Set rstacumdet = New ADODB.Recordset
'  If rstacumdet.State = 1 Then rstacumdet.Close
'  Set rs_datos19 = New ADODB.Recordset
'  If rs_datos19.State = 1 Then rs_datos19.Close
''  LblGestion
''  lblcorrelVenta
''  lblNroVenta
'  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If IsNull(rstacumdet!totbs) Then
'    VAR_AUX = 0
'    VAR_AUX2 = 0
'    VAR_CANT = 1
'  Else
'    VAR_AUX = Round(rstacumdet!totbs, 2)
'    VAR_AUX2 = Round(rstacumdet!totdl, 2)
'    VAR_CANT = rstacumdet!cantot
'  End If
'
'  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If IsNull(rs_datos19!totbs2) Then
'    Cobrobs = 0
'    VAR_COBR = 0
'  Else
'    Cobrobs = Round(rs_datos19!totbs2, 2)
'    VAR_COBR = Round(rs_datos19!totdl2, 2)
'  End If
'
'  VAR_Bs = VAR_AUX - Cobrobs
'  VAR_Dol = VAR_AUX2 - VAR_COBR
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'
'  TxtMontoBs.Text = VAR_AUX
'  TxtCobrado.Text = Cobrobs
'  TxtBstotal.Text = VAR_Bs
'
''  If IsNull(Ado_datos.Recordset!venta_monto_cobrado_bs) Then
''    Ado_datos.Recordset!venta_monto_cobrado_bs = 0
''    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs
''  Else
''    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs
''  End If
''  If VAR_AUX > 0 Then
''        VAR_AUX2 = VAR_AUX / Ado_datos.Recordset!venta_tipo_cambio
''  Else
''        VAR_AUX2 = 0
''  End If
''  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas_cabecera.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas_cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas_cabecera.saldo_p_cobrar = ao_ventas_cabecera.monto_total_Bs - ao_ventas_cabecera.deuda_cobrada Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
''  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.venta_monto_total_dol = " & rstacumdet!totdl & ", ao_ventas_cabecera.venta_cantidad_total = " & rstacumdet!cantot & ", ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_AUX & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_AUX2 & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
''
''  TxtMontoBs = rstacumdet!totbs
''  TxtCobrado = rs_datos19!totbs2    'IIf(IsNull(Ado_datos.Recordset("venta_monto_cobrado_bs")), 0, Ado_datos.Recordset("venta_monto_cobrado_bs"))
''  If IsNull(Ado_datos.Recordset("venta_saldo_p_cobrar_bs")) Then
''    Text2 = VAR_AUX 'Ado_datos.Recordset("venta_monto_total_bs") - Ado_datos.Recordset("venta_monto_cobrado_bs")
''    Ado_datos.Recordset("venta_saldo_p_cobrar_bs") = VAR_AUX
''  Else
''    Text2 = Ado_datos.Recordset("venta_saldo_p_cobrar_bs")
''  End If
'
'  If rstacumdet.State = 1 Then rstacumdet.Close
'
'  'Print ado_datos14.Recordset!ges_gestion
'  'Print ado_datos14.Recordset!correl_venta
'  'Print ado_datos14.Recordset!venta_codigo
'  'ado_datos14.Recordset!monto_Bolivianos = rstacumdet!totbs
'  'ado_datos14.Recordset!monto_dolares = rstacumdet!totdl
'  'ado_datos14.Recordset.Update
''  Set rstdestino = New ADODB.Recordset
''  If rstdestino.State = 1 Then rstdestino.Close
''  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & ges & "' and correl_venta = '" & corr & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
''  If rstdestino.RecordCount > 0 Then
''    rstdestino!monto_total_Bs = rstacumdet!totbs
''    rstdestino!monto_cobrado = rstacumdet!totbs
''    rstdestino!monto_total_Us = rstacumdet!totdl
''    rstdestino!cantidad_total_vendida = rstacumdet!cantot
''    rstdestino!saldo_p_cobrar = 0
''    rstdestino.Update
''  End If
''  'Set Ado_datos.Recordset = rstdestino
''  If rstdestino.State = 1 Then rstdestino.Close
''  If rstacumdet.State = 1 Then rstacumdet.Close
'End Sub
'
Private Sub sstab1_Click(PreviousTab As Integer)
        If SSTab1.Tab = 0 Then
            'SSTab1.TabEnabled(0) = True
            'SSTab1.TabEnabled(1) = False
        Else
    '        SSTab1.Tab = 0
    '        SSTab1.TabEnabled(0) = True
    '        SSTab1.TabEnabled(1) = False
    '        SSTab1.TabEnabled(2) = False
    '           FrmEditaDet.Visible = False
    '           DtGLista.Visible = False
    '           adoao_solicitud_lista.Visible = False
        End If
End Sub
'
'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub
'
'Private Sub TxtCantidad_LostFocus()
'  If (TxtCantidad.Text) = "" Then
'    TxtCantidad.Text = 1
'  End If
'  If dtc_codigo11.Text = "E" Then
'    If (dtc_codigo12.Text) = "" Or IsNull(dtc_codigo12.Text) Then
'        TxtDescuento.Text = "0"
'    Else
'        TxtDescuento.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) * CDbl(Dtc_aux12.Text))
'    End If
'    'TxtPrecioU.Text = dtc_precioventabase15.Text
'    'TxtTotal.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento.Text))
'  End If
'  If dtc_codigo11.Text = "C" Then
'     TxtDescuento.Text = "0"
'     'TxtDescuento.Text = CDbl(Dtc_aux12) * (CDbl(TxtCantidad) * CDbl(TxtPrecioU))
'     TxtPrecioU.Text = dtc_precioventafinal15.Text
'  End If
'  If (dtc_codigo11.Text <> "E" And dtc_codigo11.Text <> "C") Then
'     TxtDescuento.Text = "0"
'     TxtPrecioU.Text = "0"
'  End If
'  TxtTotal.Text = (CDbl(TxtCantidad.Text) * CDbl(TxtPrecioU.Text)) - CDbl(TxtDescuento.Text)
'
'End Sub
'
'Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub
'
'Private Sub TxtMonto_LostFocus()
'    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
'        TxtMontoDol = "0"
'    Else
'        TxtMontoDol = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
'    End If
'End Sub
'
'Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
'End Sub
Private Sub Form_Load()
'    swnuevo = 0
'    VAR_SW = ""
'    'parametro = "estado_codigo" + " = " + "'REG'"
'    '
'    Call ABRIR_TABLAS_AUX
'    Call OptFilGral1_Click
'    'Call ABRIR_TABLA
'    'Call ABRIR_TABLA_AUX2
'    'Call ABRIR_TABLA_DET3
'    'txt_codigo.Enabled = True
'    mbDataChanged = False
'    FrmCabecera.Enabled = False
'    FrmCobros.Enabled = False
'    dg_datos.Enabled = True
'    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'    GlNombFor = "F04"
'    'LblUsuario.Caption = GlUsuario
'    marca1 = 1
'    deta2 = 0
'    BtnImprimir2.Visible = True
'    BtnImprimir3.Visible = True
'
''    FrmEdita.Enabled = False
''    Cmd_Cliente.Visible = False
'    swnuevo = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
'	Call SeguridadSet(Me)
End Sub

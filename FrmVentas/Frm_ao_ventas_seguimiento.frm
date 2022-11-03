VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_ao_ventas_seguimiento 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Financieros - Ventas - Seguimiento de Cobranzas"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   17115
   Icon            =   "Frm_ao_ventas_seguimiento.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.4749e6
   ScaleMode       =   0  'User
   ScaleWidth      =   1.50646e7
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1560
      Left            =   15240
      ScaleHeight     =   1500
      ScaleWidth      =   1695
      TabIndex        =   187
      Top             =   4080
      Width           =   1755
   End
   Begin VB.CommandButton BtnModDetalle2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Facturadas por Dia"
      Height          =   645
      Left            =   686
      Picture         =   "Frm_ao_ventas_seguimiento.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   180
      ToolTipText     =   "Ver Detalle del Bien ..."
      Top             =   8160
      Width           =   885
   End
   Begin VB.CommandButton BntImprimir2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cobranzas"
      Height          =   645
      Left            =   210
      Picture         =   "Frm_ao_ventas_seguimiento.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   179
      ToolTipText     =   "Imprime Nota de Venta"
      Top             =   8910
      Width           =   885
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Picture         =   "Frm_ao_ventas_seguimiento.frx":19C1
      ScaleHeight     =   1755
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   31
      Top             =   8040
      Width           =   1935
      Begin VB.CommandButton BntImprimir3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobranzas Dolares"
         Height          =   645
         Left            =   960
         Picture         =   "Frm_ao_ventas_seguimiento.frx":6D9F3
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   840
         Width           =   885
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1515
      Left            =   120
      Picture         =   "Frm_ao_ventas_seguimiento.frx":6DFB0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   30
      Top             =   6405
      Width           =   1935
      Begin VB.CommandButton BtnImprimir4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cronogr."
         Height          =   640
         Left            =   600
         Picture         =   "Frm_ao_ventas_seguimiento.frx":D9FE2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprime Cronograma de Cobranzas ..."
         Top             =   75
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kardex"
         Height          =   645
         Left            =   600
         Picture         =   "Frm_ao_ventas_seguimiento.frx":DB764
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   750
         Width           =   765
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6271
      Left            =   120
      TabIndex        =   27
      Top             =   45
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11060
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   0
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "VER INICIO DE COBRANZA"
      TabPicture(0)   =   "Frm_ao_ventas_seguimiento.frx":DBD21
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraNavega1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrmCobros1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraOpciones1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "VER FACTURADOS"
      TabPicture(1)   =   "Frm_ao_ventas_seguimiento.frx":DC733
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmCobros"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraNavega"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraOpciones"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "VER Y APROBAR COBRANZAS"
      TabPicture(2)   =   "Frm_ao_ventas_seguimiento.frx":106355
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOpciones2"
      Tab(2).Control(1)=   "FrmCobros2"
      Tab(2).Control(2)=   "FraNavega2"
      Tab(2).ControlCount=   3
      Begin VB.PictureBox fraOpciones1 
         BackColor       =   &H00404040&
         Height          =   1020
         Left            =   40
         Picture         =   "Frm_ao_ventas_seguimiento.frx":106D67
         ScaleHeight     =   960
         ScaleWidth      =   14880
         TabIndex        =   124
         Top             =   360
         Width           =   14940
         Begin VB.CommandButton BtnBuscar1 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   2760
            Picture         =   "Frm_ao_ventas_seguimiento.frx":172D99
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir1 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   4680
            Picture         =   "Frm_ao_ventas_seguimiento.frx":173351
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Cerrar Ventana"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lbl_titulo1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COBRANZAS"
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
            Left            =   9495
            TabIndex        =   125
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.PictureBox fraOpciones2 
         BackColor       =   &H00404040&
         Height          =   1020
         Left            =   -74960
         Picture         =   "Frm_ao_ventas_seguimiento.frx":17355B
         ScaleHeight     =   960
         ScaleWidth      =   14880
         TabIndex        =   135
         Top             =   360
         Width           =   14940
         Begin VB.CommandButton BtnBuscar2 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   2280
            Picture         =   "Frm_ao_ventas_seguimiento.frx":1DF58D
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir2 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   4680
            Picture         =   "Frm_ao_ventas_seguimiento.frx":1DFB45
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Cerrar Ventana"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lbl_titulo2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COBRANZAS"
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
            Left            =   9480
            TabIndex        =   136
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.Frame FrmCobros2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4780
         Left            =   -69200
         TabIndex        =   144
         Top             =   1440
         Width           =   9135
         Begin VB.TextBox Txt_docnro 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "doc_numero"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7320
            TabIndex        =   188
            Text            =   "0"
            Top             =   3120
            Width           =   1590
         End
         Begin VB.TextBox TxtDscto2D 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_dol2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
            Height          =   285
            Left            =   4755
            TabIndex        =   185
            Text            =   "0"
            Top             =   4360
            Width           =   1275
         End
         Begin VB.TextBox TxtMonto02D 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
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
            Left            =   4755
            TabIndex        =   184
            Text            =   "0"
            Top             =   3740
            Width           =   1275
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6760
            TabIndex        =   178
            Top             =   2610
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFD4F
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2085
            TabIndex        =   13
            Top             =   3120
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "trans_descripcion"
            BoundColumn     =   "trans_codigo"
            Text            =   "Todos"
         End
         Begin VB.ComboBox cmd_moneda2 
            DataField       =   "tipo_moneda2"
            DataSource      =   "Ado_datos02"
            Height          =   315
            ItemData        =   "Frm_ao_ventas_seguimiento.frx":1DFD68
            Left            =   120
            List            =   "Frm_ao_ventas_seguimiento.frx":1DFD75
            TabIndex        =   18
            Text            =   "BOB"
            Top             =   4360
            Width           =   735
         End
         Begin VB.ComboBox cmd_moneda1 
            DataField       =   "tipo_moneda"
            DataSource      =   "Ado_datos02"
            Height          =   315
            ItemData        =   "Frm_ao_ventas_seguimiento.frx":1DFD88
            Left            =   120
            List            =   "Frm_ao_ventas_seguimiento.frx":1DFD95
            TabIndex        =   14
            Text            =   "BOB"
            Top             =   3740
            Width           =   735
         End
         Begin MSDataListLib.DataCombo DataCombo14 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFDA8
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   5085
            TabIndex        =   176
            Top             =   2600
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin VB.TextBox Txt_deposito3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cmpbte_deposito2"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6075
            TabIndex        =   21
            Text            =   "0"
            Top             =   4360
            Width           =   1590
         End
         Begin VB.TextBox TxtDscto2 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_bs2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
            Height          =   285
            Left            =   3405
            TabIndex        =   20
            Text            =   "0"
            Top             =   4360
            Width           =   1275
         End
         Begin VB.TextBox TxtMonto02 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
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
            Left            =   3405
            TabIndex        =   16
            Text            =   "0"
            Top             =   3740
            Width           =   1275
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos02"
            ForeColor       =   &H00FFFFFF&
            Height          =   585
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   147
            Top             =   1140
            Width           =   5900
         End
         Begin VB.TextBox Txt_deposito2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cmpbte_deposito"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6075
            TabIndex        =   17
            Text            =   "0"
            Top             =   3740
            Width           =   1590
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6730
            TabIndex        =   145
            Top             =   1890
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFDC1
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   5535
            TabIndex        =   146
            Top             =   1880
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin MSDataListLib.DataCombo dtc_cta2 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFDDB
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   960
            TabIndex        =   15
            Top             =   3720
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFDF5
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2640
            TabIndex        =   148
            Top             =   3555
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7_2 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE0F
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   915
            TabIndex        =   19
            Top             =   4360
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo9 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE28
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   3480
            TabIndex        =   149
            Top             =   2880
            Visible         =   0   'False
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "trans_codigo"
            BoundColumn     =   "trans_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo DataCombo10 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE41
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2640
            TabIndex        =   150
            Top             =   4080
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro2 
            DataField       =   "cobranza_fecha_cobro1"
            DataSource      =   "Ado_datos02"
            Height          =   300
            Left            =   7740
            TabIndex        =   11
            Top             =   3740
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   84803585
            CurrentDate     =   42043
         End
         Begin MSDataListLib.DataCombo DataCombo12 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE5A
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   1680
            TabIndex        =   151
            Top             =   1880
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo13 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE74
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   120
            TabIndex        =   175
            Top             =   2600
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro02 
            DataField       =   "cobranza_fecha_cobro"
            DataSource      =   "Ado_datos02"
            Height          =   300
            Left            =   7740
            TabIndex        =   12
            Top             =   4360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   84803585
            CurrentDate     =   42043
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   7140
            X2              =   7140
            Y1              =   0
            Y2              =   3030
         End
         Begin VB.Label TxtAutorizacion2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_nro_autorizacion"
            DataSource      =   "Ado_datos02"
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
            TabIndex        =   174
            Top             =   645
            Width           =   1875
         End
         Begin VB.Label TxtCmpbte2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_nro_factura"
            DataSource      =   "Ado_datos02"
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
            Left            =   2700
            TabIndex        =   173
            Top             =   645
            Width           =   1200
         End
         Begin VB.Label Txt_cod_cobro2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_codigo"
            DataSource      =   "Ado_datos02"
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
            Left            =   1635
            TabIndex        =   172
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de...                                                                   NIT/CI"
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
            Left            =   120
            TabIndex        =   171
            Top             =   2320
            Width           =   5625
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            Caption         =   "Código Registro"
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
            Index           =   5
            Left            =   120
            TabIndex        =   170
            Top             =   570
            Width           =   750
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "Tipo de Transacción"
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
            Left            =   120
            TabIndex        =   169
            Top             =   3135
            Width           =   1875
         End
         Begin VB.Label Label52 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Autorización"
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
            Left            =   3960
            TabIndex        =   168
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label Label48 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.de Factura"
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
            Left            =   1980
            TabIndex        =   167
            Top             =   570
            Width           =   765
         End
         Begin VB.Label Label47 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Recibo"
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
            Left            =   7200
            TabIndex        =   166
            Top             =   2775
            Width           =   1785
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrador de CGI:"
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
            Left            =   120
            TabIndex        =   165
            Top             =   1890
            Width           =   1560
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado BOB"
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
            TabIndex        =   164
            Top             =   195
            Width           =   1785
         End
         Begin VB.Label Label41 
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
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   120
            TabIndex        =   163
            Top             =   1275
            Width           =   960
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                     Nro.Venta:                    Nro.Cuota:"
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
            Left            =   120
            TabIndex        =   162
            Top             =   150
            Width           =   6150
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
            DataSource      =   "Ado_datos02"
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
            Left            =   6240
            TabIndex        =   161
            Top             =   135
            Width           =   735
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_codigo_fac"
            DataSource      =   "Ado_datos02"
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
            Left            =   900
            TabIndex        =   160
            Top             =   645
            Width           =   1005
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Facturacion"
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
            TabIndex        =   159
            Top             =   1860
            Width           =   1785
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos02"
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
            Left            =   3960
            TabIndex        =   158
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label DTPFechaProg2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_fecha_fac"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            DataSource      =   "Ado_datos02"
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
            Left            =   7260
            TabIndex        =   157
            Top             =   2145
            Width           =   1710
         End
         Begin VB.Label TxtDsctoTot2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
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
            Left            =   7260
            TabIndex        =   156
            Top             =   465
            Width           =   1710
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda  Cuenta Bancaria 1                  Cobrado1 Bs   Cobrado1 Dol   Cpbte.Deposito1  Fecha.Cobro1"
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
            Left            =   120
            TabIndex        =   155
            Top             =   3480
            Width           =   8880
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7150
            Y1              =   3025
            Y2              =   3025
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda  Cuenta Bancaria 2                  Cobrado2 Bs   Cobrado2 Dol   Cpbte.Deposito2  Fecha.Cobro2"
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
            Left            =   120
            TabIndex        =   154
            Top             =   4095
            Width           =   8880
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturado USD"
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
            TabIndex        =   153
            Top             =   1080
            Width           =   1785
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos02"
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
            Left            =   7260
            TabIndex        =   152
            Top             =   1350
            Width           =   1710
         End
      End
      Begin VB.Frame FraNavega2 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4805
         Left            =   -74960
         TabIndex        =   137
         Top             =   1440
         Width           =   5745
         Begin VB.OptionButton OptFilGral05 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Concluidos"
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
            TabIndex        =   177
            Top             =   4480
            Width           =   1335
         End
         Begin VB.OptionButton OptFilGral04 
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
            Left            =   2280
            TabIndex        =   139
            Top             =   4480
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.OptionButton OptFilGral03 
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
            Left            =   840
            TabIndex        =   138
            Top             =   4480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dg_datos2 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":1DFE8D
            Height          =   4100
            Left            =   75
            TabIndex        =   140
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   7223
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
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "cobranza_fecha_fac"
               Caption         =   "F.Facturacion"
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
            BeginProperty Column02 
               DataField       =   "cobranza_codigo"
               Caption         =   "No.Cobranza"
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
            BeginProperty Column04 
               DataField       =   "cobranza_fecha_cobro"
               Caption         =   "F.Cobranza"
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
               DataField       =   "cobranza_descuento_bs"
               Caption         =   "Tot.Cobr.Bs"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "###,###,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "cobranza_deuda_bs"
               Caption         =   "Cobrado.Bs."
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
            BeginProperty Column08 
               DataField       =   "doc_numero"
               Caption         =   "Nro.Doc.Respaldo"
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
            BeginProperty Column10 
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
            BeginProperty Column11 
               DataField       =   "beneficiario_codigo"
               Caption         =   "NIT/CI del Cliente"
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
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column10 
                  Alignment       =   2
                  ColumnWidth     =   645.165
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Ado_datos02 
            Height          =   330
            Left            =   75
            Top             =   4400
            Width           =   5595
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
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   1020
         Left            =   -74960
         Picture         =   "Frm_ao_ventas_seguimiento.frx":1DFEA7
         ScaleHeight     =   960
         ScaleWidth      =   14880
         TabIndex        =   118
         Top             =   360
         Width           =   14940
         Begin VB.CommandButton BtnVer 
            BackColor       =   &H00808000&
            Caption         =   "Digitaliza"
            Height          =   720
            Left            =   4280
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24BED9
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimir3 
            BackColor       =   &H00C0C000&
            Caption         =   "Factura"
            Height          =   720
            Left            =   2640
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24C31B
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Re-Imprime Factura"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnBuscar 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   1800
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24DA9D
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdFoto 
            BackColor       =   &H00808000&
            Caption         =   "&Reporte"
            Height          =   720
            Left            =   3480
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24E055
            Style           =   1  'Graphical
            TabIndex        =   121
            ToolTipText     =   "Carga Imagen QR"
            Top             =   120
            Visible         =   0   'False
            Width           =   740
         End
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H00808000&
            Caption         =   "Devolver"
            Height          =   720
            Left            =   4280
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24EA57
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Devuelve a Cobradores (Solicitud de Facturación)"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   5115
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24EC61
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cerrar Ventana"
            Top             =   120
            Width           =   765
         End
         Begin VB.Label lbl_titulo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COBRANZAS"
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
            Left            =   9735
            TabIndex        =   123
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.Frame FrmCobros1 
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
         Height          =   4740
         Left            =   5800
         TabIndex        =   81
         Top             =   1440
         Width           =   9135
         Begin VB.ComboBox cmd_fac 
            Height          =   315
            ItemData        =   "Frm_ao_ventas_seguimiento.frx":24EE6B
            Left            =   4320
            List            =   "Frm_ao_ventas_seguimiento.frx":24EE75
            TabIndex        =   182
            Text            =   "FACTURA"
            Top             =   2840
            Width           =   1275
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6730
            TabIndex        =   131
            Top             =   2335
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A1 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EE8A
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   5535
            TabIndex        =   130
            Top             =   2320
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3970
            TabIndex        =   128
            Top             =   4290
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3970
            TabIndex        =   127
            Top             =   3870
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3970
            TabIndex        =   126
            Top             =   3460
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Txt_deposito1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "cmpbte_deposito"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6560
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "0"
            Top             =   3450
            Width           =   2430
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos01"
            ForeColor       =   &H00FFFFFF&
            Height          =   585
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   82
            Top             =   1580
            Width           =   5900
         End
         Begin VB.TextBox TxtMonto1 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
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
            Left            =   5805
            TabIndex        =   1
            Text            =   "0"
            Top             =   3860
            Width           =   1275
         End
         Begin VB.TextBox TxtDscto1 
            Alignment       =   2  'Center
            DataField       =   "cobranza_deuda_dol2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
            Height          =   285
            Left            =   5805
            TabIndex        =   26
            Text            =   "0"
            Top             =   4280
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EEA4
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1800
            TabIndex        =   84
            Top             =   3855
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EEBE
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   2520
            TabIndex        =   85
            Top             =   3675
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EED8
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1800
            TabIndex        =   86
            Top             =   4275
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc61 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EEF1
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1800
            TabIndex        =   87
            Top             =   3450
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "trans_descripcion"
            BoundColumn     =   "trans_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo61 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EF0A
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   2760
            TabIndex        =   88
            Top             =   3240
            Visible         =   0   'False
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "trans_codigo"
            BoundColumn     =   "trans_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo DataCombo11 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EF23
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   2520
            TabIndex        =   89
            Top             =   4080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaSol 
            DataField       =   "cobranza_fecha_sol"
            DataSource      =   "Ado_datos01"
            Height          =   300
            Left            =   7260
            TabIndex        =   0
            Top             =   2820
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
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
            CheckBox        =   -1  'True
            Format          =   84803585
            CurrentDate     =   41678
         End
         Begin MSDataListLib.DataCombo dtc_desc4A1 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EF3C
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1080
            TabIndex        =   129
            Top             =   2320
            Width           =   4800
            _ExtentX        =   8467
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Elige el Documento a Emitir..."
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
            Left            =   1560
            TabIndex        =   181
            Top             =   2850
            Width           =   2595
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7100
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00FFFF80&
            X1              =   7080
            X2              =   7080
            Y1              =   0
            Y2              =   3280
         End
         Begin VB.Label Label58 
            BackColor       =   &H00404040&
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-MMM-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   2085
            TabIndex        =   117
            Top             =   2865
            Visible         =   0   'False
            Width           =   4845
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos01"
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
            Left            =   7140
            TabIndex        =   116
            Top             =   2870
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label55 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Transacción:                                                          Comprobante Deposito:"
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
            Left            =   120
            TabIndex        =   115
            Top             =   3480
            Width           =   6465
         End
         Begin VB.Label Label54 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_deuda_bs2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
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
            Left            =   7720
            TabIndex        =   114
            Top             =   4280
            Width           =   1275
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_programada_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
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
            Left            =   7260
            TabIndex        =   113
            Top             =   1215
            Width           =   1710
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Programado USD"
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
            TabIndex        =   112
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.Bancaria USD:                                                          P/Cobrar USD:                                 BOB"
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
            Left            =   120
            TabIndex        =   111
            Top             =   4305
            Width           =   7530
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   9120
            Y1              =   3285
            Y2              =   3285
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.Bancaria BOB:                                                          P/Cobrar BOB:                                  USD"
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
            Left            =   120
            TabIndex        =   110
            Top             =   3885
            Width           =   7560
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_deuda_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
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
            Left            =   7725
            TabIndex        =   109
            Top             =   3860
            Width           =   1275
         End
         Begin VB.Label TxtDsctoTot1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_programada_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos01"
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
            Left            =   7260
            TabIndex        =   108
            Top             =   465
            Width           =   1710
         End
         Begin VB.Label DTPFechaProg1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            DataSource      =   "Ado_datos01"
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
            Left            =   7260
            TabIndex        =   107
            Top             =   2070
            Width           =   1710
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   3960
            TabIndex        =   106
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Conformidad"
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
            TabIndex        =   105
            Top             =   1815
            Width           =   1785
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_codigo_fac"
            DataSource      =   "Ado_datos01"
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
            TabIndex        =   104
            Top             =   1125
            Width           =   1245
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   6240
            TabIndex        =   103
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                     Nro.Venta:                    Nro.Cuota:"
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
            Left            =   120
            TabIndex        =   102
            Top             =   270
            Width           =   6150
         End
         Begin VB.Label Label24 
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
            Left            =   120
            TabIndex        =   101
            Top             =   1635
            Width           =   960
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Programado BOB"
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
            TabIndex        =   100
            Top             =   195
            Width           =   1785
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrador de CGI:"
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
            Left            =   120
            TabIndex        =   99
            Top             =   2260
            Width           =   960
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitud.Fac"
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
            Left            =   7200
            TabIndex        =   98
            Top             =   2560
            Width           =   1785
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.de Factura"
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
            Left            =   2580
            TabIndex        =   97
            Top             =   830
            Width           =   1365
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Autorización"
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
            Left            =   5040
            TabIndex        =   96
            Top             =   830
            Width           =   1875
         End
         Begin VB.Label lbl_doc01 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   1635
            TabIndex        =   95
            Top             =   660
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lbl_docnro1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "doc_numero"
            DataSource      =   "Ado_datos01"
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
            TabIndex        =   94
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            Caption         =   "Código Registro"
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
            Index           =   0
            Left            =   120
            TabIndex        =   93
            Top             =   830
            Width           =   1470
         End
         Begin VB.Label Txt_cod_cobro1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_codigo"
            DataSource      =   "Ado_datos01"
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
            Left            =   1635
            TabIndex        =   92
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label TxtCmpbte1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_nro_factura"
            DataSource      =   "Ado_datos01"
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
            Left            =   2580
            TabIndex        =   91
            Top             =   1125
            Width           =   1320
         End
         Begin VB.Label TxtAutorizacion1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_nro_autorizacion"
            DataSource      =   "Ado_datos01"
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
            Left            =   4965
            TabIndex        =   90
            Top             =   1125
            Width           =   1995
         End
      End
      Begin VB.Frame FraNavega 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4750
         Left            =   -74960
         TabIndex        =   77
         Top             =   1440
         Width           =   5745
         Begin VB.OptionButton OptFilGral2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Facturados No Cobrados"
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
            Left            =   2520
            TabIndex        =   79
            Top             =   4395
            Width           =   2475
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
            Left            =   960
            TabIndex        =   78
            Top             =   4395
            Value           =   -1  'True
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dg_datos 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EF56
            Height          =   4020
            Left            =   75
            TabIndex        =   80
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   7091
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
               DataField       =   "cobranza_fecha_sol"
               Caption         =   "F.Solicit.Fac"
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
            BeginProperty Column02 
               DataField       =   "cobranza_codigo"
               Caption         =   "No.Cobranza"
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
            BeginProperty Column04 
               DataField       =   "cobranza_fecha_fac"
               Caption         =   "F.Facturacion"
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
               DataField       =   "cobranza_total_bs"
               Caption         =   "Facturad.Bs."
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
            BeginProperty Column07 
               DataField       =   "doc_numero"
               Caption         =   "Nro.Doc.Respaldo"
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
            BeginProperty Column09 
               DataField       =   "estado_codigo_fac"
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
               Caption         =   "NIT/CI del Cliente"
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
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
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
         Begin MSAdodcLib.Adodc Ado_datos 
            Height          =   330
            Left            =   75
            Top             =   4320
            Width           =   5595
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
         Height          =   4740
         Left            =   -69195
         TabIndex        =   38
         Top             =   1440
         Width           =   9135
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
            DataField       =   "cobranza_total_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   6045
            TabIndex        =   8
            Text            =   "0"
            Top             =   4320
            Width           =   1275
         End
         Begin VB.TextBox txt_tdc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            DataField       =   "cobranza_tdc"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   4680
            TabIndex        =   183
            Text            =   "6.96"
            Top             =   4320
            Width           =   915
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6715
            TabIndex        =   42
            Top             =   2485
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24EF6E
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5520
            TabIndex        =   43
            Top             =   2475
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "12345678901234"
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4040
            TabIndex        =   134
            Top             =   4290
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4040
            TabIndex        =   133
            Top             =   4090
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4040
            TabIndex        =   132
            Top             =   3585
            Width           =   255
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8460
            TabIndex        =   41
            Top             =   3045
            Width           =   255
         End
         Begin VB.TextBox TxtMontoDol 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            DataField       =   "cobranza_total_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            Height          =   285
            Left            =   7725
            TabIndex        =   40
            Text            =   "0"
            Top             =   4320
            Width           =   1275
         End
         Begin VB.TextBox TxtObs 
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos"
            Height          =   585
            Left            =   1080
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1620
            Width           =   5900
         End
         Begin VB.TextBox Txt_deposito 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "cmpbte_deposito"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6560
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "0"
            Top             =   3570
            Width           =   2430
         End
         Begin VB.CommandButton cmd_benef 
            BackColor       =   &H00808000&
            Height          =   320
            Left            =   8740
            Picture         =   "Frm_ao_ventas_seguimiento.frx":24EF87
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Beneficiario"
            Top             =   3030
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F111
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2085
            TabIndex        =   6
            Top             =   3030
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux5 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F12A
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7035
            TabIndex        =   44
            Top             =   3030
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_desc4A 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F143
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1695
            TabIndex        =   45
            Top             =   2475
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_cta 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F15C
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   46
            Top             =   4080
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_ctades 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F176
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2640
            TabIndex        =   47
            Top             =   3915
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo7 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F190
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   48
            Top             =   4395
            Visible         =   0   'False
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_codigo"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc6 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F1A9
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   49
            Top             =   3570
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "trans_descripcion"
            BoundColumn     =   "trans_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F1C2
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   50
            Top             =   2760
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_codigo6 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F1DB
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   51
            Top             =   3240
            Visible         =   0   'False
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "trans_codigo"
            BoundColumn     =   "trans_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_desc7 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F1F4
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2640
            TabIndex        =   52
            Top             =   4080
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "cta_descripcion"
            BoundColumn     =   "cta_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro 
            DataField       =   "cobranza_fecha_fac"
            DataSource      =   "Ado_datos"
            Height          =   300
            Left            =   7260
            TabIndex        =   5
            Top             =   2420
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
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
            CheckBox        =   -1  'True
            Format          =   84803585
            CurrentDate     =   41678
         End
         Begin VB.Label TxtDscto 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_deuda_dol2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   5925
            TabIndex        =   57
            Top             =   4455
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   $"Frm_ao_ventas_seguimiento.frx":24F20D
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
            Left            =   120
            TabIndex        =   143
            Top             =   4080
            Width           =   8715
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Transacción:                                                          Comprobante Deposito:"
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
            Left            =   120
            TabIndex        =   142
            Top             =   3600
            Width           =   6465
         End
         Begin VB.Label lbl_doc1 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   1635
            TabIndex        =   70
            Top             =   660
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                     Nro.Venta:                    Nro.Cuota:"
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
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   6150
         End
         Begin VB.Label TxtAutorizacion 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   4920
            TabIndex        =   76
            Top             =   1080
            Width           =   1995
         End
         Begin VB.Label TxtCmpbte 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   2700
            TabIndex        =   75
            Top             =   1060
            Width           =   1200
         End
         Begin VB.Label Txt_cod_cobro 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   1635
            TabIndex        =   74
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Lbl_nombre_fac 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de:                                                                                                 NIT/CI"
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
            Left            =   120
            TabIndex        =   73
            Top             =   3045
            Width           =   6885
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7100
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Registro"
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
            Left            =   120
            TabIndex        =   72
            Top             =   800
            Width           =   1470
         End
         Begin VB.Label lbl_docnro 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            TabIndex        =   71
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Autorización"
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
            Left            =   5040
            TabIndex        =   69
            Top             =   800
            Width           =   1875
         End
         Begin VB.Label lbl_factura 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.de Factura"
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
            Left            =   2580
            TabIndex        =   68
            Top             =   800
            Width           =   1485
         End
         Begin VB.Label lbl_fechas 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Facturación"
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
            Left            =   7200
            TabIndex        =   67
            Top             =   2155
            Width           =   1785
         End
         Begin VB.Label Lbl_Cobrador 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobrador de CGI:"
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
            Left            =   120
            TabIndex        =   66
            Top             =   2495
            Width           =   1560
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado BOB"
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
            TabIndex        =   65
            Top             =   75
            Width           =   1785
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
            Left            =   120
            TabIndex        =   64
            Top             =   1755
            Width           =   960
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   6240
            TabIndex        =   63
            Top             =   255
            Width           =   735
         End
         Begin VB.Label lbl_fac 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   420
            TabIndex        =   62
            Top             =   1060
            Width           =   1005
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitud Fac"
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
            TabIndex        =   61
            Top             =   1415
            Width           =   1785
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   7100
            X2              =   7100
            Y1              =   0
            Y2              =   2855
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7100
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Label TxtNroVentaC 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
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
            Left            =   3960
            TabIndex        =   60
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label DTPFechaProg 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_fecha_sol"
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
            Left            =   7260
            TabIndex        =   59
            Top             =   1680
            Width           =   1710
         End
         Begin VB.Label TxtDsctoTot 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_solicitado_bs"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   7260
            TabIndex        =   58
            Top             =   345
            Width           =   1710
         End
         Begin VB.Label lbl_nit 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "beneficiario_codigo"
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
            Left            =   1080
            TabIndex        =   56
            Top             =   3120
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   9120
            Y1              =   3465
            Y2              =   3465
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00FFFF80&
            X1              =   7095
            X2              =   9115
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado USD"
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
            TabIndex        =   55
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_solicitado_dol"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   7260
            TabIndex        =   54
            Top             =   975
            Width           =   1710
         End
         Begin VB.Label TxtMonto2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "0"
            DataField       =   "cobranza_deuda_bs2"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
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
            Left            =   7725
            TabIndex        =   53
            Top             =   4395
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00FFFF80&
            X1              =   7095
            X2              =   9115
            Y1              =   2855
            Y2              =   2855
         End
      End
      Begin VB.Frame FraNavega1 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4750
         Left            =   40
         TabIndex        =   34
         Top             =   1440
         Width           =   5745
         Begin VB.OptionButton OptFilGral01 
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
            Left            =   840
            TabIndex        =   37
            Top             =   4395
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptFilGral02 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Aprobados No Facturados"
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
            Left            =   2400
            TabIndex        =   36
            Top             =   4395
            Width           =   2595
         End
         Begin MSDataGridLib.DataGrid dg_datos1 
            Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F29A
            Height          =   4020
            Left            =   75
            TabIndex        =   35
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   7091
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
            BeginProperty Column02 
               DataField       =   "cobranza_codigo"
               Caption         =   "No.Cobranza"
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
            BeginProperty Column04 
               DataField       =   "cobranza_fecha_sol"
               Caption         =   "F.Solicit.Fac."
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
               DataField       =   "cobranza_deuda_bs"
               Caption         =   "P/Cobrar.Bs."
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
            BeginProperty Column07 
               DataField       =   "doc_numero"
               Caption         =   "Nro.Doc.Respaldo"
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
            BeginProperty Column09 
               DataField       =   "estado_codigo_sol"
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
               Caption         =   "NIT/CI del Cliente"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
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
         Begin MSAdodcLib.Adodc Ado_datos01 
            Height          =   330
            Left            =   75
            Top             =   4320
            Width           =   5595
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
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00000000&
      Caption         =   "DATOS DE LA VENTA"
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
      TabIndex        =   29
      Top             =   6345
      Width           =   12975
      Begin MSDataGridLib.DataGrid dg_datos16 
         Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F2B4
         Height          =   1170
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   2064
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Tramite"
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
         BeginProperty Column02 
            DataField       =   "edif_descripcion"
            Caption         =   "Denominacion del Edificio"
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
            DataField       =   "zona_denominacion"
            Caption         =   "Zona"
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
            DataField       =   "calle_tipo"
            Caption         =   "Via.Acceso"
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
            DataField       =   "calle_denominacion"
            Caption         =   "Nombre de Calle, Av u otro"
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
            DataField       =   "edif_nro"
            Caption         =   "Nro."
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
         BeginProperty Column08 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Cliente/Representante.Legal"
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
            DataField       =   "venta_fecha_inicio"
            Caption         =   "F.Inicio.Contrato"
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
            DataField       =   "venta_fecha_fin"
            Caption         =   "F.Fin.Contrato"
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
            DataField       =   "venta_cantidad_total"
            Caption         =   "Cantidad"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Periodicidad"
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
         BeginProperty Column13 
            DataField       =   "venta_monto_total_bs"
            Caption         =   "Total,Contrato.Bs"
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
         BeginProperty Column14 
            DataField       =   "venta_monto_cobrado_bs"
            Caption         =   "Cobrado.Bs"
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
         BeginProperty Column15 
            DataField       =   "venta_saldo_p_cobrar_bs"
            Caption         =   "Saldo.P/Cobar"
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
         BeginProperty Column16 
            DataField       =   "unidad_codigo"
            Caption         =   "Unidad.E."
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
         BeginProperty Column17 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
         BeginProperty Column18 
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
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2564.788
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3060.284
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   675.213
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE BIENES / SERVICIOS VENDIDOS"
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
      TabIndex        =   28
      Top             =   7965
      Width           =   12975
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "Frm_ao_ventas_seguimiento.frx":24F2CE
         Height          =   1620
         Left            =   120
         TabIndex        =   33
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
            Caption         =   "Codigo.Bien"
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
            Caption         =   "Descripcion y Características del Bien"
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
            Caption         =   "Modelo.Vendido"
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
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   615.118
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   240
      Top             =   9840
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
      Top             =   9840
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
   Begin Crystal.CrystalReport CryF01 
      Left            =   1200
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   6840
      Top             =   10800
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
      Left            =   9120
      Top             =   10800
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
      Left            =   11400
      Top             =   10440
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
      Left            =   13080
      Top             =   9720
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
   Begin Crystal.CrystalReport CryF02 
      Left            =   1680
      Top             =   9840
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
   Begin Crystal.CrystalReport CryQ01 
      Left            =   2160
      Top             =   9840
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
End
Attribute VB_Name = "Frm_ao_ventas_seguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ventas
'INI QR
'Enum TQRCodeEncoding
'ceALPHA
'ceBYTE
'ceNUMERIC
'ceKANJI
'ceAUTO
'End Enum
'Enum TQRCodeECLevel
'LEVEL_L
'LEVEL_M
'LEVEL_Q
'LEVEL_H
'End Enum
'Private Declare Sub FullQRCode Lib "QRCodeLib.dll" _
'(ByVal autoConfigurate As Boolean, _
' ByVal AutoFit As Boolean, _
' ByVal backColor As Long, _
' ByVal barColor As Long, _
' ByVal Texto As String, _
' ByVal correctionLevel As TQRCodeECLevel, _
' ByVal encoding As TQRCodeEncoding, _
' ByVal marginpixels As Integer, _
' ByVal moduleWidth As Integer, _
' ByVal Height As Integer, _
' ByVal Width As Integer, _
' ByVal FileName As String)
Private Declare Sub FastQRCode Lib "QRCodeLib.dll" (ByVal Texto As String, ByVal FileName As String)
'Private Declare Function QRCodeLibVer Lib "QRCodeLib.dll" () As String
'Dim sFile As String
'Dim CadenaQ As String
'FIN QR
Dim rs_datos As New ADODB.Recordset     'FACTURACION
Dim rs_datos01 As New ADODB.Recordset     'INICIO COBRANZAS
Dim rs_datos02 As New ADODB.Recordset     'REG. COBRANZAS
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos4A As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset
Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset   'Cta Bancaria

Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset

'CLASIFICADORES
Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset
'IMAGENES
Dim m_stream    As ADODB.Stream
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
Dim queryinicial As String
Dim queryinicial1 As String
Dim queryinicial2 As String

'Dim descri_bien As String
'VARIABLES
Dim iResult As Variant  ', i%, y%
Dim marca1 As Variant

Dim VAR_CANT As Integer         'Cant_Alm,
Dim correlativo1 As Integer
Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, correlv, NRO_COBR As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SW, VAR_TSOL As Integer
Dim VAR_SOL As Integer
Dim i As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, COBR_BS As Double
Dim VAR_CONTAB As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CTA, VAR_PROY3 As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3 As String
Dim VAR_ANIO, VAR_MES, VAR_DIA, VAR_FECHA As String
Dim VAR_COD4, VAR_TIPOV, VAR_CITE  As String
Dim DESAUX, VARAUX, VARCODIG As String

Dim Numero As String
Dim Autorizacion As String
Dim NroFactura As String
Dim NitCi As String
Dim Fecha As String
Dim Monto As String
Dim Llave As String
Dim CodigoContro As String
'Dim Exel As New Excel.Application
Dim fs As FileSystemObject      'Variable de tipo file System Object
    
Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  Dim descri_bien As String
'  Dim Cant_Alm As Integer
  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'            BtnModificar.Visible = True
            If Ado_datos.Recordset!doc_codigo_fac = "R-103" Then
                BtnImprimir3.Caption = "Recibo"
                lbl_factura.Caption = "Nro.de Recibo"
            Else
                BtnImprimir3.Caption = "Factura"
                lbl_factura.Caption = "Nro.de Factura"
            End If
            If (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 16) Then
                TxtDsctoTot.backColor = &HFF&             'ROJO
                DTPFechaProg.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos.Recordset("cobranza_fecha_sol") > Date - 16) And (Ado_datos.Recordset("cobranza_fecha_sol") <= Date - 1) Then
                    TxtDsctoTot.backColor = &H80FF&           'NARANJA
                    DTPFechaProg.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
        Else
'            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
'            BtnAprobar.Visible = False
'            BtnVer.Visible = True
'            FrmABMDet.Visible = False
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
            TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
            If glusuario = "RVEIZAGA" Or glusuario = "HBUSTILLOS" Or glusuario = "ADMIN" Then
                BtnImprimir3.Visible = True
            Else
                BtnImprimir3.Visible = False
            End If
        End If

        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and correl_venta = " & Ado_datos.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
        Else
            deta2 = 0
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
    BtnImprimir3.Visible = False
'                BtnDesAprobar.Visible = True
'    BtnModificar.Visible = False
'    BtnEliminar.Visible = False
'    BtnVer.Visible = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub

Private Sub Ado_datos01_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos01.Recordset.BOF) And (Not Ado_datos01.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos01.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos01.Recordset("estado_codigo_sol") = "REG") Then          'REG
            If (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 16) Then
                TxtDsctoTot1.backColor = &HFF&             'ROJO
                DTPFechaProg1.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos01.Recordset("cobranza_fecha_prog") > Date - 16) And (Ado_datos01.Recordset("cobranza_fecha_prog") <= Date - 1) Then
                    TxtDsctoTot1.backColor = &H80FF&           'NARANJA
                    DTPFechaProg1.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
'            BtnModificar1.Visible = True
'            BtnAprobar1.Visible = True
            If Ado_datos01.Recordset!doc_codigo_fac = "R-103" Then
                cmd_fac = "RECIBO"
            Else
                cmd_fac = "FACTURA"
            End If
        Else
'            BtnModificar1.Visible = False
'            BtnAprobar1.Visible = False
            TxtDsctoTot1.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg1.backColor = &H404040       '&H80000013      'Fondo Oscuro
        End If
'        If Ado_datos01.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos01.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
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
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos01.Recordset!venta_codigo & " and correl_venta = " & Ado_datos01.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'TxtMontoBs.Text = Ado_datos01.Recordset!monto_total_bS
            'TxtMontoUs.Text = Ado_datos01.Recordset!deuda_cobrada
            'Text2.Text = Ado_datos01.Recordset!saldo_p_cobrar
            'Call AbreAlmacen
        Else
            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos01.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos01.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos01.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
'    BtnAprobar1.Visible = False
'    BtnModificar1.Visible = False
    'BtnEliminar1.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF
End Sub


Private Sub AbreAlmacen()
'    Set rs_datos13 = New ADODB.Recordset
'    If rs_datos13.State = 1 Then rs_datos13.Close
'    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_datos13.Recordset = rs_datos13
'    Ado_datos13.Refresh

End Sub

Private Sub Ado_datos02_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Ado_datos02.Recordset.BOF) And (Not Ado_datos02.Recordset.EOF) Then   'EOF
     If Not IsNull(Ado_datos02.Recordset("venta_codigo")) Then            'venta_codigo
        If (Ado_datos02.Recordset("estado_codigo_bco") = "REG") Then          'REG
            If (Ado_datos02.Recordset("cobranza_fecha_prog") <= Date - 16) Then
                TxtDsctoTot1.backColor = &HFF&             'ROJO
                DTPFechaProg1.backColor = &HFF&             'ROJO
            Else
                If (Ado_datos02.Recordset("cobranza_fecha_prog") > Date - 16) And (Ado_datos02.Recordset("cobranza_fecha_prog") <= Date - 1) Then
                    TxtDsctoTot2.backColor = &H80FF&           'NARANJA
                    DTPFechaProg2.backColor = &H80FF&           'NARANJA
                Else
                    TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
                    DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
                End If
            End If
            BtnModificar2.Visible = True
            BtnAprobar2.Visible = True
            OptFilGral05.Visible = False
            If (glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS") Then
                OptFilGral05.Visible = True
            Else
                OptFilGral05.Visible = False
            End If
        Else
            TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
            If Ado_datos02.Recordset!estado_codigo = "APR" Then
'                BtnAprobar.Visible = False
'                BtnAprobar2.Visible = False
'                BtnModificar2.Visible = False
                OptFilGral05.Visible = False
            Else
                If (glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS") Then
'                    BtnAprobar.Visible = True
'                    BtnAprobar2.Visible = False
'                    BtnModificar2.Visible = True
                    OptFilGral05.Visible = True
                Else
                    BtnAprobar.Visible = False
                    BtnAprobar2.Visible = False
                    BtnModificar2.Visible = False
                    OptFilGral05.Visible = False
                End If
            End If
        End If

'        If Ado_datos02.Recordset("beneficiario_codigo") <> "" Then
'            Set RS_BENEF = New ADODB.Recordset
'            If RS_BENEF.State = 1 Then RS_BENEF.Close
'            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos02.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
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
        Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        'queryinicial2 = "select * from ao_ventas_detalle where venta_codigo = " & Ado_datos02.Recordset!venta_codigo & " and correl_venta = " & Ado_datos02.Recordset!correl_venta & " "
        'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
            'TxtMontoBs.Text = Ado_datos02.Recordset!monto_total_bS
            'TxtMontoUs.Text = Ado_datos02.Recordset!deuda_cobrada
            'Text2.Text = Ado_datos02.Recordset!saldo_p_cobrar
            'Call AbreAlmacen
        Else
            deta2 = 0
'            'TxtMontoBs.Text = 0
'            'TxtMontoUs.Text = 0
'            'Text2.Text = 0
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from av_ventas_cabecera where venta_codigo = '" & Ado_datos02.Recordset!venta_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            VAR_PROY3 = Ado_datos16.Recordset!edif_codigo
            FrmCobranza.Visible = True
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        ''Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        
        FrmDetalle.Caption = "VENTA NRO. " + Str((Ado_datos02.Recordset("venta_codigo")))
        
        FrmCobranza.Caption = "DETALLE DE BIENES DE LA VENTA NRO. " + Str((Ado_datos02.Recordset("venta_codigo")))
        
'        TxtCobrador1 = Trim(dtc_desc4A.Text)
        
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo = '" & Ado_datos02.Recordset!cobranza_codigo & "' ", "Foto")
'        Image2 = Img_Foto
'        'If adoLista.Recordset!estado_codigo = "APR" Then
'        CmdFoto.Visible = True
     End If                         'venta_codigo
     FrmDetalle.Enabled = True
     FrmCobranza.Visible = True
  Else
'    BtnAprobar2.Visible = False
'    BtnModificar2.Visible = False
    'BtnEliminar2.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmABMDet.Visible = False
    FrmABMDet2.Visible = False
  End If                            'EOF

End Sub

Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
'        BtnModDetalle.Visible = True
        BtnImprimir4.Visible = True


    Else
        'BtnAprobar2.Visible = False
        'BtnImprimir2.Visible = False
        BtnImprimir4.Visible = False
        'BtnAnlDetalle2.Visible = False
'        BtnModDetalle.Visible = False
    End If
 Else
    'BtnAprobar2.Visible = False
    'BtnImprimir2.Visible = False
    BtnImprimir4.Visible = False
    'BtnAnlDetalle2.Visible = False
'    BtnModDetalle.Visible = False
 End If
End Sub

Private Sub BntImprimir2_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'MODULO DE COBRANZA' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS - KARDEX' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub

Private Sub BtnBuscar_Click()
'JQA
 If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial1
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnBuscar1_Click()
'JQA
 If Ado_datos01.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos1
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos01.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If

End Sub

Private Sub BtnBuscar2_Click()
 If Ado_datos02.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos2
      ClBuscaGrid.QueryUtilizado = queryinicial2
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos02.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "No se puede Procesar el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir_Click()
    Select Case SSTab1.Tab
        Case 0
          If Ado_datos01.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos01.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos01.Recordset!cobranza_codigo
            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos01.Recordset!Literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 1
          If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
        Case 2
          If Ado_datos02.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
            CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
    End Select

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
End Sub

'Private Sub BtnImprimir1_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
'
'    CryR01.Formulas(1) = "literalcobro = '" & Ado_datos02.Recordset!Literal & "' "
'    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_codigo & "' "
'    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'    iResult = CryR01.PrintReport
'    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If
'
'End Sub

Private Sub BtnImprimir3_Click()
  If Ado_datos.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
    If (Ado_datos.Recordset!factura_impresa = "S") And (glusuario = "RVEIZAGA" Or glusuario = "HBUSTILLOS" Or glusuario = "ADMIN") Then
      If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
        '===== ini GENERA EL CODIGO DE FACTURA ====
            gestion0 = Ado_datos.Recordset("ges_gestion")
            correlv = Ado_datos.Recordset("venta_codigo")
            nroventa = Ado_datos.Recordset("venta_codigo")
            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'
'            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_total_dol, 2)
'            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_total_bs, 2)
'            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
            var_literal = Ado_datos.Recordset!Literal
'
'            Llave = Trim(rs_aux1!dosifica_llave)
'            If dtc_aux5.Text Like " " Then
'                MsgBox "Error en el NIT del Cliente, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
'                Exit Sub
'            Else
'                NitCi = IIf(dtc_aux5.Text = "", Ado_datos.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
'            End If
'            Autorizacion = rs_aux1!dosifica_autorizacion
'            'Fecha = Val(Format((Date), "YYYYMMDD"))
'            'Monto = Redondeo((VAR_BS2), 0)
'            'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
'            'CodigoContro = CodigoControl(NroFactura)
'            If Autorizacion <> "" And NitCi <> "" And Llave <> "" And VAR_BS2 <> "0" And rs_aux1!CORREL >= 0 Then
'                VAR_SW = 1
'            Else
'                VAR_SW = 0
'                MsgBox "Error en Autorizacion, NIT o Llave, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
'                Exit Sub
'            End If
            VAR_COD1 = Ado_datos.Recordset!cobranza_nro_factura
            sino = MsgBox("Esta seguro(a) de RE-IMPRIMIR la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
            If sino = vbYes Then
'                'GENERA CORREL NOTA DEBITO POR DEPTO INI
'                Set rs_aux5 = New ADODB.Recordset
'                If rs_aux5.State = 1 Then rs_aux5.Close
'                'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
'                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'                If Not rs_aux5.EOF Then
'                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'                End If
'                'rs_aux5!Codigo = VAR_CONTAB
'                'rs_aux5.Update
'
                VAR_COD2 = Ado_datos.Recordset!cobranza_nro_autorizacion
'                NroFactura = Trim(Str(VAR_COD1))
'                Fecha = Val(Format((Date), "YYYYMMDD"))
'                Monto = Redondeo((VAR_BS2), 0)
'
'                CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'                If CodigoContro = "" Or CodigoContro = "0" Then
'                    VAR_SW = 0
'                    MsgBox "Error en Codigo de Control, Contactese con el Administrador o vuelva a intentar ...", , "Atención"
'                    Exit Sub
'                Else
'                    VAR_SW = 1
'                End If
'                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'                'Ado_datos.Recordset!correl_contab = VAR_CONTAB
'                If VAR_CONTAB < 10 Then
'                    'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'                End If
'                If VAR_CONTAB > 99 Then
''                    If VAR_CONTAB > 1200 Then
''                        MsgBox "El ND Finaliza en 6564 ... ", , "Atención"
''                    End If
'                   'Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'                End If
'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
''               'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'
'                '===== ini nombre archivo de la FACTURA ====
'                'db.Execute "update ao_ventas_cobranza set archivo_foto = '" & doc_codigo & "' + '-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R101-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                '===== fin nombre archivo de la FACTURA ====
'                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                'IMPRIMIR FACTURA
''                VAR_ANIO = CStr(glGestion)
''                VAR_MES = CStr(Month(Date))
''                VAR_DIA = CStr(Day(Date))
''                VAR_FECHA = VAR_ANIO & VAR_MES & VAR_DIA
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
'                'Dim F1
'                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
'                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
'                'frm_qr.Show vbModal
'                'NIT del emisor, Nombre o Razón Social del emisor, Número correlativo de Factura, Número de Autorización, Fecha de emisión, Importe de la compra, Código de Control, Fecha Límite de Emisión, 0, 0, NIT / NDI Comprador, Nombre o Razón Social del comprador
'
'                'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"
'                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'
'                db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'
''                Ado_datos.Recordset!estado_codigo_fac = "APR"
''                Ado_datos.Recordset.Update
'                'INI QR
''                'sFile = "C:\Tmp\QRCode.bmp"
''                '1003579028
''                '& "|" & Format(Trim("0"), "###0.00") _
''                'dtc_aux5.Text
''                sFile = App.Path & "\CLIENTES\QRCode.bmp"
''                CadenaQ = Trim("1018533029") _
''                & "|" & Trim(VAR_COD1) _
''                & "|" & Trim(VAR_COD2) _
''                & "|" & Format(Trim(Date), "DD/MM/YYYY") _
''                & "|" & Format(Trim(VAR_BS2), "###0.00") _
''                & "|" & Format(Trim(VAR_BS2), "###0.00") _
''                & "|" & Trim(CodigoContro) _
''                & "|" & Trim(dtc_aux5.Text) _
''                & "|" & Trim("0") _
''                & "|" & Trim("0") _
''                & "|" & Trim("0") _
''                & "|" & Trim("0")
''
''                'CadenaQ = Trim(txtNitEmisor.Text) _
''                '& "|" & Trim(txtNumeroFactura.Text) _
''                '& "|" & Trim(txtNumeroAutorizacion.Text) _
''                '& "|" & Format(Trim(txtFechaEmision.Text), "DD/MM/YYYY") _
''                '& "|" & Format(Trim(txtImporteCompra.Text), "###0.00") _
''                '& "|" & Format(Trim(txtFiscal.Text), "###0.00") _
''                '& "|" & Trim(txtCodigoControl.Text) _
''                '& "|" & Trim(txtNitComprador.Text) _
''                '& "|" & Trim(txtImporteICE.Text) _
''                '& "|" & Trim(txtGravadas.Text) _
''                '& "|" & Trim(txtNoFiscal) _
''                '& "|" & Trim(TxtDescuento)
'''                MsgBox CadenaQ
''                FastQRCode CadenaQ, sFile
''                Set Picture1.Picture = LoadPicture(sFile)
'                'FIN QR
                
                Call IMPRIME_QR
                'MsgBox CadenaQ
'                If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                End If
            Else
                'VAR_COD1 = "0"
                If rs_aux1.State = 1 Then rs_aux1.Close
                Exit Sub
            End If
        
        If rs_aux1.State = 1 Then rs_aux1.Close
        '===== fin TERMINA GENERACION DE FACTURA =====
        

'        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
'          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
'          'rs_aux1.Update
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
        
'        Dim iResult As Variant  ', i%, y%
'        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
'        CryF01.WindowShowRefreshBtn = True
'        CryF01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        CryF01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        CryF01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'
'        CryF01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
'        CryF01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryF01.PrintReport
'        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
        
'        TxtCmpbte.Caption = VAR_COD1
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
      End If
      If Ado_datos.Recordset!doc_codigo_fac = "R-103" Then
      'WWWWWWWWWWWWWWWWWWWWWWWWW
        '===== ini GENERA EL CODIGO DE RECIBO ====
        Set rs_aux1 = New ADODB.Recordset
        rs_aux1.CursorLocation = adUseClient
        If rs_aux1.State = 1 Then rs_aux1.Close
        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
            correlv = Ado_datos.Recordset("venta_codigo")
            nroventa = Ado_datos.Recordset("venta_codigo")
            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
'            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
'            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
            var_literal = Ado_datos.Recordset!Literal
'            'Llave = Trim(rs_aux1!dosifica_llave)
'            NitCi = IIf(dtc_aux5.Text = "", Ado_datos.Recordset!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
'            'Autorizacion = rs_aux1!dosifica_autorizacion
'            VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
'            VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
'            VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
'            VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
'            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
'
'            VAR_COD1 = CDbl(rs_aux1!correl_doc) + 1
            sino = MsgBox("Esta seguro(a) de RE-IMPRIMIR la Recibo Nro. " + Str(TxtCmpbte) + " ?", vbYesNo, "Confirmando")
            If sino = vbYes Then
                'GENERA CORREL NOTA DEBITO POR DEPTO INI
''                Set rs_aux5 = New ADODB.Recordset
''                If rs_aux5.State = 1 Then rs_aux5.Close
''                rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
''                If Not rs_aux5.EOF Then
''                    VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
''                End If
''                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
''                db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
''                If VAR_CONTAB < 10 Then
''                    VAR_GLOSA = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
''                   VAR_GLOSA = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
''                End If
''                If VAR_CONTAB > 99 And VAR_CONTAB < 6564 Then
''                    If VAR_CONTAB > 1200 Then
''                        MsgBox "El ND Finaliza en 6564 ... ", , "Atención"
''                    End If
''                   VAR_GLOSA = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
''                End If
'                VAR_GLOSA = TxtObs.Text
'                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
''                'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'
'                VAR_COD2 = "0"  'rs_aux1!dosifica_autorizacion
'                NroFactura = Trim(Str(VAR_COD1))
'                '===== ini nombre archivo de la FACTURA ====
'                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R103-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                '===== fin nombre archivo de la FACTURA ====
'                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                'IMPRIMIR FACTURA
'                Fecha = Val(Format((Date), "YYYYMMDD"))
'                Monto = Redondeo((VAR_BS2), 0)
'                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
'                'Dim F1
'                'FI = Ado_datos.Recordset!cobranza_fecha_cobro
'                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & Ado_datos.Recordset!cobranza_fecha_cobro & "' + "|" + " & Ado_datos.Recordset!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & Ado_datos.Recordset!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
'                'frm_qr.Show vbModal
'                'NIT del emisor, Nombre o Razón Social del emisor, Número correlativo de Factura, Número de Autorización, Fecha de emisión, Importe de la compra, Código de Control, Fecha Límite de Emisión, 0, 0, NIT / NDI Comprador, Nombre o Razón Social del comprador
'
'                'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"
'                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'
'                VAR_SW = 1
'                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
'                'db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                Call IMPRIME_RECIBO
'                If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                End If
            Else
'                VAR_COD1 = "0"
                If rs_aux1.State = 1 Then rs_aux1.Close
                Exit Sub
            End If
        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'        'TxtCmpbte.Caption = VAR_COD1
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
      'WWWWWWWWWWWWWWWWWWWWWWWWW
      End If
    Else
        MsgBox "La Factura Nro. " + Ado_datos.Recordset!cobranza_nro_factura + " NO fue Impresa o el Usuario NO está Autorizado...", , "Atención"
        'Call IMPRIME_FACTURA
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
    End If
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub generar(Autorizacion As String, Numero As String, NitCi As String, Fecha As String, Monto As String, Llave As String)
' paso 1
'    Dim suma As String
'    Dim digitos As String
'    Dim digitossum(4) As Integer
'    Dim cadenas(4) As String
'    Dim inicio As Integer
'    Dim x As Integer
'
'    Dim arc4 As String
'    Dim suma_total As Long
'    Dim sumas(4) As Long
'    Dim strlen_arc4 As Integer
'    Dim i As Integer
'    Dim total As Long
'
'    Dim mensaje As String
'    Dim last As String
'
'        numero = verhoeff_add_recursive(numero, 2)
'        nitci = verhoeff_add_recursive(nitci, 2)
'        fecha = verhoeff_add_recursive(fecha, 2)
'        monto = verhoeff_add_recursive(monto, 2)
''            Dim suma As String = CType((Long.Parse(numero) _
''                        + (Long.Parse(nitci) _
''                        + (Long.Parse(fecha) + Long.Parse(monto)))),Long).ToString
'        suma = (CStr(numero) + (CStr(nitci) + (Trim(fecha) + CStr(monto))))
'        suma = verhoeff_add_recursive(suma, 5)
'' paso2
''            Dim digitos As String = ("" + suma.Substring((suma.Length - 5), 5))
''            Dim digitossum() As Integer = New Integer() {0, 0, 0, 0, 0}
''            Dim cadenas() As String = New String() {"", "", "", "", ""}
''            Dim inicio As Integer = 0
''            Dim x As Integer = 0
'    digitos = ("" + suma.Substring((suma.Length - 5), 5))
'    digitossum(0) = 0
'    digitossum(1) = 0
'    digitossum(2) = 0
'    digitossum(3) = 0
'    digitossum(4) = 0
'    cadenas(0) = ""
'    cadenas(1) = ""
'    cadenas(2) = ""
'    cadenas(3) = ""
'    cadenas(4) = ""
'    inicio = 0
'    x = 0
''    For Each d As Char In digitos.ToCharArray
''                digitossum(x) = (Integer.Parse(d.ToString) + 1)
''                cadenas(x) = llave.Substring(inicio, (Integer.Parse(d.ToString) + 1))
''                inicio = (inicio _
''                            + (Integer.Parse(d.ToString) + 1))
''                x = (x + 1)
''    Next
'    For x = 0 To Len(digitos)
'        digitossum(x) = (CInt(digitos) + 1)
'        cadenas(x) = llave.Substring(inicio, (CInt(digitos) + 1))
'        inicio = (inicio + (CInt(digitos) + 1))
'        x = (x + 1)
'    Next x
'            autorizacion = (autorizacion + cadenas(0))
'            numero = (numero + cadenas(1))
'            nitci = (nitci + cadenas(2))
'            fecha = (fecha + cadenas(3))
'            monto = (monto + cadenas(4))
'' paso3
'    arc4 = allegedrc4((autorizacion + (numero + (nitci + (fecha + monto)))), (llave + digitos))
'' paso4
'    suma_total = 0
'    sumas(0) = 0
'    sumas(1) = 0
'    sumas(2) = 0
'    sumas(3) = 0
'    sumas(4) = 0
'    strlen_arc4 = Len(arc4)
'    i = 0
'    Do While (i < strlen_arc4)
'                x = CInt(arc4(i))
'                sumas((i Mod 5)) = (sumas((i Mod 5)) + x)
'                suma_total = (suma_total + x)
'                i = (i + 1)
'    Loop
'' paso5
'    total = 0
'    i = 0
'    Do While (i < Len(sumas))
'                total = (total + (suma_total * (sumas(i) / digitossum(i))))
'                i = (i + 1)
'    Loop
'    mensaje = big_base_convert(total, 64)
'    last = allegedrc4(mensaje, (llave + digitos)).Insert(2, "-").Insert(5, "-").Insert(8, "-")
'            If (last.Length > 11) Then
'                last = last.Insert(11, "-")
'            End If
'    'Return last

End Sub

Private Sub big_base_convert(ByVal Numero As Long, ByVal baseconv As Long)
'    Dim dic(63) As Char
'    Dim cociente As Long
'    Dim resto As Long
'    Dim palabra As String
'
'    dic(0) = Microsoft.VisualBasic.ChrW(48)
'    dic(1) = Microsoft.VisualBasic.ChrW(49)
'    dic(2) = Microsoft.VisualBasic.ChrW(50)
'    dic(3) = Microsoft.VisualBasic.ChrW(51)
'    dic(4) = Microsoft.VisualBasic.ChrW(52)
'    dic(5) = Microsoft.VisualBasic.ChrW(53)
'    dic(6) = Microsoft.VisualBasic.ChrW(54)
'    dic(7) = Microsoft.VisualBasic.ChrW(55)
'    dic(8) = Microsoft.VisualBasic.ChrW(56)
'    dic(9) = Microsoft.VisualBasic.ChrW(57)
'    dic(10) = Microsoft.VisualBasic.ChrW(65)
'    dic(11) = Microsoft.VisualBasic.ChrW(66)
'    dic(12) = Microsoft.VisualBasic.ChrW(67)
'    dic(13) = Microsoft.VisualBasic.ChrW(68)
'    dic(14) = Microsoft.VisualBasic.ChrW(69)
'    dic(15) = Microsoft.VisualBasic.ChrW(70)
'    dic(16) = Microsoft.VisualBasic.ChrW(71)
'    dic(17) = Microsoft.VisualBasic.ChrW(72)
'    dic(18) = Microsoft.VisualBasic.ChrW(73)
'    dic(19) = Microsoft.VisualBasic.ChrW(74)
'    dic(20) = Microsoft.VisualBasic.ChrW(75)
'    dic(21) = Microsoft.VisualBasic.ChrW(76)
'    dic(22) = Microsoft.VisualBasic.ChrW(77)
'    dic(23) = Microsoft.VisualBasic.ChrW(78)
'    dic(24) = Microsoft.VisualBasic.ChrW(79)
'    dic(25) = Microsoft.VisualBasic.ChrW(80)
'    dic(26) = Microsoft.VisualBasic.ChrW(81)
'    dic(27) = Microsoft.VisualBasic.ChrW(82)
'    dic(28) = Microsoft.VisualBasic.ChrW(83)
'    dic(29) = Microsoft.VisualBasic.ChrW(84)
'    dic(30) = Microsoft.VisualBasic.ChrW(85)
'    dic(31) = Microsoft.VisualBasic.ChrW(86)
'    dic(32) = Microsoft.VisualBasic.ChrW(87)
'    dic(33) = Microsoft.VisualBasic.ChrW(88)
'    dic(34) = Microsoft.VisualBasic.ChrW(89)
'    dic(35) = Microsoft.VisualBasic.ChrW(90)
'    dic(36) = Microsoft.VisualBasic.ChrW(97)
'    dic(37) = Microsoft.VisualBasic.ChrW(98)
'    dic(38) = Microsoft.VisualBasic.ChrW(99)
'    dic(39) = Microsoft.VisualBasic.ChrW(100)
'    dic(40) = Microsoft.VisualBasic.ChrW(101)
'    dic(41) = Microsoft.VisualBasic.ChrW(102)
'    dic(42) = Microsoft.VisualBasic.ChrW(103)
'    dic(43) = Microsoft.VisualBasic.ChrW(104)
'    dic(44) = Microsoft.VisualBasic.ChrW(105)
'    dic(45) = Microsoft.VisualBasic.ChrW(106)
'    dic(46) = Microsoft.VisualBasic.ChrW(107)
'    dic(47) = Microsoft.VisualBasic.ChrW(108)
'    dic(48) = Microsoft.VisualBasic.ChrW(109)
'    dic(49) = Microsoft.VisualBasic.ChrW(110)
'    dic(50) = Microsoft.VisualBasic.ChrW(111)
'    dic(51) = Microsoft.VisualBasic.ChrW(112)
'    dic(52) = Microsoft.VisualBasic.ChrW(113)
'    dic(53) = Microsoft.VisualBasic.ChrW(114)
'    dic(54) = Microsoft.VisualBasic.ChrW(115)
'    dic(55) = Microsoft.VisualBasic.ChrW(116)
'    dic(56) = Microsoft.VisualBasic.ChrW(117)
'    dic(57) = Microsoft.VisualBasic.ChrW(118)
'    dic(58) = Microsoft.VisualBasic.ChrW(119)
'    dic(59) = Microsoft.VisualBasic.ChrW(120)
'    dic(60) = Microsoft.VisualBasic.ChrW(121)
'    dic(61) = Microsoft.VisualBasic.ChrW(122)
'    dic(62) = Microsoft.VisualBasic.ChrW(43)
'    dic(63) = Microsoft.VisualBasic.ChrW(47)
'
'    cociente = 1
'    resto = 0
'    palabra = ""
'    While (cociente > 0)
'                cociente = (numero / baseconv)
'                resto = (numero Mod baseconv)
'                palabra = (dic(resto) + palabra)
'                numero = cociente
'
'    End
'    '        Return palabra
End Sub
        
Private Sub SWAP(ByRef num1 As Integer, ByRef num2 As Integer)
    Dim temp As Integer
    temp = num2
    num2 = num1
    num1 = temp
End Sub
        
'Private Sub allegedrc4(mensaje As String, llaverc4 As String)
'            Dim state() As Integer = New Integer((256) - 1) {}
'            Dim x As Integer = 0
'            Dim y As Integer = 0
'            Dim index1 As Integer = 0
'            Dim index2 As Integer = 0
'            Dim nmen As Integer = 0
'            Dim i As Integer = 0
'            Dim cifrado As String = ""
'            i = 0
'            Do While (i < 256)
'                state(i) = i
'                i = (i + 1)
'            Loop
'            Dim strlen_llave As Integer = llaverc4.Length
'            Dim strlen_mensaje As Integer = mensaje.Length
'            i = 0
'            Do While (i < 256)
'                index2 = ((CType(llaverc4(index1),Integer) _
'                            + (state(i) + index2)) _
'                            Mod 256)
'                swap(state(index2), state(i))
'                index1 = ((index1 + 1) _
'                            Mod strlen_llave)
'                i = (i + 1)
'            Loop
'            Dim cadtemp As String = ""
'            i = 0
'            Do While (i < strlen_mensaje)
'                x = ((x + 1) _
'                            Mod 256)
'                y = ((state(x) + y) _
'                            Mod 256)
'                swap(state(y), state(x))
'                ' ^ = XOR function
'                nmen = (CType(mensaje(i),Integer) Or state(((state(x) + state(y)) _
'                            Mod 256)))
'                'The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
'                cadtemp = ("0" + big_base_convert(nmen, 16))
'                cifrado = (cifrado + cadtemp.Substring((cadtemp.Length - 2), 2))
'                i = (i + 1)
'            Loop
'            Return cifrado
'End Sub
'
'Private Shared Function calcsum(ByVal number As String) As Integer
'            Dim c As Integer = 0
'            Dim n As String = reverse(number)
'            Dim len As Integer = n.Length
'            Dim nchar() As Char = n.ToCharArray
'            Dim i As Integer = 0
'            Do While (i < len)
'                c = table_d(c, table_p(((i + 1) _
'                            Mod 8), Integer.Parse(nchar(i).ToString)))
'                i = (i + 1)
'            Loop
'            Return table_inv(c)
'End Sub
'
'Private Shared Function verhoeff_add_recursive(ByVal number As String, ByVal digits As Integer) As String
'            Dim temp As String = number
'
'            While (digits > 0)
'                temp = (temp + calcsum(temp))
'                digits = (digits - 1)
'
'            End While
'            Return temp
'End Sub
'
'Private Shared Function reverse(ByVal cadena As String) As String
'            Dim str() As Char = cadena.ToCharArray
'            Array.Reverse(str)
'            Return New String(str)
'End Sub

Private Sub IMPRIME_FACTURA()
        'IMPRIMIR FACTURA
    Dim iResult As Variant  ', i%, y%
    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior_rep.rpt"
    Else
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_anterior.rpt"
    End If
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

End Sub

Private Sub IMPRIME_QR()
    'RE-IMPRIMIR FACTURA con QR
    'Dim Exel As Object
    'Set Exel = CreateObject("Excel.Application")
    'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
    'Exel.Visible = True
    Call CmdFoto_Click
    ImagenQr = App.Path & "\FACTURAS\" & Trim(VAR_COD2) & "\" & Trim(Ado_datos.Recordset!doc_codigo_fac) & "-" & Trim(Str(VAR_COD1)) & ".JPG"
    ''QRCode.bmp"
    
'    Picture2.AutoRedraw = True
'    Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'    ImagenQr = App.Path & "\FACTURAS\" & Trim(VAR_COD2) & "\" & Trim(Ado_datos.Recordset!doc_codigo_fac) & "-" & Trim(Str(VAR_COD1)) & ".JPG"
'    'ImagenQr = App.Path & "\CLIENTES\" & Trim(VARCOD2) & "\" & Trim(Str(VARCOD1)) & ".JPG"      'QRCode.bmp"
    ' MsgBox CadenaQr
'    FastQRCode CadenaQr, ImagenQr
    Picture1.AutoRedraw = True
    Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Clipboard.Clear
'    Clipboard.SetData Picture2.Image
''    Exel.Application.Range("a2").Select
''    Exel.Application.ActiveSheet.Paste

    Dim iResult As Variant  ', i%, y%
    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
    Else
        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
    End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"

End Sub

Private Sub IMPRIME_RECIBO()
        'IMPRIMIR FACTURA
        Dim iResult As Variant  ', i%, y%
        'CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_oficial.rpt"
        CryF01.WindowShowRefreshBtn = True
        CryF01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
        CryF01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryF01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        'var_literal = "-"   'Ado_datos.Recordset!Literal
        CryF01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryF01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryF01.PrintReport
        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"

End Sub
Private Sub BtnImprimir4_Click()
    Select Case SSTab1.Tab
        Case 0
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos01.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos01.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos01.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos01.Recordset!cobranza_prog_codigo & "' "
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
        Case 1
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_prog_codigo & "' "
              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
        Case 2  'Ado_datos02
            If Ado_datos16.Recordset.RecordCount > 0 Then
              'CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_R105_kardex.rpt"
              CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
              CryV01.WindowShowRefreshBtn = True
              CryV01.StoredProcParam(0) = Me.Ado_datos02.Recordset!ges_gestion            'glGestion
              CryV01.StoredProcParam(1) = Me.Ado_datos02.Recordset!venta_codigo           'nroventa        '
              CryV01.StoredProcParam(2) = Me.Ado_datos02.Recordset!cobranza_prog_codigo   'NRO_COBR        '
              'Literal por el Total de la Compra
              var_literal = Literal(CStr(Ado_datos16.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
              CryV01.Formulas(1) = "literalcobro = '" & var_literal & "' "
              'CryV01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
              CryV01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_prog_codigo & "' "
              '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
              iResult = CryV01.PrintReport
              If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
            Else
              MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
            End If
    End Select
  
End Sub
Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

Private Sub BtnSalir2_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmCabecera.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir3_Click()
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'
'    FrmEdita.Visible = False
''    BtnImprimir2.Visible = True
''    BtnImprimir3.Visible = True
End Sub

Private Sub BtnSalir1_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If

End Sub

Private Sub cmd_benef_Click()
    Set rs_datos8 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_beneficiario where tipoben_codigo <> '0' and tipoben_codigo <> '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    If Ado_datos8.Recordset.RecordCount > 0 Then
        dtc_desc8.BoundText = dtc_codigo8.BoundText
'        FraGrabarCancelar.Enabled = False
'        frm_benef.Visible = True
    End If
End Sub

Private Sub cmd_moneda1_LostFocus()
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda1.Text & "' ", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub cmd_moneda2_LostFocus()
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria where tipo_moneda = '" & cmd_moneda2.Text & "' ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub CmdFoto_Click()
'    Frm_Imprime_Factura.Show

    On Error GoTo QError
    Set fs = New FileSystemObject   'Creamos la Nueva referencia Fso
    
    Set rs_aux6 = New ADODB.Recordset     'Iniciales del Cliente - gc_beneficiario
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        db.Execute "update ao_ventas_cobranza set beneficiario_iniciales = '" & rs_aux6!beneficiario_iniciales & "'   Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
    End If
    'If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
    If Ado_datos.Recordset!archivo_foto_cargado = "N" Or IsNull(Ado_datos.Recordset!archivo_foto_cargado) Then
'      NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
'      DirOrigen = App.Path & "\CLIENTES\"
'      DirDestino = App.Path & "\CLIENTES\"
'      'DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
'      fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
'      Ado_datos.Recordset!ARCHIVO_Foto = Trim(Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      
      NombreCarpeta = App.Path & "\FACTURAS\" & Trim(VAR_COD2) & "\"
      DirOrigen = App.Path & "\FACTURAS\" & Trim(VAR_COD2) & "\"
      DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
      Ado_datos.Recordset!ARCHIVO_Foto = Trim(Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      fs.CopyFile DirOrigen & Ado_datos.Recordset!ARCHIVO_Foto, DirDestino & "\" & Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG"       'Ado_datos.Recordset!cobranza_nro_factura        'ARCHIVO_Foto
      'Ado_datos.Recordset!ARCHIVO_Foto = Trim(Ado_datos.Recordset!doc_codigo_fac & "-" & Trim(Str(VAR_COD1)) & ".JPG")
      Ado_datos.Recordset!archivo_foto_cargado = "S"
      
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "Q_R"
''      If GlServidor = "SERVIDOR2" Then
''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
''      Else
'         e = NombreCarpeta
''      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          DirOrigen = App.Path & "\CLIENTES\"
          DirDestino = App.Path & "\CLIENTES\" & Trim(rs_aux6!beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
          fs.CopyFile DirOrigen & "\QRCode.bmp", DirDestino & "\" & Ado_datos.Recordset!ARCHIVO_Foto
          frmBeneficiario_Admin.adoLista.Recordset!archivo_foto_cargado = "S"
          
    '      Frmexporta.DirDestino.Path = NombreCarpeta
    '      GlArch = "Q_R"
    ''      If GlServidor = "SERVIDOR2" Then
    ''         e = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" & Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) & "-" & Trim(Ado_datos.Recordset!beneficiario_codigo) & "\"
    ''      Else
    '         e = NombreCarpeta
    ''      End If
    '      Frmexporta.DirDestino2.Path = e
    '      Frmexporta.Show vbModal      End If
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SERVIDOR2" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\CLIENTES\" + Trim(Ado_datos.Recordset!beneficiario_beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        'ARCH_FOTO = App.Path + "\CLIENTES\" + Trim(rs_aux6!beneficiario_iniciales) + "-" + Trim(Ado_datos.Recordset("beneficiario_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
        ARCH_FOTO = App.Path & "\FACTURAS\" & Trim(VAR_COD2) & "\" & Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "CLIENTES" + "\" + Ado_datos.Recordset!beneficiario_codigo + "\" + Ado_datos.Recordset("beneficiario_codigo") + "-FOTO.JPG"
    CodBenef = Ado_datos.Recordset!cobranza_codigo
    'If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where beneficiario_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
    If Guardar_Imagen(db, "Select Foto From ao_ventas_cobranza Where cobranza_codigo= '" & CodBenef & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
        Exit Sub
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    'dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub BntImprimir3_Click()
    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas_dol.rpt"
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub

Private Sub dtc_codigo4A1_Click(Area As Integer)
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_cta_Click(Area As Integer)
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub dtc_ctades_Click(Area As Integer)
    dtc_cta.BoundText = dtc_ctades.BoundText
End Sub

Private Sub dtc_desc4A1_Click(Area As Integer)
    dtc_codigo4A1.BoundText = dtc_desc4A1.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
End Sub

Private Sub dtc_aux5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    dtc_aux5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodmanejo.BoundText
    DtCDescripcion.BoundText = dtccodmanejo.BoundText
    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
    DtCCodigo.BoundText = dtccodpeso.BoundText
    DtCDescripcion.BoundText = dtccodpeso.BoundText
    dtcunidadmedida.BoundText = dtccodpeso.BoundText
    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub


Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
    dtcdespoa.Text = dtccodpoa.BoundText
End Sub

Private Sub dtccodpuesto_Click(Area As Integer)
    dtcdenopuesto.Text = dtccodpuesto.BoundText
End Sub

Private Sub dtccodtipoid_Click(Area As Integer)
    dtcdescrtipoid.BoundText = dtccodtipoid.BoundText
End Sub

Private Sub dtccoduni_Click(Area As Integer)
    dtcdescripuni.Text = dtccoduni.BoundText
End Sub

Private Sub dtccorrcompromiso_Click(Area As Integer)
    dtcfechacompromiso.BoundText = dtccorrcompromiso.BoundText
End Sub

Private Sub dtccorrsol_Click(Area As Integer)
 dtcfechasol.BoundText = dtccorrsol.BoundText
End Sub

Private Sub dtcdenominacionruc_Click(Area As Integer)
    dtcnroruc.BoundText = dtcdenominacionruc.BoundText
End Sub

Private Sub dtcdenopuesto_Click(Area As Integer)
    dtccodpuesto.Text = dtcdenopuesto.BoundText
End Sub

Private Sub DtCDescripcion_Click(Area As Integer)
    DtCCodigo.BoundText = DtCDescripcion.BoundText
    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
    dtccodmanejo.BoundText = DtCDescripcion.BoundText
    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub dtcdescripar_Click(Area As Integer)
    dtccodpar.Text = dtcdescripar.BoundText
End Sub

Private Sub dtcdescripuni_Click(Area As Integer)
    dtccoduni.Text = dtcdescripuni.BoundText
End Sub

Private Sub dtcdescrtipoid_Click(Area As Integer)
    dtccodtipoid.BoundText = dtcdescrtipoid.BoundText
End Sub

Private Sub dtcfechacompromiso_Click(Area As Integer)
    dtccorrcompromiso.BoundText = dtcfechacompromiso.BoundText
End Sub

Private Sub dtcfechasol_Click(Area As Integer)
    dtccorrsol.BoundText = dtcfechasol.BoundText
End Sub

Private Sub dtcnroruc_Click(Area As Integer)
    dtcdenominacionruc.Text = dtcnroruc.BoundText
End Sub


Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub dtctipodoc_Click(Area As Integer)
    dtcdenodoc.Text = dtctipodoc.BoundText
End Sub

Private Sub dtcunidadmedida_Click(Area As Integer)
    DtCCodigo.BoundText = dtcunidadmedida.BoundText
    DtCDescripcion.BoundText = dtcunidadmedida.BoundText
    dtccodmanejo.BoundText = dtcunidadmedida.BoundText
    dtccodpeso.BoundText = dtcunidadmedida.BoundText
End Sub

Private Sub dtcdespoa_Click(Area As Integer)
    dtccodpoa.Text = dtcdespoa.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    dtc_aux5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub DTPFechaCobro02_LostFocus()
    If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
        MsgBox "La <<Fecha Cobranza2>> No puede ser MENOR a la <<Fecha Cobranza1>>, Vuelva a Intentar !! ", vbExclamation, "Atención!"
        DTPFechaCobro02.SetFocus
    End If
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = 0
    parametro = Aux
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral01_Click
    'Call ABRIR_TABLA
    'Call ABRIR_TABLA_AUX2
    'Call ABRIR_TABLA_DET3
    'txt_codigo.Enabled = True
    mbDataChanged = False
'    FrmCabecera.Enabled = False
    FrmCobros.Enabled = False
    FrmCobros1.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
'    FraGrabarCancelar1.Visible = False
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = True
    If glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
    End If
'    FrmEdita.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    FraNavega.Caption = lbl_titulo.Caption
    'lbl_titulo2.Caption = lbl_titulo.Caption
    'lbl_titulo1.Caption = lbl_titulo.Caption
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    Set rs_datos4 = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador en Fac.
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
    
    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    'rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario ", db, adOpenStatic  '4333735
    'rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
    dtc_desc4A1.BoundText = dtc_codigo4A1.BoundText

    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from gc_tipo_transaccion order by trans_descripcion", db, adOpenStatic
    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "ac_tipo_compra_venta", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh
    
    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    'rs_datos15.Open "select * from av_lista_productos where saldo_actual >= 0 order by DescDetalle ", db, adOpenKeyset, adLockReadOnly  'JQA 06/2008
    rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh
    
   'wwwwwwwwwwwwwwwwwwww
    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
    'Call ABREVENTAS
  
'    Set rs_Dsctos = New ADODB.Recordset
'    If rs_Dsctos.State = 1 Then rs_Dsctos.Close
'    rs_Dsctos.Open "select * from ac_ventas_descuentos ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    Set AdoDsctos.Recordset = rs_Dsctos
'    AdoDsctos.Refresh

    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
       
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    rs_datos20.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    dtc_ctades.BoundText = dtc_cta.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_cuenta_bancaria", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'    frmmo_formulario_M1.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre = rs_Personal!pers_nombres
'    frmmo_formulario_M1.Dtc_Pers_Cargo = rs_Personal!cargo_codigo
'  End If
'  If glPersNew = "L" Then
'    frmmo_formulario_M1.Dtc_doc_id_lab = rs_Personal!pers_doc_id
'    frmmo_formulario_M1.Dtc_pers_1apell_lab = rs_Personal!pers_primer_apellido
'    frmmo_formulario_M1.Dtc_pers_2apell_lab = rs_Personal!pers_segundo_apellido
'    frmmo_formulario_M1.Dtc_Pers_nombre_lab = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PL" Then
'    frmeo_Larvas_mosquitos.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_Larvas_mosquitos.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_Larvas_mosquitos.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_Larvas_mosquitos.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  If glPersNew = "PMA" Then
'    frmeo_mosquito_adulto.Dtc_pers_id = rs_Personal!pers_doc_id
'    frmeo_mosquito_adulto.Dtc_pers_1apell = rs_Personal!pers_primer_apellido
'    frmeo_mosquito_adulto.Dtc_pers_2Apell = rs_Personal!pers_segundo_apellido
'    frmeo_mosquito_adulto.Dtc_Pers_nombre = rs_Personal!pers_nombres
'  End If
'  glPersNew = "N"

End Sub

Private Sub OpMod1_Click()
'    Txt_modelo.Text = Txt_modelo1.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
'    End If
'    'Set ado_datos17.Recordset = rs_datos18
'    'ado_datos17.Refresh
End Sub

Private Sub OpMod2_Click()
'    Txt_modelo.Text = Txt_modelo2.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_h
'    End If
End Sub

Private Sub OpMod3_Click()
'    Txt_modelo.Text = Txt_modelo3.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs_x
'    End If
End Sub

Private Sub OptFilGral01_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    Set rs_datos01 = New Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    If glusuario = "ADMIN" Then
        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'REG' "
    Else
        If glusuario = "HBUSTILLOS" Then
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'REG' AND unidad_codigo = 'DVTA' "
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos01.Sort = "cobranza_fecha_prog"
    Set Ado_datos01.Recordset = rs_datos01.DataSource
    Set dg_datos1.DataSource = Ado_datos01.Recordset
    
End Sub

Private Sub OptFilGral02_Click()
'===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos01 = New Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    'queryinicial1 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
    If glusuario = "ADMIN" Then
        queryinicial = "select * From av_ventas_cobranza where estado_codigo_sol = 'APR' and estado_codigo_fac = 'REG'  "
        'queryinicial = "SELECT ao_ventas_cobranza.*, ao_ventas_cabecera.* FROM ao_ventas_cobranza INNER JOIN ao_ventas_cabecera ON ao_ventas_cobranza.venta_codigo = ao_ventas_cabecera.venta_codigo"
    Else
        If glusuario = "HBUSTILLOS" Then
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' and estado_codigo_fac = 'REG' and unidad_codigo = 'DVTA' "
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' and estado_codigo_fac = 'REG' and  beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
    rs_datos01.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos01.Sort = "cobranza_fecha_prog"
    Set Ado_datos01.Recordset = rs_datos01.DataSource
    Set dg_datos1.DataSource = Ado_datos01.Recordset
End Sub

Private Sub OptFilGral03_Click()
    '===== Proceso para filtrado de datos(registros Pendientes para Cobrar)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
'    If glusuario = "ADMIN" Then
'        queryinicial2 = "select * From av_ventas_cobranza  "
'    Else
    If glusuario = "HBUSTILLOS" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE unidad_codigo = 'DVTA' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND unidad_codigo = 'DVTA' "
    Else
        If glusuario = "RVEIZAGA" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' and doc_codigo_fac <> 'R-103' "
        Else
            If glusuario = "ADMIN" Then
                'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
                queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' "
            Else
                queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' and beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "'  "
            End If
        End If
    End If
'    End If
        'queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
'    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset
End Sub

Private Sub OptFilGral04_Click()
    '===== Proceso para filtrado general de datos(Todos los registros)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
    If glusuario = "HBUSTILLOS" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE unidad_codigo = 'DVTA' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND unidad_codigo = 'DVTA' "
        End If
        If glusuario = "RVEIZAGA" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' and doc_codigo_fac <> 'R-103' "
        End If
        If glusuario = "ADMIN" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' "
        End If
        
'    If glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Then
'        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR'  "
'    Else
'        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
'    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_fac"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset
End Sub

Private Sub OptFilGral05_Click()
'===== Proceso para filtrado de datos(registros Pendientes para Cobrar)
    Set rs_datos02 = New Recordset
    If rs_datos02.State = 1 Then rs_datos02.Close
        'queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG'  "
    If glusuario = "RVEIZAGA" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo = 'APR' and doc_codigo_fac <> 'R-103' "
    End If
        If glusuario = "HBUSTILLOS" Then
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo = 'APR' AND unidad_codigo = 'DVTA' "
        End If
    If glusuario = "ADMIN" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo = 'APR' "
    End If
    rs_datos02.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    rs_datos02.Sort = "cobranza_fecha_cobro"
    Set Ado_datos02.Recordset = rs_datos02.DataSource
    Set dg_datos2.DataSource = Ado_datos02.Recordset

End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    'Set Ado_datos9.Recordset = rs_datos9
    'dtc_desc1.BoundText = dtc_codigo1.BoundText
        Set rs_datos = New Recordset
        If rs_datos.State = 1 Then rs_datos.Close
        If glusuario = "RVEIZAGA" Then
            queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and doc_codigo_fac <> 'R-103' "      'ORDER BY cobranza_fecha_prog
        Else
            If glusuario = "HBUSTILLOS" Then
                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Then
                    queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' "      'ORDER BY cobranza_fecha_prog
                End If
            End If
        End If
        
        rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
        rs_datos.Sort = "cobranza_fecha_sol"
        Set Ado_datos.Recordset = rs_datos.DataSource
        Set dg_datos.DataSource = Ado_datos.Recordset
'    If glusuario = "HBUSTILLOS" Then
'        Set rs_datos = New Recordset
'        If rs_datos.State = 1 Then rs_datos.Close
'        'queryinicial = "select * From ao_ventas_cobranza WHERE estado_codigo = 'REG' ORDER BY cobranza_fecha_prog "
'        'If glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Then
'        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
'        'Else
'        '    queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
'        'End If
'        'queryinicial = "Select * from ao_solicitud where " + parametro
'        rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'        rs_datos.Sort = "cobranza_fecha_sol"
'        Set Ado_datos.Recordset = rs_datos.DataSource
'        Set dg_datos.DataSource = Ado_datos.Recordset
'    End If
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If glusuario = "RVEIZAGA" Then
        queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' "     'ORDER BY cobranza_fecha_prog
    Else
        If glusuario = "HBUSTILLOS" Then
                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR'  AND estado_codigo_bco = 'REG' and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Then
                    queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' "      'ORDER BY cobranza_fecha_prog
                End If
            End If
    '    queryinicial = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
    End If
    'queryinicial = "select * From ao_ventas_cobranza  ORDER BY cobranza_fecha_prog "
    rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "cobranza_fecha_sol"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub Option1_Click()
'    Fra_Total.Visible = True
'End Sub
'
'Private Sub Option2_Click()
'    FrmCobranza.Visible = True
'End Sub

Private Sub TxtCantPedi_KeyPress(KeyAscii As Integer)
 If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtcaracteristicas_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares_contra.Text = 0
    End If
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
    Else
       Txtmonto_dolares.Text = 0
    End If
  End If

End Sub

Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos_contra.Text = 0
    End If
  End If
End Sub

Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(TxtTipo_cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
    Else
      TxtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub Txtobservaciones_KeyPress(KeyAscii As Integer)
    'convertir a mayusculas
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtsolpeso_KeyPress(KeyAscii As Integer)
'solo numeros y , .
    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
    
    Else
      KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub txtterref_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Then
        KeyAscii = Asc(UCase(Chr(0)))
    Else
        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            KeyAscii = Asc(UCase(Chr(0)))
            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
        End If
    End If
End Sub

Private Sub cerea()
  txt_venta = " "
  dtc_codigo4.Text = " "
  Dtcpaternosol.Text = " "  'dtc_codigo4.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  dtc_codigo2 = ""
  dtc_desc2 = ""
  txtTDC.Text = GlTipoCambioMercado ' GlTipoCambioOficial
  
'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txt_venta = ""
'  txtterref = ""
End Sub
'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub

Sub creaVista()
db.Execute "drop view vwF04"

db.Execute "create view vwF04 as " & _
            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipoben_codigo, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
            "from ao_solicitud_lista  ,     " & _
                 "ao_solicitud       ,     " & _
                 "fc_unidad_ejecutora,     " & _
                 "rc_personal,             " & _
                 "ac_bienes                " & _
            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
                    "ao_solicitud_lista.codigo_unidad    = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
                    "ao_solicitud_lista.codigo_solicitud =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipoben_codigo, " & _
            "ao_solicitud.codigo_solicitud, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.razon_s, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, " & _
            "ao_solicitud_lista.telefono, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.nacional_extranjero, ao_solicitud.por_tiempo, ao_solicitud.codigo_bien, ac_bienes.bie_descripcion, ao_solicitud.duracion_estimada_numero, ao_solicitud.duracion_estimada_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.esparaRH, ao_solicitud.tr_adjuntos, ao_solicitud.observaciones, ao_solicitud.caracteristicas, fc_unidad_ejecutora.Uni_descripcion_larga, ao_solicitud.fecha_solicitud, (rc_personal.paterno)+' '+(rc_personal.materno)+' '+(rc_personal.nombres)+' ['+ao_solicitud.ci+']', ao_solicitud_lista.id_beneficiario "
                 
'            "trim$(rc_personal.paterno) + ' ' + trim$(rc_personal.materno) + ' ' +trim$(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _

'''db.Execute "create view vwF05 as " & _
'''            "select  ao_solicitud_lista.* " & _
'''            "from ao_solicitud_lista"
End Sub

Sub CREAVISTAF11()
db.Execute "drop view VWF11"
db.Execute "create view VWF11 as " & _
    "SELECT ao_Solicitud.Ges_Gestion, ao_Solicitud.codigo_unidad, " & _
    "ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, " & _
    "ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, " & _
    "ao_Solicitud.fecha_solicitud, ao_Solicitud.codigo_bien, " & _
    "ac_bienes_grupo.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
"FROM ao_Solicitud, ao_Solicitud_detalle, ac_bienes_grupo, RC_Personal, ao_solicitud_lista " & _
"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    "(ao_Solicitud.codigo_unidad) = '" & Me.Ado_datos.Recordset!codigo_unidad & "' and " & _
    "(ao_Solicitud.codigo_solicitud) =  " & Me.Ado_datos.Recordset!codigo_solicitud & " and " & _
    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
    "ao_Solicitud.CodGrupo = ac_bienes_grupo.CodGrupo AND " & _
    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(ges, Nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  Set rs_datos19 = New ADODB.Recordset
  If rs_datos19.State = 1 Then rs_datos19.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!CANTOT
  End If
  
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If
  
  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & Nro & " "
  
'  TxtMontoBs.Text = VAR_AUX
'  TxtCobrado.Text = Cobrobs
'  TxtBstotal.Text = VAR_Bs
  
  If rstacumdet.State = 1 Then rstacumdet.Close
  
End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            lbl_titulo1.Caption = SSTab1.Caption
'            lbl_titulo3.Caption = SSTab1.Caption
            FraNavega1.Caption = SSTab1.Caption
'            FraGrabarCancelar1.Visible = False
            OptFilGral01.Value = True
            Call OptFilGral01_Click
        Case 1
            If glusuario = "RVEIZAGA" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Then
                lbl_titulo = SSTab1.Caption
'                lbl_titulo2 = SSTab1.Caption
                FraNavega.Caption = SSTab1.Caption
'                FraGrabarCancelar.Visible = False
                Call ABRIR_TABLAS_AUX
                OptFilGral1.Value = True
                Call OptFilGral1_Click
                'FACTURA O RECIBO
            Else
                SSTab1.Tab = 0
            End If
            Picture1.Visible = True
        Case 2
            lbl_titulo2 = SSTab1.Caption
'            lbl_titulo5 = SSTab1.Caption
            FraNavega2.Caption = SSTab1.Caption
'            FraGrabarCancelar2.Visible = False
            OptFilGral03.Value = True
            Call OptFilGral03_Click
    End Select
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtCantidad_LostFocus()
  If (TxtCantidad.Text) = "" Then
    TxtCantidad.Text = 1
  End If
  If dtc_codigo11.Text = "E" Then
    If (dtc_codigo12.Text) = "" Or IsNull(dtc_codigo12.Text) Then
        TxtDescuento.Text = "0"
    Else
        TxtDescuento.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) * CDbl(Dtc_aux12.Text))
    End If
    'TxtPrecioU.Text = dtc_precioventabase15.Text
    'TxtTotal.Text = CDbl(TxtCantidad.Text) * (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento.Text))
  End If
  If dtc_codigo11.Text = "C" Then
     TxtDescuento.Text = "0"
     'TxtDescuento.Text = CDbl(Dtc_aux12) * (CDbl(TxtCantidad) * CDbl(TxtPrecioU))
     TxtPrecioU.Text = dtc_precioventafinal15.Text
  End If
  If (dtc_codigo11.Text <> "E" And dtc_codigo11.Text <> "C") Then
     TxtDescuento.Text = "0"
     TxtPrecioU.Text = "0"
  End If
  TxtTotal.Text = (CDbl(TxtCantidad.Text) * CDbl(TxtPrecioU.Text)) - CDbl(TxtDescuento.Text)
  
End Sub

Private Sub TxtCobrado_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtDscto2_LostFocus()
    TxtDscto2D.Text = Round(CDbl(TxtDscto2.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtDscto2D_LostFocus()
    TxtDscto2.Text = Round(CDbl(TxtDscto2D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtMontoDol = "0"
    Else
        'TxtMontoDol = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
        TxtMontoDol = Round(CDbl(TxtMonto.Text) / CDbl(txt_tdc), 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub
'adelante
Private Function CodigoControl(NAuto As String, NFactura As String, Nit As String, Fecha As String, Monto As String, Key As String) As String
Dim Suma As Currency
Dim CodControl As String, Cadena As String, NroVer As String
Dim Pos As Integer, i As Integer, Nro As Integer, j As Integer
Dim SumTot As Long, SumPar(1 To 5) As Currency

  Suma = 0
  Cadena = NFactura
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NFactura = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox NFactura
  'Para el Nit o CI del Cliente.
  Cadena = Nit
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Nit = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Nit
  'Para la Fecha de transaccion.
  Cadena = Fecha
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Fecha = Cadena
  Suma = Suma + CDbl(Cadena)
  'MsgBox Fecha
  'Para el monto de transaccion.
  Cadena = Monto
  For i = 1 To 2
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  Monto = Cadena
  'MsgBox Monto
  Suma = Suma + CDbl(Cadena)
  'MsgBox Suma
  
  'Para Obtener los 5 numeros Verhoeff.
  Cadena = Str(Suma)
  For i = 1 To 5
    Cadena = Cadena & Verhoeff(Cadena)
  Next i
  NroVer = Right(Cadena, 5)
  'MsgBox NroVer
  
  'Para obtener las nuevas cadenas.
  Cadena = ""
  Pos = 1
  For i = 1 To 5
    Nro = (Val(Mid(NroVer, i, 1)) + 1)
    Select Case i
      Case 1: Cadena = Cadena & NAuto & Mid(Key, Pos, Nro)
      Case 2: Cadena = Cadena & NFactura & Mid(Key, Pos, Nro)
      Case 3: Cadena = Cadena & Nit & Mid(Key, Pos, Nro)
      Case 4: Cadena = Cadena & Fecha & Mid(Key, Pos, Nro)
      Case 5: Cadena = Cadena & Monto & Mid(Key, Pos, Nro)
    End Select
    Pos = Pos + Nro
  Next i

  Cadena = AllegedRC4(Cadena, (Key & NroVer))

  
  SumTot = 0
  i = 0
  Do While i < Len(Cadena)
    i = i + 1
    SumTot = SumTot + Asc(Mid(Trim(Cadena), i, 1))
  Loop
 
  
  For i = 1 To 5
    SumPar(i) = 0
    j = i
    Do While j <= Len(Cadena)
      SumPar(i) = SumPar(i) + Asc(Mid(Cadena, j, 1))
      j = j + 5
    Loop
  
  Next i
  
  Suma = 0
  For i = 1 To 5
    SumPar(i) = Int((SumTot * SumPar(i)) / (Val(Mid(NroVer, i, 1)) + 1))
    Suma = Suma + SumPar(i)
  Next i
  Cadena = Base64(Str(Suma))
  
  Cadena = AllegedRC4(Cadena, (Key & NroVer))
  

  CodigoControl = ""
  i = 0
  j = 1
  
  Do While i < Len(Cadena)
    i = i + 1
    If i Mod 2 = 0 Then
      CodigoControl = CodigoControl & Mid(Cadena, j, 2) & "-"
      j = i + 1
    End If
  Loop
  
  CodigoControl = Mid(CodigoControl, 1, (Len(CodigoControl) - 1))
End Function
Public Function Redondear(dNumero As Double, iDecimales As Integer) As Double
    Dim lMultiplicador As Long
    Dim dRetorno As Double
    
    If iDecimales > 9 Then iDecimales = 9
    lMultiplicador = 10 ^ iDecimales
    dRetorno = CDbl(CLng(dNumero * lMultiplicador)) / lMultiplicador
    
    Redondear = dRetorno
End Function
Private Function Redondeo(ByVal Numero, ByVal Decimales)
      Redondeo = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Private Sub TxtMonto02_LostFocus()
    TxtMonto02D.Text = Round(CDbl(TxtMonto02.Text) / Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMonto02D_LostFocus()
    TxtMonto02.Text = Round(CDbl(TxtMonto02D.Text) * Ado_datos02.Recordset!cobranza_tdc, 2)
End Sub

Private Sub TxtMontoDol_Change()
    'TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
End Sub

Private Sub TxtMontoDol_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoDol_LostFocus()
    TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
End Sub

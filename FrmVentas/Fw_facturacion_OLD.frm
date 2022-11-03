VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Fw_facturacion 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Financiero - Tesoreria - Facturacion Electrónica"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   17115
   Icon            =   "Fw_facturacion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.97602e6
   ScaleMode       =   0  'User
   ScaleWidth      =   2.25671e7
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BtnModDetalle2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Lista de Facturas"
      Height          =   645
      Left            =   570
      Picture         =   "Fw_facturacion.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   141
      ToolTipText     =   "Ver Detalle del Bien ..."
      Top             =   8280
      Width           =   1005
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      Picture         =   "Fw_facturacion.frx":1404
      ScaleHeight     =   1755
      ScaleMode       =   0  'User
      ScaleWidth      =   1875
      TabIndex        =   39
      Top             =   8040
      Width           =   1935
      Begin VB.CommandButton BntImprimir2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobranzas"
         Height          =   645
         Left            =   405
         Picture         =   "Fw_facturacion.frx":6D436
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   960
         Width           =   1005
      End
      Begin VB.CommandButton BntImprimir3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cobranzas Dolares"
         Height          =   645
         Left            =   960
         Picture         =   "Fw_facturacion.frx":6D9F3
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1515
      Left            =   120
      Picture         =   "Fw_facturacion.frx":6DFB0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   38
      Top             =   6405
      Width           =   1935
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kardex Bs."
         Height          =   675
         Left            =   420
         Picture         =   "Fw_facturacion.frx":D9FE2
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   750
         Width           =   1050
      End
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kardex.Dol."
         Height          =   645
         Left            =   840
         Picture         =   "Fw_facturacion.frx":DA59F
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   870
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton BtnImprimir4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cronograma"
         Height          =   675
         Left            =   420
         Picture         =   "Fw_facturacion.frx":DAB5C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Imprime Cronograma de Cobranzas ..."
         Top             =   45
         Width           =   1050
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6270
      Left            =   120
      TabIndex        =   35
      Top             =   45
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   11060
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Fw_facturacion.frx":DC2DE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FraNavega1"
      Tab(0).Control(1)=   "FrmCobros1"
      Tab(0).Control(2)=   "fraOpciones1"
      Tab(0).Control(3)=   "FraGrabarCancelar1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "PARA FACTURACION"
      TabPicture(1)   =   "Fw_facturacion.frx":DCCF0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraNavega"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraOpciones"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraGrabarCancelar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FrmCobros"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frm_benef"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Fw_facturacion.frx":106912
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraGrabarCancelar2"
      Tab(2).Control(1)=   "fraOpciones2"
      Tab(2).Control(2)=   "FrmCobros2"
      Tab(2).Control(3)=   "FraNavega2"
      Tab(2).ControlCount=   4
      Begin VB.Frame frm_benef 
         BackColor       =   &H00404040&
         Caption         =   "Elije un Nuevo Beneficiario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   2655
         Left            =   5880
         TabIndex        =   78
         Top             =   1485
         Visible         =   0   'False
         Width           =   10335
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8760
            TabIndex        =   79
            Top             =   975
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton BtnCancelarBen 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   5280
            MaskColor       =   &H00000000&
            Picture         =   "Fw_facturacion.frx":107324
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cancelar"
            Top             =   1680
            Width           =   1125
         End
         Begin VB.CommandButton BtnGrabarBen 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Aceptar"
            Height          =   675
            Left            =   3480
            Picture         =   "Fw_facturacion.frx":10752E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1680
            Width           =   1125
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "Fw_facturacion.frx":107738
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6720
            TabIndex        =   80
            Top             =   1320
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "Fw_facturacion.frx":107751
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   480
            TabIndex        =   9
            Top             =   960
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   14737632
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux8 
            Bindings        =   "Fw_facturacion.frx":10776A
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6720
            TabIndex        =   140
            Top             =   960
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   4210752
            ForeColor       =   16777215
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "00000000000004"
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de:"
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
            Left            =   480
            TabIndex        =   81
            Top             =   585
            Width           =   2025
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
         Left            =   5805
         TabIndex        =   161
         Top             =   1440
         Width           =   10575
         Begin VB.CommandButton cmd_benef 
            BackColor       =   &H00808000&
            Height          =   320
            Left            =   9705
            Picture         =   "Fw_facturacion.frx":107783
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "Buscar Beneficiario"
            Top             =   3150
            Width           =   375
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
            Left            =   6915
            Locked          =   -1  'True
            TabIndex        =   175
            Text            =   "0"
            Top             =   3570
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.TextBox TxtObs 
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos"
            Height          =   585
            Left            =   1440
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   174
            Top             =   1740
            Width           =   6255
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
            Left            =   6525
            TabIndex        =   173
            Text            =   "0"
            Top             =   4200
            Width           =   1275
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   9300
            TabIndex        =   172
            Top             =   3165
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4040
            TabIndex        =   171
            Top             =   3585
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2475
            TabIndex        =   170
            Top             =   4090
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4040
            TabIndex        =   169
            Top             =   4290
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7080
            TabIndex        =   167
            Top             =   2610
            Width           =   255
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
            Left            =   3480
            TabIndex        =   166
            Text            =   "6.96"
            Top             =   4200
            Width           =   915
         End
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
            Left            =   4845
            TabIndex        =   165
            Text            =   "0"
            Top             =   4200
            Width           =   1275
         End
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   1500
            Left            =   5520
            ScaleHeight     =   1500
            ScaleMode       =   0  'User
            ScaleWidth      =   1411.765
            TabIndex        =   164
            Top             =   1080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox TxtCmpbte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            DataField       =   "cobranza_nro_factura"
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
            Height          =   420
            Left            =   3720
            TabIndex        =   163
            Text            =   "0"
            Top             =   1060
            Width           =   1455
         End
         Begin VB.TextBox lbl_docnro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            DataSource      =   "Ado_datos"
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
            Height          =   420
            Left            =   1800
            TabIndex        =   162
            Text            =   "0"
            Top             =   1065
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            Bindings        =   "Fw_facturacion.frx":10790D
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5760
            TabIndex        =   168
            Top             =   2595
            Width           =   1590
            _ExtentX        =   2805
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
         Begin MSDataListLib.DataCombo dtc_desc5 
            Bindings        =   "Fw_facturacion.frx":107926
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2085
            TabIndex        =   177
            Top             =   3150
            Width           =   5160
            _ExtentX        =   9102
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
            Bindings        =   "Fw_facturacion.frx":10793F
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7395
            TabIndex        =   178
            Top             =   3150
            Width           =   2190
            _ExtentX        =   3863
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
            Bindings        =   "Fw_facturacion.frx":107958
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1815
            TabIndex        =   179
            Top             =   2595
            Width           =   4320
            _ExtentX        =   7620
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
            Bindings        =   "Fw_facturacion.frx":107971
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   300
            TabIndex        =   180
            Top             =   4080
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
         Begin MSDataListLib.DataCombo dtc_ctades 
            Bindings        =   "Fw_facturacion.frx":10798B
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2640
            TabIndex        =   181
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
            Bindings        =   "Fw_facturacion.frx":1079A5
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   182
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
            Bindings        =   "Fw_facturacion.frx":1079BE
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   183
            Top             =   3570
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
            ListField       =   "trans_descripcion"
            BoundColumn     =   "trans_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo5 
            Bindings        =   "Fw_facturacion.frx":1079D7
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4920
            TabIndex        =   184
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
            Bindings        =   "Fw_facturacion.frx":1079F0
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   185
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
            Bindings        =   "Fw_facturacion.frx":107A09
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2640
            TabIndex        =   186
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
            Left            =   8400
            TabIndex        =   187
            Top             =   2295
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   92471297
            CurrentDate     =   42873
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "cobranza_fecha_fac"
            DataSource      =   "Ado_datos"
            Height          =   300
            Left            =   8400
            TabIndex        =   188
            Top             =   2280
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
            Format          =   92471297
            CurrentDate     =   41678
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
            TabIndex        =   211
            Top             =   4395
            Visible         =   0   'False
            Width           =   1275
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
            Left            =   8460
            TabIndex        =   210
            Top             =   975
            Visible         =   0   'False
            Width           =   1710
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
            Left            =   8400
            TabIndex        =   209
            Top             =   720
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00FFFF80&
            X1              =   8055
            X2              =   10560
            Y1              =   1605
            Y2              =   1605
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   10560
            Y1              =   3705
            Y2              =   3705
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
            TabIndex        =   208
            Top             =   3120
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   8460
            TabIndex        =   207
            Top             =   465
            Width           =   1710
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
            Left            =   8460
            TabIndex        =   206
            Top             =   1680
            Visible         =   0   'False
            Width           =   1710
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
            Left            =   4320
            TabIndex        =   205
            Top             =   255
            Width           =   1125
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   8055
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   8055
            X2              =   8055
            Y1              =   0
            Y2              =   2855
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
            Left            =   8400
            TabIndex        =   204
            Top             =   1410
            Visible         =   0   'False
            Width           =   1785
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
            TabIndex        =   203
            Top             =   1060
            Width           =   1005
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
            Left            =   6840
            TabIndex        =   202
            Top             =   255
            Width           =   855
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
            TabIndex        =   201
            Top             =   1875
            Width           =   960
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
            Left            =   8400
            TabIndex        =   200
            Top             =   195
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
            TabIndex        =   199
            Top             =   2625
            Width           =   1560
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
            Left            =   8400
            TabIndex        =   198
            Top             =   1920
            Width           =   1785
         End
         Begin VB.Label lbl_factura 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Registro       Nro. de Recibo            Nro.de Factura              Nro. de Autorización"
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
            Left            =   300
            TabIndex        =   197
            Top             =   795
            Width           =   7215
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   8055
            Y1              =   1605
            Y2              =   1605
         End
         Begin VB.Label Lbl_nombre_fac 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura a Nombre de:"
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
            TabIndex        =   196
            Top             =   3165
            Width           =   1950
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
            TabIndex        =   195
            Top             =   255
            Width           =   1245
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
            Left            =   5520
            TabIndex        =   194
            Top             =   1080
            Width           =   2235
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                         Nro.Venta:                        Nro.Cuota:"
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
            TabIndex        =   193
            Top             =   240
            Width           =   6630
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
            Left            =   1515
            TabIndex        =   192
            Top             =   1020
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "."
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
            TabIndex        =   191
            Top             =   3600
            Visible         =   0   'False
            Width           =   6705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "                                                                      Tipo.Cambio             BOB (Bs)                   USD (Dol)"
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
            TabIndex        =   190
            Top             =   3840
            Width           =   7500
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
            TabIndex        =   189
            Top             =   4455
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.PictureBox FraGrabarCancelar1 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   1020
         Left            =   -74960
         Picture         =   "Fw_facturacion.frx":107A22
         ScaleHeight     =   960
         ScaleWidth      =   5805
         TabIndex        =   93
         Top             =   360
         Width           =   5865
         Begin VB.CommandButton BtnCancelar1 
            BackColor       =   &H00808000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   3240
            MaskColor       =   &H00000000&
            Picture         =   "Fw_facturacion.frx":173A54
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnGrabar1 
            BackColor       =   &H00808000&
            Caption         =   "Grabar"
            Height          =   675
            Left            =   1560
            Picture         =   "Fw_facturacion.frx":173C5E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox fraOpciones1 
         BackColor       =   &H00404040&
         Height          =   1020
         Left            =   -74960
         Picture         =   "Fw_facturacion.frx":173E68
         ScaleHeight     =   960
         ScaleWidth      =   16320
         TabIndex        =   90
         Top             =   360
         Width           =   16380
         Begin VB.CommandButton BtnVer 
            BackColor       =   &H00808000&
            Caption         =   "REF"
            Height          =   315
            Left            =   14640
            Picture         =   "Fw_facturacion.frx":1DFE9A
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00808000&
            Caption         =   "REC"
            Height          =   315
            Left            =   15240
            Picture         =   "Fw_facturacion.frx":1E02DC
            Style           =   1  'Graphical
            TabIndex        =   158
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00808000&
            Caption         =   "DEI"
            Height          =   315
            Left            =   14040
            Picture         =   "Fw_facturacion.frx":1E071E
            Style           =   1  'Graphical
            TabIndex        =   157
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00808000&
            Caption         =   "Correl"
            Height          =   315
            Left            =   15840
            Picture         =   "Fw_facturacion.frx":1E0B60
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Guarda en Archivo Digital"
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton BtnAprobar1 
            BackColor       =   &H00808000&
            Caption         =   "Aprobar"
            Height          =   720
            Left            =   1920
            Picture         =   "Fw_facturacion.frx":1E0FA2
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Aprueba Envia a Facturación"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnBuscar1 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   2760
            Picture         =   "Fw_facturacion.frx":1E11AC
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir1 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   4680
            Picture         =   "Fw_facturacion.frx":1E1764
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cerrar Ventana"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00808000&
            Caption         =   "Anular"
            Height          =   720
            Left            =   1080
            Picture         =   "Fw_facturacion.frx":1E196E
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Anula Registro Activo"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnModificar1 
            BackColor       =   &H00808000&
            Caption         =   "Modificar"
            Height          =   720
            Left            =   240
            Picture         =   "Fw_facturacion.frx":1E2638
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Modifica Registro Activo"
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
            TabIndex        =   92
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.PictureBox FraGrabarCancelar2 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   1020
         Left            =   -74955
         Picture         =   "Fw_facturacion.frx":1E2C18
         ScaleHeight     =   960
         ScaleWidth      =   6000
         TabIndex        =   97
         Top             =   360
         Width           =   6060
         Begin VB.CommandButton BtnGrabar2 
            BackColor       =   &H00808000&
            Caption         =   "Grabar"
            Height          =   675
            Left            =   1560
            Picture         =   "Fw_facturacion.frx":24EC4A
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnCancelar2 
            BackColor       =   &H00808000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   3240
            MaskColor       =   &H00000000&
            Picture         =   "Fw_facturacion.frx":24EE54
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox fraOpciones2 
         BackColor       =   &H00404040&
         Height          =   1020
         Left            =   -74960
         Picture         =   "Fw_facturacion.frx":24F05E
         ScaleHeight     =   960
         ScaleWidth      =   16440
         TabIndex        =   98
         Top             =   360
         Width           =   16500
         Begin VB.CommandButton BtnAprobar 
            BackColor       =   &H00808000&
            Caption         =   "Aprobar"
            Height          =   720
            Left            =   2760
            Picture         =   "Fw_facturacion.frx":2BB090
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Aprueba Fin de la Cobranza"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnBuscar2 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   3600
            Picture         =   "Fw_facturacion.frx":2BB29A
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnModificar2 
            BackColor       =   &H00808000&
            Caption         =   "Modificar"
            Height          =   720
            Left            =   240
            Picture         =   "Fw_facturacion.frx":2BB852
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00808000&
            Caption         =   "Anular"
            Height          =   720
            Left            =   1080
            Picture         =   "Fw_facturacion.frx":2BBE32
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Anula Registro Activo"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir2 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   5280
            Picture         =   "Fw_facturacion.frx":2BCAFC
            Style           =   1  'Graphical
            TabIndex        =   29
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
            TabIndex        =   100
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.Frame FrmCobros2 
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
         Height          =   4780
         Left            =   -68115
         TabIndex        =   105
         Top             =   1440
         Width           =   9615
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
            Left            =   7680
            TabIndex        =   148
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
            TabIndex        =   145
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
            TabIndex        =   144
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
            TabIndex        =   139
            Top             =   2610
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "Fw_facturacion.frx":2BCD06
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2085
            TabIndex        =   18
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
            ItemData        =   "Fw_facturacion.frx":2BCD1F
            Left            =   120
            List            =   "Fw_facturacion.frx":2BCD2C
            TabIndex        =   23
            Text            =   "BOB"
            Top             =   4360
            Width           =   735
         End
         Begin VB.ComboBox cmd_moneda1 
            DataField       =   "tipo_moneda"
            DataSource      =   "Ado_datos02"
            Height          =   315
            ItemData        =   "Fw_facturacion.frx":2BCD3F
            Left            =   120
            List            =   "Fw_facturacion.frx":2BCD4C
            TabIndex        =   19
            Text            =   "BOB"
            Top             =   3740
            Width           =   735
         End
         Begin MSDataListLib.DataCombo DataCombo14 
            Bindings        =   "Fw_facturacion.frx":2BCD5F
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   5085
            TabIndex        =   137
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
            Left            =   6100
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   21
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
            TabIndex        =   108
            Top             =   1140
            Width           =   6015
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
            Left            =   6100
            TabIndex        =   22
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
            TabIndex        =   106
            Top             =   1890
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Fw_facturacion.frx":2BCD78
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   5535
            TabIndex        =   107
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
            Bindings        =   "Fw_facturacion.frx":2BCD92
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   915
            TabIndex        =   20
            Top             =   3740
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
            Bindings        =   "Fw_facturacion.frx":2BCDAC
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2640
            TabIndex        =   109
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
            Bindings        =   "Fw_facturacion.frx":2BCDC6
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   915
            TabIndex        =   24
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
            Bindings        =   "Fw_facturacion.frx":2BCDDF
            DataField       =   "trans_codigo"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   3480
            TabIndex        =   110
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
            Bindings        =   "Fw_facturacion.frx":2BCDF8
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   2640
            TabIndex        =   111
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
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "Fw_facturacion.frx":2BCE11
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   1680
            TabIndex        =   112
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
            Bindings        =   "Fw_facturacion.frx":2BCE2B
            DataField       =   "beneficiario_codigo_fac"
            DataSource      =   "Ado_datos02"
            Height          =   315
            Left            =   120
            TabIndex        =   136
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
            Left            =   7800
            TabIndex        =   17
            Top             =   4365
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   92471297
            CurrentDate     =   42955
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro2A 
            Height          =   300
            Left            =   7800
            TabIndex        =   154
            Top             =   3720
            Visible         =   0   'False
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   92471297
            CurrentDate     =   42884
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro2 
            DataField       =   "cobranza_fecha_cobro1"
            DataSource      =   "Ado_datos02"
            Height          =   300
            Left            =   7800
            TabIndex        =   155
            Top             =   3720
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   92471297
            CurrentDate     =   42885
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   7380
            X2              =   7380
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
            TabIndex        =   135
            Top             =   645
            Width           =   1995
         End
         Begin VB.Label TxtCmpbte2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_nro_factura"
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
            Left            =   2700
            TabIndex        =   134
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
            TabIndex        =   133
            Top             =   135
            Width           =   1245
         End
         Begin VB.Label Lbl_nombre_fac3 
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label lbl_factura3 
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
            ForeColor       =   &H00C0FFFF&
            Height          =   480
            Left            =   1980
            TabIndex        =   128
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
            Left            =   7560
            TabIndex        =   127
            Top             =   2805
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
            TabIndex        =   126
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
            Left            =   7560
            TabIndex        =   125
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
            TabIndex        =   124
            Top             =   1275
            Width           =   960
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cobranza:                      Nro.Venta:                     Nro.Cuota:"
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
            TabIndex        =   123
            Top             =   150
            Width           =   6270
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
            Left            =   6360
            TabIndex        =   122
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
            TabIndex        =   121
            Top             =   645
            Width           =   1005
         End
         Begin VB.Label lbl_fechas3 
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
            Left            =   7560
            TabIndex        =   120
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
            Left            =   4080
            TabIndex        =   119
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
            Left            =   7620
            TabIndex        =   118
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
            Left            =   7620
            TabIndex        =   117
            Top             =   465
            Width           =   1710
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda  Cuenta Bancaria 1                  Cobrado1 Bs   Cobrado1 Dol   Cpbte.Deposito1    Fecha.Cobro1"
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
            TabIndex        =   116
            Top             =   3480
            Width           =   8970
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7380
            Y1              =   3030
            Y2              =   3030
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda  Cuenta Bancaria 2                  Cobrado2 Bs   Cobrado2 Dol   Cpbte.Deposito2    Fecha.Cobro2"
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
            Top             =   4095
            Width           =   8970
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
            Left            =   7560
            TabIndex        =   114
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
            Left            =   7620
            TabIndex        =   113
            Top             =   1350
            Width           =   1710
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   1020
         Left            =   0
         Picture         =   "Fw_facturacion.frx":2BCE44
         ScaleHeight     =   960
         ScaleWidth      =   5880
         TabIndex        =   89
         Top             =   360
         Visible         =   0   'False
         Width           =   5940
         Begin VB.CommandButton BtnCancelar 
            BackColor       =   &H00808000&
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   3240
            MaskColor       =   &H00000000&
            Picture         =   "Fw_facturacion.frx":328E76
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancelar"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnGrabar 
            BackColor       =   &H00808000&
            Caption         =   "Grabar"
            Height          =   675
            Left            =   1560
            Picture         =   "Fw_facturacion.frx":329080
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame FraNavega2 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4805
         Left            =   -74960
         TabIndex        =   101
         Top             =   1440
         Width           =   6825
         Begin VB.OptionButton OptFilGral05 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cobrados y No Contabilizados"
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
            Left            =   3240
            TabIndex        =   138
            Top             =   4480
            Width           =   2775
         End
         Begin VB.OptionButton OptFilGral04 
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
            Left            =   3720
            TabIndex        =   103
            Top             =   4480
            Visible         =   0   'False
            Width           =   1275
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
            TabIndex        =   102
            Top             =   4480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dg_datos2 
            Bindings        =   "Fw_facturacion.frx":32928A
            Height          =   4095
            Left            =   75
            TabIndex        =   104
            Top             =   240
            Width           =   6660
            _ExtentX        =   11748
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
               DataField       =   "cobranza_deuda_bs"
               Caption         =   "Cobrado1.Bs."
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
               DataField       =   "cobranza_deuda_bs2"
               Caption         =   "Cobrado2.Bs"
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
               DataField       =   "estado_codigo_bco"
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
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
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
            Top             =   4395
            Width           =   6675
            _ExtentX        =   11774
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
         Left            =   40
         Picture         =   "Fw_facturacion.frx":3292A4
         ScaleHeight     =   960
         ScaleWidth      =   16320
         TabIndex        =   82
         Top             =   360
         Width           =   16380
         Begin VB.CommandButton BtnAprobar2 
            BackColor       =   &H00808000&
            Caption         =   "Enviar"
            Height          =   720
            Left            =   2640
            Picture         =   "Fw_facturacion.frx":3952D6
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Verifica la Cobranza"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimir5 
            BackColor       =   &H00C0C000&
            Caption         =   "Re-Impr"
            Height          =   720
            Left            =   7635
            Picture         =   "Fw_facturacion.frx":3954E0
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Re-Imprime Factura"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimir2 
            BackColor       =   &H00C0C000&
            Caption         =   "Recibo"
            Height          =   720
            Left            =   6720
            Picture         =   "Fw_facturacion.frx":396C62
            Style           =   1  'Graphical
            TabIndex        =   149
            ToolTipText     =   "Imprime Recibo"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton BtnModificar 
            BackColor       =   &H00808000&
            Caption         =   "Modificar"
            Height          =   720
            Left            =   120
            Picture         =   "Fw_facturacion.frx":3983E4
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Modifica Registro Activo"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnImprimir3 
            BackColor       =   &H00C0C000&
            Caption         =   "Subir.Fac"
            Height          =   720
            Left            =   5880
            Picture         =   "Fw_facturacion.frx":3989C4
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Imprime Factura"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnBuscar 
            BackColor       =   &H00808000&
            Caption         =   "Buscar"
            Height          =   720
            Left            =   1800
            Picture         =   "Fw_facturacion.frx":39A146
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdFoto 
            BackColor       =   &H00C0C000&
            Caption         =   "&Reporte"
            Height          =   720
            Left            =   4680
            Picture         =   "Fw_facturacion.frx":39A6FE
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Enviar a Facturación"
            Top             =   120
            Width           =   740
         End
         Begin VB.CommandButton BtnDesAprobar 
            BackColor       =   &H00808000&
            Caption         =   "Devolver"
            Height          =   720
            Left            =   3465
            Picture         =   "Fw_facturacion.frx":39B100
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Devuelve a Cobradores (Solicitud de Facturación)"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnSalir 
            BackColor       =   &H00808000&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   15360
            Picture         =   "Fw_facturacion.frx":39B30A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cerrar Ventana"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnEliminar 
            BackColor       =   &H00808000&
            Caption         =   "Anular"
            Height          =   720
            Left            =   960
            Picture         =   "Fw_facturacion.frx":39B514
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Anula Factura Emitida"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton BtnAñadir 
            BackColor       =   &H00808000&
            Caption         =   "Nuevo"
            Height          =   720
            Left            =   120
            Picture         =   "Fw_facturacion.frx":39C1DE
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Nuevo Registro"
            Top             =   120
            Visible         =   0   'False
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
            TabIndex        =   88
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
         Left            =   -69200
         TabIndex        =   50
         Top             =   1440
         Width           =   10575
         Begin VB.ComboBox cmd_fac 
            Height          =   315
            ItemData        =   "Fw_facturacion.frx":39C802
            Left            =   2880
            List            =   "Fw_facturacion.frx":39C80C
            TabIndex        =   143
            Text            =   "FACTURA"
            Top             =   3075
            Width           =   1275
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6975
            TabIndex        =   96
            Top             =   2335
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A1 
            Bindings        =   "Fw_facturacion.frx":39C821
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   5775
            TabIndex        =   95
            Top             =   2325
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
            TabIndex        =   51
            Top             =   1580
            Width           =   6135
         End
         Begin VB.TextBox TxtMonto1 
            Alignment       =   2  'Center
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
            TabIndex        =   2
            Text            =   "0"
            Top             =   3615
            Width           =   1275
         End
         Begin VB.TextBox TxtDscto1 
            Alignment       =   2  'Center
            DataField       =   "cobranza_solicitado_dol"
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
            TabIndex        =   34
            Text            =   "0"
            Top             =   4155
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "Fw_facturacion.frx":39C83B
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1800
            TabIndex        =   52
            Top             =   3615
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
            Bindings        =   "Fw_facturacion.frx":39C855
            DataField       =   "cta_codigo"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   2520
            TabIndex        =   53
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
            Bindings        =   "Fw_facturacion.frx":39C86F
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1800
            TabIndex        =   54
            Top             =   4155
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
         Begin MSDataListLib.DataCombo DataCombo11 
            Bindings        =   "Fw_facturacion.frx":39C888
            DataField       =   "cta_codigo2"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   2520
            TabIndex        =   55
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
            Left            =   7620
            TabIndex        =   1
            Top             =   3075
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
            Format          =   92471297
            CurrentDate     =   41678
         End
         Begin MSDataListLib.DataCombo dtc_desc4A1 
            Bindings        =   "Fw_facturacion.frx":39C8A1
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos01"
            Height          =   315
            Left            =   1080
            TabIndex        =   94
            Top             =   2325
            Width           =   5040
            _ExtentX        =   8890
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Registro       Nro. de Recibo       Nro.de Factura         Nro. de Autorización"
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
            Left            =   120
            TabIndex        =   152
            Top             =   840
            Width           =   7005
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Elige el Documento a Emitir"
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
            TabIndex        =   142
            Top             =   3090
            Width           =   2460
         End
         Begin VB.Line Line11 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   7845
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00FFFF80&
            X1              =   7845
            X2              =   7845
            Y1              =   0
            Y2              =   2925
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
            Left            =   7725
            TabIndex        =   77
            Top             =   4155
            Width           =   1275
         End
         Begin VB.Label txtProg_Bs 
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
            Left            =   8340
            TabIndex        =   76
            Top             =   1260
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
            Left            =   8280
            TabIndex        =   75
            Top             =   975
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
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   120
            TabIndex        =   74
            Top             =   4185
            Width           =   7530
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00FFFF80&
            X1              =   0
            X2              =   10560
            Y1              =   2925
            Y2              =   2925
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
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   120
            TabIndex        =   73
            Top             =   3645
            Width           =   7560
         End
         Begin VB.Label Label36 
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
            TabIndex        =   72
            Top             =   3615
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
            Left            =   8340
            TabIndex        =   71
            Top             =   480
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
            Left            =   8340
            TabIndex        =   70
            Top             =   2115
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
            TabIndex        =   69
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
            Left            =   8280
            TabIndex        =   68
            Top             =   1845
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
            Top             =   240
            Width           =   6030
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
            ForeColor       =   &H00FFFFC0&
            Height          =   240
            Left            =   120
            TabIndex        =   64
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
            Left            =   8280
            TabIndex        =   63
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
            TabIndex        =   62
            Top             =   2240
            Width           =   960
         End
         Begin VB.Label lbl_fechaFac 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitud"
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
            Left            =   4560
            TabIndex        =   61
            Top             =   3090
            Width           =   2985
         End
         Begin VB.Label lbl_doc01 
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
            Left            =   1875
            TabIndex        =   60
            Top             =   1140
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
            TabIndex        =   59
            Top             =   660
            Visible         =   0   'False
            Width           =   1365
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
            TabIndex        =   58
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
            Left            =   3540
            TabIndex        =   57
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
            Left            =   5205
            TabIndex        =   56
            Top             =   1125
            Width           =   1995
         End
      End
      Begin VB.Frame FraNavega 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4750
         Left            =   40
         TabIndex        =   46
         Top             =   1440
         Width           =   5745
         Begin VB.OptionButton OptFilGral2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Enviados y NO Facturados"
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
            TabIndex        =   48
            Top             =   4395
            Width           =   2595
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
            Left            =   840
            TabIndex        =   47
            Top             =   4395
            Value           =   -1  'True
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dg_datos 
            Bindings        =   "Fw_facturacion.frx":39C8BB
            Height          =   4020
            Left            =   75
            TabIndex        =   49
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
      Begin VB.Frame FraNavega1 
         BackColor       =   &H00000000&
         Caption         =   "LISTA"
         ForeColor       =   &H00FFFFC0&
         Height          =   4750
         Left            =   -74960
         TabIndex        =   42
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
            Left            =   720
            TabIndex        =   45
            Top             =   4395
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptFilGral02 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Aprobados y No Facturados"
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
            TabIndex        =   44
            Top             =   4395
            Width           =   2595
         End
         Begin MSDataGridLib.DataGrid dg_datos1 
            Bindings        =   "Fw_facturacion.frx":39C8D3
            Height          =   4020
            Left            =   75
            TabIndex        =   43
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
      TabIndex        =   37
      Top             =   6345
      Width           =   14535
      Begin MSDataGridLib.DataGrid dg_datos16 
         Bindings        =   "Fw_facturacion.frx":39C8ED
         Height          =   1170
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   14280
         _ExtentX        =   25188
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
      TabIndex        =   36
      Top             =   7965
      Width           =   14535
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "Fw_facturacion.frx":39C907
         Height          =   1620
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   14295
         _ExtentX        =   25215
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
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   5235.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   705.26
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
   Begin VB.PictureBox Picture1 
      Height          =   1560
      Left            =   13680
      ScaleHeight     =   1500
      ScaleWidth      =   1695
      TabIndex        =   147
      Top             =   4080
      Visible         =   0   'False
      Width           =   1755
   End
   Begin Crystal.CrystalReport crRecibo 
      Left            =   2640
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
Attribute VB_Name = "Fw_facturacion"
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
'Private Declare Sub FastQRCode Lib "QRCodeLib.dll" _
'(ByVal Texto As String, _
' ByVal FileName As String)
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
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux20 As New ADODB.Recordset

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
Dim msgSalir, Caracter As String
'Dim queryinicial As String
Dim queryinicial1 As String
Dim queryinicial2 As String

'Dim descri_bien As String
'VARIABLES
Dim iResult As Variant  ', i%, y%
Dim marca1 As Variant

Dim VAR_CANT As Integer         'Cant_Alm,
Dim correlativo1 As Long
Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, correlv, NRO_COBR As Long
Dim VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SW, VAR_TSOL As Integer
Dim VAR_SOL, VAR_TIPOS As Integer
Dim i As Integer
Dim VAR_COMPM As Integer
Dim VAR_DOCR, VAR_DDIF As Integer

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, COBR_BS As Double
Dim VAR_CONTAB As Double
Dim VAR_PORC As Double
Dim VAR_13, VAR_87 As Double

Dim var_literal, VAR_PROY2, VAR_CTA, VAR_PROY3 As String
Dim VAR_CODTIPO, VAR_BENEF, VAR_GLOSA, VAR_MONEDA As String
Dim VAR_COD1, VAR_COD2, VAR_COD3 As String
Dim VAR_ANIO, VAR_MES, VAR_DIA, VAR_FECHA, VAR_FFAC As String
Dim VAR_COD4, VAR_TIPOV, VAR_CITE  As String
Dim DESAUX, VARAUX, VARCODIG As String
Dim VAR_EST, VAR_FAC, VAR_DOC As String
Dim VAR_ORG, VAR_FTE, VAR_PARTIDA As String
Dim VAR_ETAPA, VAR_TCOMP, EST_PROG As String
Dim gestion0, VAR_JQ, VAR_VTIPO As String

Dim codigo_doc As String
Dim Numero As String
Dim Autorizacion As String
Dim NroFactura As String
Dim NitCi As String
Dim Fecha As String
Dim Monto As String
Dim Llave As String
Dim CodigoContro As String
Dim VAR_NOMD, VAR_NOMH As String
Dim VAR_DCORR, VAR_HCORR As String

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
            BtnModificar.Visible = True
            If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
               'BtnImprimir3.Visible = True
               'BtnImprimir2.Visible = False
               'BtnImprimir3.Caption = "Facturar"
'               lbl_factura.Caption = "Nro.de Factura"
'               TxtCmpbte.Visible = True
'               TxtCmpbte.Locked = True
               lbl_docnro.Visible = False
               'TxtCmpbte.backColor = &H404040
               'TxtCmpbte.ForeColor = &HFFFFFF
               Lbl_nombre_fac.Caption = "Factura a Nombre de:                                                                                                      NIT/CI"
               lbl_fechas.Caption = "Fecha Facturación"
            Else
               'BtnImprimir3.Visible = False
               'BtnImprimir2.Visible = True
               'BtnImprimir3.Caption = "Recibo"
'               lbl_factura.Caption = "Nro.de Recibo"
               lbl_docnro.Visible = True
'               TxtCmpbte.Visible = False
               'TxtCmpbte.Locked = False     ' CAMBIAR DE Objeto
               'TxtCmpbte.backColor = &H80000005
               'TxtCmpbte.ForeColor = &H80000008
               Lbl_nombre_fac.Caption = "Recibo a Nombre de:                                                                                                       NIT/CI"
               lbl_fechas.Caption = "Fecha de Recibo"
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
            BtnModificar.Visible = False
'            BtnEliminar.Visible = False
'            BtnAprobar.Visible = False
'            BtnVer.Visible = True
'            FrmABMDet.Visible = False
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
            TxtDsctoTot.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg.backColor = &H404040       '&H80000013      'Fondo Oscuro
            BtnImprimir3.Visible = False
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
'        If glusuario = "ADMIN" Or glusuario = "RVALDIVIEZO" Or glusuario = "VPAREDES" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
'                BtnImprimir5.Visible = True
'        Else
'                BtnImprimir5.Visible = False
'        End If
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
    BtnModificar.Visible = False
'    BtnEliminar.Visible = False
'    BtnVer.Visible = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
'    FrmABMDet.Visible = False
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
            BtnModificar1.Visible = True
            BtnAprobar1.Visible = True
            If Ado_datos01.Recordset!doc_codigo_fac = "R-103" Then
                cmd_fac = "RECIBO"
'                lbl_factura1 = "Nro. de Recibo"
                lbl_fechaFac = "Fecha Solicitud Recibo"
            Else
                cmd_fac = "FACTURA"
'                lbl_factura1 = "Nro.de Factura"
                lbl_fechaFac = "Fecha Solicitud Factura"
            End If
        Else
            BtnModificar1.Visible = False
            BtnAprobar1.Visible = False
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
    BtnAprobar1.Visible = False
    BtnModificar1.Visible = False
    'BtnEliminar1.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
'    FrmABMDet.Visible = False
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
            If (glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS") Or glusuario = "VPAREDES" Or glusuario = "CNUÑEZ" Or glusuario = "FACTURACION" Then
                OptFilGral05.Visible = True
            Else
                OptFilGral05.Visible = False
            End If
            If Ado_datos02.Recordset!doc_codigo_fac = "R-103" Then
                lbl_factura3 = "Nro. de Recibo"
                Lbl_nombre_fac3.Caption = "Factura a Nombre de:                                                                                                 NIT/CI"
                lbl_fechas3.Caption = "Fecha de Recibo"
            Else
                lbl_factura3 = "Nro.de Factura"
            End If

        Else
            TxtDsctoTot2.backColor = &H404040        '&H80000013      'Fondo Oscuro
            DTPFechaProg2.backColor = &H404040       '&H80000013      'Fondo Oscuro
            If Ado_datos02.Recordset!estado_codigo = "APR" Then
                BtnAprobar.Visible = False
                BtnAprobar2.Visible = False
                BtnModificar2.Visible = False
                OptFilGral05.Visible = False
            Else
                'Or glusuario = "MVALDIVIA"
                If (glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "CNUÑEZ" Or glusuario = "VPAREDES") Or glusuario = "CNUÑEZ" Or glusuario = "FACTURACION" Then
                    BtnAprobar.Visible = True
                    BtnAprobar2.Visible = False
                    BtnModificar2.Visible = True
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
    BtnAprobar2.Visible = False
    BtnModificar2.Visible = False
    'BtnEliminar2.Visible = False

    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
'    FrmABMDet.Visible = False
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
        'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
        CryF02.ReportFileName = App.Path & "\reportes\ventas\fr_cobranzas_facturadas_unidad.rpt"
        CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'MODULO DE COBRANZAS' "
        CryF02.Formulas(2) = "subtitulo = 'ESTADO DE CUENTAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If
End Sub

Private Sub BtnAñadir_Click()
marca1 = Ado_datos.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
'        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
'        FrmABMDet.Visible = False
        FrmABMDet2.Visible = False
'        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = True
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
    Else
        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAprobar_Click()
 If Ado_datos02.Recordset.RecordCount > 0 Then
     If IsNull(Ado_datos02.Recordset("cobranza_observaciones")) Or (Ado_datos02.Recordset("cobranza_deuda_bs") = 0) Then
        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
     Else
        If Ado_datos02.Recordset("estado_codigo") = "REG" Then
           sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
           If sino = vbYes Then
               'If Ado_datos02.Recordset("venta_tipo") = "C" Or Ado_datos02.Recordset("venta_tipo") = "V" Then
               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               'End If
               gestion0 = glGestion                 'Ado_datos02.Recordset("ges_gestion")
               correlv = Ado_datos02.Recordset("venta_codigo")
               nroventa = Ado_datos02.Recordset("venta_codigo")
               
               VAR_BENEF = Ado_datos02.Recordset!beneficiario_codigo
               VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
               VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) '+ " - Nro.: " + Trim(VAR_CITE)
               VAR_DOL2 = Round(Ado_datos02.Recordset!cobranza_deuda_dol, 2)
               VAR_BS2 = Round(Ado_datos02.Recordset!cobranza_deuda_bs, 2)
               VAR_CTA = IIf(Ado_datos02.Recordset!cta_codigo = "", "NN", Ado_datos02.Recordset!cta_codigo)
               VAR_PROY2 = Ado_datos16.Recordset!edif_codigo
               VAR_COD4 = Ado_datos16.Recordset!unidad_codigo
               VAR_TIPOV = Ado_datos16.Recordset!venta_tipo
               VAR_SOL = Ado_datos16.Recordset!solicitud_codigo
               If Ado_datos02.Recordset!cta_codigo <> "NN" Then
                    VAR_FFAC = Ado_datos02.Recordset!cobranza_fecha_cobro1
               Else
                    VAR_FFAC = Ado_datos02.Recordset!cobranza_fecha_cobro
               End If
               NRO_COBR = Me.Ado_datos02.Recordset!cobranza_codigo
               var_literal = Ado_datos02.Recordset!Literal
               VAR_MONEDA = Ado_datos02.Recordset!tipo_moneda
               VAR_CODTIPO = "REC"
               VAR_DOC = "R-110"
               VAR_ETAPA = "FIN-02-03"
               VAR_TCOMP = "RECAUDADO (INGRESOS)"
               VAR_ANIO = Year(VAR_FFAC)
               VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
               
'               'Correlativo por Mes y Tipo de Comprobante
'                Set rs_aux2 = New ADODB.Recordset
'                SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(VAR_MES) & "' and cta_codigo2 = '" & VAR_DOC & "' ) "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                      rs_aux2!numero_correlativo = rs_aux2!numero_correlativo + 1
'                      VAR_DOCR = rs_aux2!numero_correlativo
'                      rs_aux2.Update
'                End If
'                'Correlativo General por Documento I-E-T
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos02.Recordset!doc_codigo & "'  "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                    'Ado_datos02.Recordset!doc_numero = rs_aux2!correl_doc
'                    'Txt_campo1.Caption = rs_aux2!correl_doc
'                    rs_aux2.Update
'                End If
                ' GRABA Nombre de Archivo en ao_ventas_cabecera
                
                'Llave = Trim(rs_aux1!dosifica_llave)
                'NitCi = Ado_datos.Recordset!beneficiario_codigo_fac     'VAR_BENEF
                'Autorizacion = rs_aux1!dosifica_autorizacion
               
               ' APRUEBA ao_ventas_cabecera
               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And cobranza_codigo = " & Ado_datos02.Recordset("cobranza_codigo") & " "
                
                'VAR_ARCH = RTrim(RTrim(Ado_datos02.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos02.Recordset!doc_numero))
                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
                'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos02.Recordset("venta_codigo") & " "
           
               
               'marca1 = Ado_datos02.Recordset.Bookmark
               'Ado_datos02.Recordset.Requery
        '       Ado_datos02.Refresh
               'Ado_datos02.Recordset.Move marca1 - 1
               
               '  Set rstacumdet = New ADODB.Recordset
                '  If rstacumdet.State = 1 Then rstacumdet.Close
                '  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranzas where ges_gestion = '" & Ado_datos02.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos02.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
                '
                '  Set rstdestino = New ADODB.Recordset
                '  If rstdestino.State = 1 Then rstdestino.Close
                '  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & gestion0 & "' and venta_codigo = " & nroventa, db, adOpenKeyset, adLockOptimistic
                '  If rstdestino.RecordCount > 0 Then
                '    rstdestino!deuda_cobrada = rstacumdet!Cobrobs
                '    rstdestino!saldo_p_cobrar = (rstdestino!monto_total_Bs - rstdestino!monto_cobrado - rstdestino!deuda_cobrada)
                '    rstdestino.Update
                '  End If
                '  If rstdestino.State = 1 Then rstdestino.Close
                '  If rstacumdet.State = 1 Then rstacumdet.Close
               VAR_SW = 2
               Call Contabiliza_venta
               db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
               db.Execute "UPDATE co_diario SET co_diario.estado_codigo = co_comprobante_m.estado_codigo FROM co_diario INNER JOIN co_comprobante_m ON co_diario.Cod_Comp =co_comprobante_m.Cod_Comp where co_diario.estado_codigo Is Null "
               Call OptFilGral1_Click
           End If
        End If
     End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAprobar1_Click()
 If Ado_datos01.Recordset.RecordCount > 0 Then
     If IsNull(Ado_datos01.Recordset("cobranza_solicitado_bs")) Or (Ado_datos01.Recordset("cobranza_solicitado_bs") = 0) Then
        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
     Else
        If Ado_datos01.Recordset("estado_codigo_sol") = "REG" Then
           Select Case cmd_fac.Text
              Case "RECIBO"
                  sino = MsgBox("Esta seguro de Enviar a Registro de Cobranza mediante RECIBO ?", vbYesNo, "Confirmando")
                  VAR_EST = "REG"
                  VAR_FAC = "R-103"
              Case "FACTURA"
                  sino = MsgBox("Esta seguro de Enviar a Facturación ?", vbYesNo, "Confirmando")
                  VAR_EST = "REG"
                  VAR_FAC = "R-101"
              Case Else
                  sino = MsgBox("Esta seguro de Enviar a Facturación ?", vbYesNo, "Confirmando")
                  VAR_EST = "REG"
                  VAR_FAC = "R-101"
           End Select
           'sino = MsgBox("Esta seguro de Enviar a Facturación ?", vbYesNo, "Confirmando")
           If sino = vbYes Then
                Ado_datos01.Recordset!cobranza_solicitado_bs = Ado_datos01.Recordset!cobranza_programada_bs               'Monto Solicitado Bs
                Ado_datos01.Recordset!cobranza_solicitado_dol = Ado_datos01.Recordset!cobranza_programada_dol           'Monto Solicitado Dol
                Ado_datos01.Recordset!cobranza_fecha_sol = Date
                Ado_datos01.Recordset!estado_codigo_sol = "APR"
                Ado_datos01.Recordset!estado_codigo_fac = VAR_EST   'REG
                Ado_datos01.Recordset!doc_codigo_fac = VAR_FAC      'R-101
                'If Ado_datos01.Recordset!unidad_codigo <> "DVTA" Then
                '    Ado_datos01.Recordset!doc_codigo_fac = "R-101"
                'Else
                '    'COMPLETAR CON TIPO DE RECIBO (DE DONDE VIENE)
                '    If cmd_fac.Text = "RECIBO" Then
                '        Ado_datos01.Recordset!doc_codigo_fac = "R-103"
                '    Else
                '        Ado_datos01.Recordset!doc_codigo_fac = "R-101"
                '    End If
                'End If
                Ado_datos01.Recordset.Update
              'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
           End If
        End If
     End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
 End If
End Sub

Private Sub BtnAprobar2_Click()

    db.Execute "update ao_ventas_cobranza set estado_codigo1 = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "

' If Ado_datos02.Recordset.RecordCount > 0 Then
'   If Ado_datos02.Recordset!cmpbte_deposito <> "0" And Ado_datos02.Recordset!cmpbte_deposito <> "" Then
'      If Ado_datos02.Recordset!Cta_Codigo <> "NN" And Ado_datos02.Recordset!Cta_Codigo <> "" Then
'        COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'        If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'           MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'           Exit Sub
'        Else
'           If COBR_BS < Ado_datos02.Recordset!cobranza_total_bs Then
'               'MsgBox "No se puede APROBAR, hasta que el Monto Cobrado sea igual al Monto Facturado. Vuelva a intentar ...", , "Atención"
'               MsgBox "No se puede APROBAR hasta que el Total Monto Cobrado sea igual al Monto Facturado ...", , "Atención"
'               Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'               Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'               Ado_datos02.Recordset!estado_codigo = "REG"
'               Ado_datos02.Recordset.Update
'               'Exit Sub
'           Else
'               If Ado_datos02.Recordset("estado_codigo_bco") = "REG" Then
'                  sino = MsgBox("Esta seguro de Verificar la Cobranza ?", vbYesNo, "Confirmando")
'                  If sino = vbYes Then
'                    If TxtDscto2.Text = "0.00" Or TxtDscto2.Text = "" Then
'                       Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro2.Value
'                    Else
'                       Ado_datos02.Recordset!cobranza_fecha_cobro = DTPFechaCobro02.Value
'                    End If
'                    Ado_datos02.Recordset!cobranza_fecha_cobro1 = DTPFechaCobro2.Value
'                    Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                    Ado_datos02.Recordset!estado_codigo_bco = "APR"
'                    Ado_datos02.Recordset!estado_codigo = "REG"
'                    Ado_datos02.Recordset.Update
'                     'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'                  End If
'               Else
'                   MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'               End If
'           End If
'        End If
'      Else
'        MsgBox "No se puede APROBAR, debe elegir una Cuenta Bancaria !! ", vbExclamation, "Atención!"
'      End If
'   Else
'    MsgBox "No se puede APROBAR, debe registrar el Comprobante (Cpbte) de Depósito !! ", vbExclamation, "Atención!"
'   End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
' End If
End Sub

Private Sub BtnAprobar3_Click()
' If Ado_datos02.Recordset.RecordCount > 0 Then
'     COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs '+ Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
'     If IsNull(Ado_datos02.Recordset!cobranza_deuda_bs) Or (COBR_BS = 0) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'        Exit Sub
'     Else
'        If COBR_BS <= Ado_datos02.Recordset!cobranza_total_bs Then
'            If Ado_datos02.Recordset("estado_codigo_bco1") = "REG" Then
'               sino = MsgBox("Esta seguro de Verificar la Cobranza 1 ?", vbYesNo, "Confirmando")
'               If sino = vbYes Then
'                  db.Execute "UPDATE ao_ventas_cobranza SET  "
'                    Ado_datos02.Recordset!cobranza_fecha_cobro = Date
'                    Ado_datos02.Recordset!estado_codigo_bco1 = "APR"
'                    Ado_datos02.Recordset!estado_codigo = "REG"
'                    Ado_datos02.Recordset.Update
'                  'db.Execute "update ao_ventas_cobranza set estado_codigo_sol = 'APR' Where cobranza_codigo = " & Ado_datos01.Recordset("cobranza_codigo") & " "
'               End If
'            Else
'                MsgBox "No se puede APROBAR, el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
'            End If
'
'        Else
'            MsgBox "No se puede APROBAR, un Monto Cobrado Mayor al Monto Facturado. Vuelva a intentar ...", , "Atención"
'            Exit Sub
'        End If
'     End If
' Else
'    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
' End If
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

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
'  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
  If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  FraNavega.Enabled = True
  FrmCobros.Enabled = False
  'Fra_datos.Enabled = True
  FrmDetalle.Enabled = True
  FrmCobranza.Visible = True
  'Fra_Total.Visible = True
  dg_datos.Visible = True
'  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True

  SSTab1.Tab = 1
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  SSTab1.TabEnabled(2) = False
  'Ado_datos.Recordset.Move marca1 - 1
'  BtnImprimir2.Visible = True
  BtnImprimir3.Visible = True
  
  swnuevo = 0
   
End Sub

Private Sub BtnCancelar1_Click()

  FraGrabarCancelar1.Visible = False
  marca1 = Ado_datos01.Recordset.Bookmark
  If Ado_datos01.Recordset("estado_codigo_sol") = "REG" Then
    Call OptFilGral01_Click
  Else
    Call OptFilGral02_Click
  End If
  FraNavega1.Enabled = True
  FrmCobros1.Enabled = False
  FrmDetalle.Enabled = True
  FrmCobranza.Enabled = True
  
'  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    End If
  'Ado_datos01.Recordset.Move marca1 - 1
  swnuevo = 0

End Sub

Private Sub BtnCancelar2_Click()
  FraGrabarCancelar2.Visible = False
  marca1 = Ado_datos02.Recordset.Bookmark
  If (Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG") Then
    Call OptFilGral03_Click
  Else
    Call OptFilGral04_Click
  End If
  FraNavega2.Enabled = True
  FrmCobros2.Enabled = False
  FrmDetalle.Enabled = True
  FrmCobranza.Enabled = True
  
'  FrmABMDet.Visible = True
  FrmABMDet2.Visible = True
'  BtnAprobar3.Visible = True
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    End If
  'Ado_datos01.Recordset.Move marca1 - 1
  DTPFechaCobro2A.Visible = False
  swnuevo = 0

End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
    FraGrabarCancelar.Enabled = True
End Sub

Private Sub BtnEliminar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "APR" And Ado_datos.Recordset!estado_codigo_bco = "REG" Then      'Ado_datos.Recordset("estado_codigo_anl") = "REG"
      sino = MsgBox("Esta seguro de ANULAR la facturación registrada ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
        sino = MsgBox("Volverá a emitir otra FACTURA con este mismo registro ? (Si elige NO, se cierra el registro)", vbYesNo, "Confirmando")
        If sino = vbYes Then
          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'REG' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set factura_impresa = 'N' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
        Else
          db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'ANL' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
        End If
          db.Execute "update ao_ventas_cobranza set cobranza_nro_factura_anl = '" & Ado_datos.Recordset!cobranza_nro_factura & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_fecha_anl = '" & Format(Date, "dd/mm/yyyy") & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set usr_codigo_anl = '" & glusuario & "' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set estado_codigo_anl = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_fecha_ant = cobranza_fecha_fac Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_codigo_control_anl = cobranza_codigo_control Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set correl_contab_anl = correl_contab Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion_anl = cobranza_nro_autorizacion Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  and cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & "  "
          
          Set rs_datos12 = New ADODB.Recordset
          If rs_datos12.State = 1 Then rs_datos12.Close
          rs_datos12.Open "Select * from ao_ventas_cobro_anl where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " and cobranza_nro_factura_anl = " & Ado_datos.Recordset!cobranza_nro_factura & " ", db, adOpenKeyset, adLockOptimistic
          If rs_datos12.RecordCount > 0 Then
            MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
          Else
            'wwwwwwwwwwwwwwwwwwwww
              ' hora_registro
            rs_datos12.AddNew
            rs_datos12!ges_gestion = glGestion
            rs_datos12!cobranza_codigo = Ado_datos.Recordset!cobranza_codigo
            rs_datos12!venta_codigo = Ado_datos.Recordset!venta_codigo
            
            rs_datos12!cobranza_nro_factura_anl = Ado_datos.Recordset!cobranza_nro_factura
            rs_datos12!cobranza_prog_codigo = Ado_datos.Recordset!cobranza_prog_codigo
            rs_datos12!beneficiario_codigo_fac = Ado_datos.Recordset!beneficiario_codigo_fac
            rs_datos12!cobranza_anuladal_bs = Ado_datos.Recordset!cobranza_total_bs
            rs_datos12!cobranza_anulada_dol = Ado_datos.Recordset!cobranza_total_dol
            
            rs_datos12!cobranza_fecha_anl = Ado_datos.Recordset!cobranza_fecha_fac      'Format(Date, "dd/mm/yyyy")
            rs_datos12!cobranza_fecha_fac2 = Ado_datos.Recordset!cobranza_fecha_fac2
            rs_datos12!cobranza_observaciones = Ado_datos.Recordset!cobranza_observaciones
            rs_datos12!cobranza_codigo_control_anl = Ado_datos.Recordset!cobranza_codigo_control
            rs_datos12!Literal = Ado_datos.Recordset!Literal
        
            rs_datos12!cobranza_nro_autorizacion_anl = Ado_datos.Recordset!cobranza_nro_autorizacion
            rs_datos12!correl_contab_anl = Ado_datos.Recordset!correl_contab
            rs_datos12!estado_codigo_anl = "APR"            'Ado_datos.Recordset!estado_codigo_anl
            rs_datos12!usr_codigo_anl = glusuario           'Ado_datos.Recordset!usr_codigo_anl
            rs_datos12!Fecha_Registro = Ado_datos.Recordset!Fecha_Registro
        
            rs_datos12!trans_codigo = Ado_datos.Recordset!trans_codigo
            rs_datos12!cmpbte_deposito = Ado_datos.Recordset!cmpbte_deposito
            rs_datos12!cta_codigo = Ado_datos.Recordset!cta_codigo
            rs_datos12.Update
          End If
      End If
        '  rs_datos12!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
          'wwwwwwwwwwwwwwwwwwwww
          'marca1 = Ado_datos.Recordset.Bookmark
          'Call OptFilGral2_Click
          'Ado_datos.Recordset.Move marca1 - 1
    Else
      MsgBox "NO se puede ANULAR, porque el registro NO fue Facturado o ya fue Cobrado...", , "Atencion"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub cambiarEtiquetaFactura()
'    If lbl_fac.Caption <> "R-101" Then
'       TxtCmpbte = False
'       TxtCmpbte.backColor = &H80000005
'       TxtCmpbte.ForeColor = &H80000008
''       lbl_factura.Caption = "Nro.de Recibo"
'    Else
'       TxtCmpbte = True
'       TxtCmpbte.backColor = &H404040
'       TxtCmpbte.ForeColor = &HFFFFFF
''       lbl_factura.Caption = "Nro.de Factura"
'    End If
End Sub

Private Sub BtnGrabar_Click()
'  Call cambiarEtiquetaFactura
  If dtc_codigo4A.Text = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo5.Text = "" Then
    MsgBox "Debe Elejir <<Factura a Nombre de:>> !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtMonto = "" Or TxtMonto = "0" Or TxtMonto = "0.00" Then
    MsgBox "Debe Registrar el " + lbl_monto.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If TxtObs = "" Then
    MsgBox "Debe Registrar el " + lbl_obs.Caption + " de la Cobranza, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    If swnuevo = 1 Then
'      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
'      Set Ado_datos16.Recordset = rstdestino
'      Ado_datos16.Recordset.AddNew
      Ado_datos.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
      Ado_datos.Recordset!ges_gestion = glGestion       'Ado_datos.Recordset("ges_gestion")
      'Ado_datos.Recordset!cobranza_fecha_prog = DTPFechaProg                                'Fecha Programada a Cobrar
    End If
'      If Ado_datos.Recordset!beneficiario_codigo = "0" Then
'        Ado_datos.Recordset!beneficiario_codigo = dtc_codigo5.Text        'lbl_nit.Caption                                  'Codigo Beneficiario (Cliente)
'      End If
      Ado_datos.Recordset!beneficiario_codigo_fac = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)       ' dtc_codigo5.Text  'dtc_codigo2A.Text                            'Beneficiario (Factura a nombre de ...)
      Ado_datos.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      Ado_datos.Recordset!trans_codigo = IIf(dtc_codigo6.Text = "", "O", dtc_codigo6.Text) 'tipo de Transaccion
      'Ado_datos.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito.Text = "", "0", Txt_deposito.Text)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      Ado_datos.Recordset!cta_codigo = IIf(dtc_cta.Text = "", "NN", dtc_cta.Text)
      Ado_datos.Recordset!cta_codigo2 = IIf(dtc_codigo7.Text = "", "NN", dtc_codigo7.Text)
      If TxtMonto.Text = "" Then
        Ado_datos.Recordset!cobranza_deuda_bs = "0"                                  'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos.Recordset!cobranza_tdc = IIf(IsNull(Txt_tdc = ""), 6.96, CDbl(Txt_tdc.Text))                               'Monto Cobrado Bs.
'        Ado_datos.Recordset!cobranza_deuda_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
        Ado_datos.Recordset!cobranza_total_bs = Round(CDbl(TxtMonto.Text), 2)                                 'Monto Cobrado Bs.
        Ado_datos.Recordset!cobranza_total_dol = Round(CDbl(TxtMontoDol), 2)       'CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      Ado_datos.Recordset!cobranza_descuento_bs = Round(CDbl(TxtMonto.Text) * 0.13, 2)                                                              'Monto Cobrado Bs. * 13%
      Ado_datos.Recordset!cobranza_descuento_dol = Round(Ado_datos.Recordset!cobranza_total_bs - Ado_datos.Recordset!cobranza_descuento_bs, 2)      'Monto Cobrado Bs. * 87%
      'VAR_GLOSA = Trim(Ado_datos02.Recordset!cobranza_observaciones) + " - Nro.: " + Trim(VAR_CITE)
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos.Recordset!cobranza_total_bs <> 0 Then
            Ado_datos.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!cobranza_total_bs)) + " BOLIVIANOS"
      End If
      'Ado_datos.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
      'Call acumulaMont(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!correl_venta, Ado_datos.Recordset!venta_codigo)
      'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
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
'        'GENERA CORREL NOTA DEBITO POR DEPTO INI
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        'rs_aux5.Open "Select correl_contab as Codigo from gc_departamento where depto_codigo = '" & Left(VAR_PROY3, 1) & "'    ", db, adOpenStatic
'        rs_aux5.Open "Select * from fc_correl where tipo_tramite  = 'NDEBITO '    ", db, adOpenStatic
'        If Not rs_aux5.EOF Then
'            VAR_CONTAB = IIf(IsNull(rs_aux5!numero_correlativo), 1, CDbl(rs_aux5!numero_correlativo) + 1)
'        End If
'        'rs_aux5!Codigo = VAR_CONTAB
'        'rs_aux5.Update
'        db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
'        db.Execute "update fc_correl set numero_correlativo = " & VAR_CONTAB & " Where tipo_tramite = 'NDEBITO' "
'        'Ado_datos.Recordset!correl_contab = VAR_CONTAB
'        'GENERA CORREL NOTA DEBITO POR DEPTO FIN
'        If VAR_CONTAB < 10 Then
'            Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-000" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 9 And VAR_CONTAB < 100 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-00" + Str(VAR_CONTAB) + ")"
'        End If
'        If VAR_CONTAB > 99 And VAR_CONTAB < 1000 Then
'           Ado_datos.Recordset!cobranza_observaciones = TxtObs.Text + " (ND-0" + Str(VAR_CONTAB) + ")"
'        End If
      Ado_datos.Recordset!proceso_codigo = "FIN"
      Ado_datos.Recordset!subproceso_codigo = "FIN-02"
      Ado_datos.Recordset!etapa_codigo = "FIN-02-02"
      'Ado_datos.Recordset!clasif_codigo = "ADM"
      'Ado_datos.Recordset!doc_codigo = IIf(lbl_doc1 = "", "R-105", lbl_doc1)
      'Ado_datos.Recordset!doc_numero = IIf(lbl_docnro = "", "0", lbl_docnro)
      Ado_datos.Recordset!cmpbte_deposito = IIf(Txt_deposito = "", "0", Txt_deposito)
      If lbl_fac <> "R-101" Then
        Ado_datos.Recordset!doc_codigo_fac = "R-103"
      Else
        Ado_datos.Recordset!doc_codigo_fac = "R-101"
      End If
      If Ado_datos.Recordset!factura_impresa = "N" Then
         TxtCmpbte = "0"
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      Else
         Ado_datos.Recordset!cobranza_nro_factura = IIf(TxtCmpbte = "", "0", Trim(TxtCmpbte))
      End If
      Ado_datos.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion = "", "0", Trim(TxtAutorizacion))
      Ado_datos.Recordset!poa_codigo = "3.1.2"
      Ado_datos.Recordset!cobranza_fecha_fac = Date 'DTPFechaCobro.Value         'Fecha de Facturacion
        'VAR_ANIO = CStr(glGestion)
        'VAR_MES = CStr(Month(Date))
        'VAR_DIA = CStr(Day(Date))
      'Ado_datos.Recordset!estado_codigo1 = "APR"
      Ado_datos.Recordset!cobranza_fecha_fac2 = ""        'VAR_ANIO & VAR_MES & VAR_DIA          'Fecha de Facturacion Texto
      Ado_datos.Recordset!estado_codigo_fac = "REG"
      Ado_datos.Recordset!usr_codigo = glusuario
      Ado_datos.Recordset!Fecha_Registro = Format(Date, "dd/mm/yyyy")
      Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos.Recordset.Update
    db.CommitTrans
    MsgBox "El registro se guardo correctamente"
    'Ado_datos.Recordset!doc_numero = Ado_datos.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos.Refresh
    'Ado_datos.Recordset.MoveLast
  End If
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
'    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FrmDetalle.Enabled = True
    FrmCobranza.Visible = True
    FrmCobros.Enabled = False
'    TxtCobrador.Visible = True
'    FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
'    BtnImprimir2.Visible = True
    BtnImprimir3.Visible = True
    
    swnuevo = 0
    
  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If

End Sub

Private Sub BtnGrabar1_Click()
  If TxtMonto1 = "" Or TxtMonto1 = "0" Or TxtMonto1 = "0.00" Then
    MsgBox "Debe Registrar el Monto por Cobrar BOB , !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
    db.BeginTrans
      Ado_datos01.Recordset!cmpbte_deposito = "0"       'IIf(Txt_deposito1.Text = "", "0", Txt_deposito1.Text)
      Ado_datos01.Recordset!cta_codigo = IIf(dtc_cta.Text = "", "NN", dtc_cta.Text)
      Ado_datos01.Recordset!cta_codigo2 = IIf(dtc_codigo7.Text = "", "NN", dtc_codigo7.Text)
'      If TxtMonto1.Text = "" Then
        Ado_datos01.Recordset!cobranza_deuda_bs = "0"         'Monto Cobrado Bs.
        Ado_datos01.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
'      Else
'        Ado_datos01.Recordset!cobranza_deuda_bs = CDbl(TxtMonto1.Text)                                  'Monto Cobrado Bs.
'        Ado_datos01.Recordset!cobranza_deuda_dol = CDbl(TxtMonto1.Text) / GlTipoCambioMercado        'Monto en Dolares
'      End If
'      If TxtDscto1.Text = "" Then
        Ado_datos01.Recordset!cobranza_deuda_dol2 = "0"                              'Monto Cobrado Dolares
        Ado_datos01.Recordset!cobranza_deuda_bs2 = "0"         'Monto en Bs
'      Else
'        Ado_datos01.Recordset!cobranza_deuda_dol2 = IIf(TxtDscto1.Text = "", 0, CDbl(TxtDscto1.Text))                              'Monto Cobrado Dolares
'        Ado_datos01.Recordset!cobranza_deuda_bs2 = CDbl(TxtDscto1.Text) * GlTipoCambioMercado         'Monto en Bs
'      End If
      Ado_datos01.Recordset!cobranza_solicitado_bs = IIf(TxtMonto1.Text = "0.00", Ado_datos01.Recordset!cobranza_programada_bs, CDbl(TxtMonto1.Text))                                                 'Monto Cobrado Bs
      Ado_datos01.Recordset!cobranza_solicitado_dol = IIf(TxtDscto1.Text = "0.00", Ado_datos01.Recordset!cobranza_programada_dol, CDbl(TxtDscto1.Text))                                                 'Monto Corado Dol
'      Ado_datos01.Recordset!cobranza_total_bs = Ado_datos01.Recordset!cobranza_deuda_bs + Ado_datos01.Recordset!cobranza_deuda_bs2                  'Monto Total Bs
'      Ado_datos01.Recordset!cobranza_total_dol = Ado_datos01.Recordset!cobranza_deuda_dol + Ado_datos01.Recordset!cobranza_deuda_dol2               'Monto Total Dol
'      Ado_datos01.Recordset!cobranza_descuento_bs = Ado_datos01.Recordset!cobranza_total_bs * 0.013                                                 'Monto Total Bs * 13%
'      Ado_datos01.Recordset!cobranza_descuento_dol = Ado_datos01.Recordset!cobranza_total_bs - Ado_datos01.Recordset!cobranza_descuento_bs          'Monto Total Dol * 87%
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos01.Recordset!cobranza_solicitado_bs <> 0 Then
            Ado_datos01.Recordset!Literal = Literal(CStr(Ado_datos01.Recordset!cobranza_solicitado_bs)) + " BOLIVIANOS"
      End If
      'Call acumulaMont(Ado_datos01.Recordset("ges_gestion"), Ado_datos01.Recordset("venta_codigo"))

      Ado_datos01.Recordset!proceso_codigo = "FIN"
      Ado_datos01.Recordset!subproceso_codigo = "FIN-02"
      Ado_datos01.Recordset!etapa_codigo = "FIN-02-03"
      Ado_datos01.Recordset!clasif_codigo = "ADM"
      Ado_datos01.Recordset!doc_codigo = IIf(lbl_doc01 = "", "R-105", lbl_doc01)
      Ado_datos01.Recordset!doc_numero = "0"          'IIf(lbl_docnro1 = "", "0", lbl_docnro1)
      If cmd_fac = "RECIBO" Then
        Ado_datos01.Recordset!doc_codigo_fac = "R-103"
      Else
        Ado_datos01.Recordset!doc_codigo_fac = "R-101"
      End If
      'Ado_datos01.Recordset!cobranza_nro_factura = IIf(TxtCmpbte1 = "", "0", Trim(TxtCmpbt1))
      'Ado_datos01.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion1 = "", "0", Trim(TxtAutorizacion1))
      Ado_datos01.Recordset!poa_codigo = "3.1.2"
      Ado_datos01.Recordset!cobranza_fecha_sol = DTPfechasol.Value         'Fecha de Cobranza
      Ado_datos01.Recordset!estado_codigo_sol = "REG"
      Ado_datos01.Recordset!usr_codigo_sol = glusuario
      Ado_datos01.Recordset!Fecha_Registro = Format(Date, "dd/mm/yyyy")
      Ado_datos01.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos01.Recordset.Update
    db.CommitTrans
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    End If
    FraNavega1.Enabled = True
'      fraOpciones1.Visible = True
      FraGrabarCancelar1.Visible = False
      FrmDetalle.Enabled = True
      FrmCobranza.Enabled = True
'      FrmABMDet.Visible = True
      FrmABMDet2.Visible = True
      FrmCobros1.Enabled = False
    swnuevo = 0
  End Sub

Private Sub BtnGrabar2_Click()
    If (TxtMonto02 = "" Or TxtMonto02 = "0" Or TxtMonto02 = "0.00") Then
      'MsgBox "Debe Registrar el " + lbl_monto1.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
        MsgBox "Debe Registrar el Monto Cobrado Bs. o Cobrado USD, !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
    End If
    If (DataCombo9.Text = "") Then
        MsgBox "Debe Registrar el Tipo de Transacción, !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
    End If
    'RONALD
'    If (CDate(Ado_datos02.Recordset!cobranza_fecha_fac) > CDate(DTPFechaCobro2.Value)) Then
'        MsgBox "La <<Fecha Cobranza1>> No puede ser MENOR a la <<Fecha de Facturación = " + CStr(Ado_datos02.Recordset!cobranza_fecha_fac) + ">>, Vuelva a Intentar !! ", vbExclamation, "Atención!"
'        Exit Sub
'    End If
    db.BeginTrans
      Ado_datos02.Recordset!cmpbte_deposito = IIf(Txt_deposito2.Text = "", "0", Txt_deposito2.Text)
      Ado_datos02.Recordset!cmpbte_deposito2 = IIf(Txt_deposito3.Text = "", "0", Txt_deposito3.Text)
      Ado_datos02.Recordset!cta_codigo = IIf(dtc_cta2.Text = "", "NN", dtc_cta2.Text)
      Ado_datos02.Recordset!cta_codigo2 = IIf(dtc_codigo7_2.Text = "", "NN", dtc_codigo7_2.Text)
      If dtc_cta2.Text = "NN" Or dtc_cta2.Text = "" Then
        Ado_datos02.Recordset!cobranza_deuda_bs = "0"         'Monto Cobrado Bs.
        Ado_datos02.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      End If
      If TxtMonto02.Text = "" Then
        Ado_datos02.Recordset!cobranza_deuda_bs = "0"         'Monto Cobrado Bs.
        Ado_datos02.Recordset!cobranza_deuda_dol = "0"        'Monto en Dolares
      Else
        Ado_datos02.Recordset!cobranza_deuda_bs = CDbl(TxtMonto02.Text)                               'Monto Cobrado Bs.
        Ado_datos02.Recordset!cobranza_deuda_dol = CDbl(TxtMonto02D.Text)        'CDbl(TxtMonto02.Text) / GlTipoCambioMercado        'Monto en Dolares
      End If
      If TxtDscto2.Text = "" Or TxtDscto2.Text = "0" Or TxtDscto2.Text = "0.00" Then
        Ado_datos02.Recordset!cobranza_deuda_dol2 = "0"                             'Segundo Monto Cobrado Dolares
        Ado_datos02.Recordset!cobranza_deuda_bs2 = "0"                              'Segundo Monto en Bs
      Else
        Ado_datos02.Recordset!cobranza_deuda_bs2 = IIf(TxtDscto2.Text = "", 0, CDbl(TxtDscto2.Text))   'Segundo Monto en Bs                           'Monto Cobrado Dolares
        Ado_datos02.Recordset!cobranza_deuda_dol2 = CDbl(TxtDscto2.Text) / GlTipoCambioMercado         'Segundo Monto en Dolares
      End If
        'Ado_datos02.Recordset!cobranza_descuento_bs = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2              'Monto Total Bs
        'Ado_datos02.Recordset!cobranza_descuento_dol = Ado_datos02.Recordset!cobranza_deuda_dol + Ado_datos02.Recordset!cobranza_deuda_dol2               'Monto Total Dol
        
      Ado_datos02.Recordset!cobranza_descuento_bs = Ado_datos02.Recordset!cobranza_total_bs * 0.13                  'Monto Total Bs * 13%
      Ado_datos02.Recordset!cobranza_descuento_dol = Ado_datos02.Recordset!cobranza_total_dol - Ado_datos02.Recordset!cobranza_descuento_bs               'Monto Total Dol * 87%
      
      Ado_datos02.Recordset!tipo_moneda = cmd_moneda1                                 'Tipo Moneda1
      Ado_datos02.Recordset!tipo_moneda2 = cmd_moneda2                                'Tipo Moneda2
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW RONALD
'      If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
'        Ado_datos02.Recordset!cobranza_fecha_cobro1 = Format(DTPFechaCobro2.Value, "dd/mm/yyyy")         'Fecha de Cobranza1
'        Ado_datos02.Recordset!cobranza_fecha_cobro = Format(DTPFechaCobro2.Value, "dd/mm/yyyy")         'Fecha de Cobranza2
'      Else
      'If (CDate(DTPFechaCobro2.Value) = "01/01/1900") Then
        Ado_datos02.Recordset!cobranza_fecha_cobro1 = IIf(IsNull(DTPFechaCobro2.Value) Or (CDate(DTPFechaCobro2.Value) = "01/01/1900"), Date, Format(DTPFechaCobro2.Value, "dd/mm/yyyy"))         'Fecha de Cobranza1
        Ado_datos02.Recordset!cobranza_fecha_cobro = IIf(IsNull(DTPFechaCobro02.Value) Or (CDate(DTPFechaCobro02.Value) = "01/01/1900"), Date, Format(DTPFechaCobro02.Value, "dd/mm/yyyy"))        'Fecha de Cobranza2
      'Else
      '  Ado_datos02.Recordset!cobranza_fecha_cobro1 = IIf(IsNull(DTPFechaCobro2.Value), Date, Format(DTPFechaCobro2.Value, "dd/mm/yyyy"))        'Fecha de Cobranza1
      '  Ado_datos02.Recordset!cobranza_fecha_cobro = IIf(IsNull(DTPFechaCobro02.Value), Date, Format(DTPFechaCobro02.Value, "dd/mm/yyyy"))        'Fecha de Cobranza2
      'End If
'      End If
      
      COBR_BS = Ado_datos02.Recordset!cobranza_deuda_bs + Ado_datos02.Recordset!cobranza_deuda_bs2            'Monto Total Cobrado Bs
      If COBR_BS > 0 Then
            Ado_datos02.Recordset!Literal = Literal(CStr(COBR_BS)) + " BOLIVIANOS"
      Else
            Ado_datos02.Recordset!Literal = "CERO 00/100 BOLIVIANOS"
      End If
'      Call acumulaMont(Ado_datos02.Recordset("ges_gestion"), Ado_datos02.Recordset("venta_codigo"))
      
      'Ado_datos02.Recordset!proceso_codigo = "FIN"
      Ado_datos02.Recordset!subproceso_codigo = "FIN-02"
      Ado_datos02.Recordset!etapa_codigo2 = "FIN-02-03"
      'Ado_datos02.Recordset!clasif_codigo = "ADM"
      Ado_datos02.Recordset!doc_codigo_cobr = "RE-426"        'IIf(lbl_doc01 = "", "R-105", lbl_doc01)
      Ado_datos02.Recordset!doc_numero = IIf(Txt_docnro = "", "0", Txt_docnro)
      'Ado_datos02.Recordset!doc_codigo_fac = "R-101"
      'Ado_datos02.Recordset!cobranza_nro_factura = IIf(TxtCmpbte2 = "", "0", Trim(TxtCmpbte2))
      'Ado_datos02.Recordset!cobranza_nro_autorizacion = IIf(TxtAutorizacion2 = "", "0", Trim(TxtAutorizacion2))
      'Ado_datos02.Recordset!poa_codigo = "3.1.2"
      If Ado_datos02.Recordset!estado_codigo_bco <> "APR" Then
            Ado_datos02.Recordset!estado_codigo_bco = "REG"
            Ado_datos02.Recordset!usr_codigo_bco = glusuario
      End If
      Ado_datos02.Recordset!Fecha_Registro = Format(Date, "dd/mm/yyyy")
      Ado_datos02.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos02.Recordset.Update
    db.CommitTrans
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    End If
    FraNavega2.Enabled = True
'      fraOpciones1.Visible = True
    FraGrabarCancelar2.Visible = False
    FrmDetalle.Enabled = True
    FrmCobranza.Enabled = True
'      FrmABMDet.Visible = True
    FrmABMDet2.Visible = True
    FrmCobros2.Enabled = False
'      BtnAprobar3.Visible = False
    DTPFechaCobro2A.Visible = False
    DataCombo2.Locked = True
    swnuevo = 0

End Sub

Private Sub BtnGrabarBen_Click()
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' and beneficiario_codigo = '" & dtc_codigo8.Text & "'  ", db, adOpenStatic
    If rs_datos10.RecordCount = 0 Then
        'abrir gc_edificio_vs_beneficiario
        db.Execute "INSERT INTO gc_edificio_vs_beneficiario (edif_codigo, beneficiario_codigo, estado_codigo, fecha_registro, usr_codigo) VALUES ('" & VAR_PROY3 & "', '" & dtc_codigo8.Text & "', 'APR', '" & Date & "', '" & glusuario & "')"
        'Beneficiario Personas Nat. y Juridicas Relacionadas al Edificio
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gv_edificio_vs_beneficiario where edif_codigo = '" & VAR_PROY3 & "' ", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        dtc_aux5.BoundText = dtc_codigo5.BoundText
        'FraGrabarCancelar.Enabled = True
'        lbl_nit_fac.Caption = dtc_codigo8.Text
'        lbl_benef_fac.Caption = dtc_desc8.Text

'        lbl_nit_fac.Visible = True
'        lbl_benef_fac.Visible = True
        
    Else
        MsgBox "Ya existe el Beneficiario relacionado, en: <<Facturado a Nombre de>>. Vuelva a intentar ...", , "Atención"
    End If
    FraGrabarCancelar.Enabled = True
    frm_benef.Visible = False
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

Private Sub BtnImprimir1_Click()
    Select Case SSTab1.Tab
        Case 0
          If Ado_datos01.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos01.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos01.Recordset!cobranza_codigo
            var_literal = Literal(CStr(Ado_datos01.Recordset!cobranza_programada_dol)) + " DOLARES "
            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
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
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
            'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            var_literal = Literal(CStr(Ado_datos.Recordset!cobranza_programada_dol)) + " DOLARES "
            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
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
            CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_dol.rpt"
            CryR01.WindowShowRefreshBtn = True
            CryR01.StoredProcParam(0) = Me.Ado_datos02.Recordset!venta_codigo
            CryR01.StoredProcParam(1) = Me.Ado_datos02.Recordset!cobranza_codigo
            var_literal = Literal(CStr(Ado_datos02.Recordset!cobranza_programada_dol)) + " DOLARES "
            CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
            CryR01.Formulas(2) = "correlcobro = '" & Ado_datos02.Recordset!cobranza_codigo & "' "
            '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
            iResult = CryR01.PrintReport
            If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
          Else
            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
          End If
    End Select

End Sub

Private Sub BtnImprimir2_Click()
    Call generarRepRecibo
End Sub

' Genera reporte de recibo
Private Sub generarRepRecibo()
    ' Verifica si codigo y numero son validos para recibo.
    'If Label38.Caption <> "" And TxtCmpbte2 <> "" And Label38.Caption <> "R-101" Then
    If Ado_datos.Recordset!doc_codigo_fac <> "R-101" Then
        Dim iResult As Integer
        Dim montoLiteral As String
        Dim Monto As Double
        Monto = 0
        If Ado_datos.Recordset!estado_codigo_fac = "REG" Then
            Set rs_aux1 = New ADODB.Recordset
            rs_aux1.CursorLocation = adUseClient
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103'  ", db, adOpenDynamic, adLockOptimistic
            If rs_aux1.RecordCount > 0 Then
                lbl_docnro = rs_aux1!correl_doc + 1
            End If
            db.Execute "update gc_documentos_respaldo set correl_doc = " & lbl_docnro.Text & " Where doc_codigo = 'R-103' "
            db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR', estado_codigo_bco = 'REG', doc_numero = " & lbl_docnro.Text & " Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
            
            db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
        End If
        If TxtMonto.Text = "0" Or TxtMonto.Text = "" Then
            MsgBox "No se puede emitir un Recibo con Monto cero, vuelva a intentar ..."
            Exit Sub
        Else
            Monto = TxtMonto.Text
        End If
        'If TxtCmpbte2 <> "" Then Monto = TxtCmpbte2
        'If TxtMonto.Text <> "" Then Monto = Monto + CInt(TxtDscto2.Text)
        montoLiteral = TxtMonto.Text          'Monto
        montoLiteral = Literal(montoLiteral)
        
        crRecibo.WindowShowPrintSetupBtn = True
        crRecibo.WindowShowRefreshBtn = True
        crRecibo.ReportFileName = App.Path & "\Reportes\Ventas\ar_recibo_oficial.rpt"
        crRecibo.StoredProcParam(0) = lbl_fac.Caption    ' codigo R-103
        crRecibo.StoredProcParam(1) = lbl_docnro.Text ' numero Recibo
        crRecibo.StoredProcParam(2) = "Bs. " + TxtMonto + "(" + montoLiteral + " BOLIVIANOS)" ' monto
        crRecibo.WindowState = crptMaximized
        iResult = crRecibo.PrintReport
        If iResult <> 0 Then
              MsgBox crRecibo.LastErrorNumber & " : " & crRecibo.LastErrorString, vbExclamation + vbOKOnly, "Error"
        End If
        db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
        db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
        db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset("cobranza_codigo") & " "
                
   Else
       MsgBox "No se puede generar reporte por falta de codigo y numero de recibo."
   End If
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
  'IMPRIME FACTURA
  db.Execute "update ao_ventas_cobranza set estado_codigo1 = 'FAC' Where cobranza_codigo = '18314' "
  db.Execute "update ao_ventas_cobranza set cobranza_nro_fac = '100' Where cobranza_codigo = '18314' "
  db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where cobranza_codigo = '18314' "
  db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = '18314' "
  db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '25/08/2020' Where cobranza_codigo = '18314' "
'If Ado_datos.Recordset.RecordCount > 0 And (dtc_aux5.Text <> "") Then
'  If (Ado_datos.Recordset!factura_impresa = "N") Then
'    If Ado_datos.Recordset!cobranza_total_bs >= 3000 And dtc_aux5.Text = "0" Then
'        MsgBox "No se puede Imprimir una Factura >= Bs.300, sin NIT, debe registrar el NIT del Beneficiario... ", , "Atención"
'    Else
'      If Ado_datos.Recordset!doc_codigo_fac = "R-101" Then
'        '===== ini GENERA EL CODIGO DE FACTURA ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from fc_dosificacion_docs where doc_codigo = 'R-101' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
'        'rs_aux1.Open "select * from fc_dosificacion_docs  where doc_codigo = 'R-101'  ", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            If Date > rs_aux1!dosifica_fecha_limite Then
'                MsgBox "No se puede EMITIR mas Facturas, la fecha límite de emisión EXPIRÓ, debe realizar una nueva Dosificación ... ", , "Atención"
'                Exit Sub
'            End If
'            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
'            correlv = Ado_datos.Recordset("venta_codigo")
'            nroventa = Ado_datos.Recordset("venta_codigo")
'            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            'VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
'            'VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
'            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_total_dol, 2)
'            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_total_bs, 2)
'            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
'            var_literal = Ado_datos.Recordset!Literal
'            VAR_FFAC = Format((Date), "DD/MM/YYYY")
'            VAR_CODTIPO = "REF"     'Tipo Comprobante (paralelo VAR_DOC)
'            VAR_DOC = "R-112"       'Doc. Respaldo
'            VAR_ETAPA = "FIN-02-02"
'            VAR_TCOMP = "RECAUDADO (FACTURACION)"
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
'            VAR_VTIPO = Ado_datos16.Recordset!venta_tipo
'            'CodigoContro = CodigoControl(NroFactura)
'            If Autorizacion <> "" And NitCi <> "" And Llave <> "" And VAR_BS2 <> "0" And rs_aux1!CORREL >= 0 Then
'                VAR_SW = 1
'            Else
'                VAR_SW = 0
'                MsgBox "Error en Autorizacion, NIT o Llave, Contactese con el Administrador y vuelva a intentar ...", , "Atención"
'                Exit Sub
'            End If
'            VAR_COD1 = CDbl(rs_aux1!CORREL) + 1
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!CORREL = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'                VAR_ANIO = Year(VAR_FFAC)
'                VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
'
'                'VAR_COD1 = "4083"
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
'                'VAR_CONTAB = "17094"
'                VAR_COD2 = rs_aux1!dosifica_autorizacion
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
'                db.Execute "update ao_ventas_cobranza set correl_contab = " & VAR_CONTAB & " Where ao_ventas_cobranza.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And ao_ventas_cobranza.cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
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
'                db.Execute "update ao_ventas_cobranza set cobranza_deuda_bs = '0', cobranza_deuda_dol = '0', cobranza_deuda_bs2 = '0', cobranza_deuda_dol2 = '0' Where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'
''                Ado_datos.Recordset!estado_codigo_fac = "APR"
''                Ado_datos.Recordset.Update
'                'INI QR
'                'sFile = "C:\Tmp\QRCode.bmp"
'                '1003579028
'                '& "|" & Format(Trim("0"), "###0.00") _
'                'dtc_aux5.Text
'                sFile = App.Path & "\CLIENTES\QRCode.bmp"
'                CadenaQ = Trim("1018533029") _
'                & "|" & Trim(VAR_COD1) _
'                & "|" & Trim(VAR_COD2) _
'                & "|" & Format(Trim(Date), "DD/MM/YYYY") _
'                & "|" & Format(Trim(VAR_BS2), "###0.00") _
'                & "|" & Format(Trim(VAR_BS2), "###0.00") _
'                & "|" & Trim(CodigoContro) _
'                & "|" & Trim(dtc_aux5.Text) _
'                & "|" & Trim("0") _
'                & "|" & Trim("0") _
'                & "|" & Trim("0") _
'                & "|" & Trim("0")
'
'                'CadenaQ = Trim(txtNitEmisor.Text) _
'                '& "|" & Trim(txtNumeroFactura.Text) _
'                '& "|" & Trim(txtNumeroAutorizacion.Text) _
'                '& "|" & Format(Trim(txtFechaEmision.Text), "DD/MM/YYYY") _
'                '& "|" & Format(Trim(txtImporteCompra.Text), "###0.00") _
'                '& "|" & Format(Trim(txtFiscal.Text), "###0.00") _
'                '& "|" & Trim(txtCodigoControl.Text) _
'                '& "|" & Trim(txtNitComprador.Text) _
'                '& "|" & Trim(txtImporteICE.Text) _
'                '& "|" & Trim(txtGravadas.Text) _
'                '& "|" & Trim(txtNoFiscal) _
'                '& "|" & Trim(TxtDescuento)
''                MsgBox CadenaQ
'                FastQRCode CadenaQ, sFile
'                Set Picture1.Picture = LoadPicture(sFile)
'                'FIN QR
'                'Call IMPRIME_FACTURA
'                Call IMPRIME_QR
'                'MsgBox CadenaQ
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'                db.Execute "UPDATE co_diario SET co_diario.estado_codigo = co_comprobante_m.estado_codigo FROM co_diario INNER JOIN co_comprobante_m ON co_diario.Cod_Comp =co_comprobante_m.Cod_Comp where co_diario.estado_codigo Is Null "
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'
'
''        '===== ini GENERA NRO. AUTORIZACION DE FACTURA ====
''        Set rs_aux1 = New ADODB.Recordset
''        rs_aux1.CursorLocation = adUseClient
''        If rs_aux1.State = 1 Then rs_aux1.Close
''        rs_aux1.Open "select * from fc_Correl  where tipo_tramite = 'FAC_AUTORIZA'", db, adOpenDynamic, adLockOptimistic
''        If rs_aux1.RecordCount > 0 Then
''          VAR_COD2 = CDbl(rs_aux1!numero_correlativo)
''          'rs_aux1!numero_correlativo = Trim(Str(VAR_COD2))
''          'rs_aux1.Update
''        End If
''        If rs_aux1.State = 1 Then rs_aux1.Close
''        '===== fin TERMINA GENERACION NRO. AUTORIZACION DE FACTURA =====
'
''        Dim iResult As Variant  ', i%, y%
''        CryF01.ReportFileName = App.Path & "\reportes\ventas\ar_R-101_factura.rpt"
''        CryF01.WindowShowRefreshBtn = True
''        CryF01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''        CryF01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''        CryF01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
''
''        CryF01.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
''        CryF01.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
''        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
''        iResult = CryF01.PrintReport
''        If iResult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
'
'        TxtCmpbte = VAR_COD1
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      Else
'        Call generarRepRecibo
'      End If
'      If Ado_datos.Recordset!doc_codigo_fac = "R-103" Then
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'        '===== ini GENERA EL CODIGO DE RECIBO ====
'        Set rs_aux1 = New ADODB.Recordset
'        rs_aux1.CursorLocation = adUseClient
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
'            correlv = Ado_datos.Recordset("venta_codigo")
'            nroventa = Ado_datos.Recordset("venta_codigo")
'            NRO_COBR = Me.Ado_datos.Recordset!cobranza_codigo
'            VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'            VAR_CITE = Ado_datos16.Recordset!unidad_codigo_ant
'            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
'            VAR_GLOSA = Trim(Ado_datos.Recordset!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
'            VAR_DOL2 = Round(Ado_datos.Recordset!cobranza_deuda_dol, 2)
'            VAR_BS2 = Round(Ado_datos.Recordset!cobranza_deuda_bs, 2)
'            'VAR_CTA = IIf(Ado_datos.Recordset!Cta_Codigo = "", "NN", Ado_datos.Recordset!Cta_Codigo)
'            var_literal = Ado_datos.Recordset!Literal
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
'            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Recibo Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
'                rs_aux1!correl_doc = Trim(Str(VAR_COD1))
'                rs_aux1.Update
'                'GENERA CORREL NOTA DEBITO POR DEPTO INI
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
'                Call IMPRIME_RECIBO
'                'If VAR_TIPOV = "C" Then
'                    Call Contabiliza_venta
'                'End If
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
'        End If
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        '===== fin TERMINA GENERACION DE FACTURA =====
'        TxtCmpbte = VAR_COD1
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'      End If
'
'    End If
'  Else
'        MsgBox "La Factura Nro. " + Ado_datos.Recordset!cobranza_nro_factura + " ya fue Impresa", , "Atención"
'        'Call IMPRIME_FACTURA
'        If (Ado_datos.Recordset("estado_codigo_sol") = "APR" And Ado_datos.Recordset("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'  End If
'Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'End If
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
'    'IMPRIMIR FACTURA con QR
'    'Dim Exel As Object
'    'Set Exel = CreateObject("Excel.Application")
'    'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
'    'Exel.Visible = True
'    Call CmdFoto_Click
'    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'
'    Picture2.AutoRedraw = True
'    Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'    ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'    ' MsgBox CadenaQr
'    FastQRCode CadenaQr, ImagenQr
'    Picture1.AutoRedraw = True
'    Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'    Clipboard.Clear
'    Clipboard.SetData Picture2.Image
''    Exel.Application.Range("a2").Select
''    Exel.Application.ActiveSheet.Paste
'
'    Dim iResult As Variant  ', i%, y%
'    sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'        If VAR_COD4 = "DNMAN" Then
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man.rpt"
'        Else
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
'        End If
'    Else
'        CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
'    End If
'        CryQ01.WindowShowRefreshBtn = True
'        CryQ01.StoredProcParam(0) = glGestion       'Me.Ado_datos.Recordset!ges_gestion
'        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
'        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
'        'var_literal = "-"   'Ado_datos.Recordset!Literal
'        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
'        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
'        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
'        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
'        iResult = CryQ01.PrintReport
'        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"

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

Private Sub BtnImprimir5_Click()
  'RE-IMPRIME FACTURA
  If Ado_datos.Recordset.RecordCount > 0 And (Ado_datos.Recordset!cobranza_nro_factura > 0) Then         'And (dtc_aux5.Text <> "")
        gestion0 = Ado_datos.Recordset!ges_gestion
        nroventa = Ado_datos.Recordset!venta_codigo
        NRO_COBR = Ado_datos.Recordset!cobranza_codigo
        VAR_COD1 = Ado_datos.Recordset!cobranza_nro_factura
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'        'Dim Exel As Object
'        'Set Exel = CreateObject("Excel.Application")
'        'Exel.Workbooks.Open "c:\tmp\Factura.xlt", , , , "123", "123"
'        'Exel.Visible = True
'        Call CmdFoto_Click
'        'ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'
'        Picture2.AutoRedraw = True
'        Picture2.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
'
'        ImagenQr = App.Path & "\CLIENTES\QRCode.bmp"
'        ' MsgBox CadenaQr
'        FastQRCode CadenaQr, ImagenQr
'        Picture1.AutoRedraw = True
'        Picture1.PaintPicture LoadPicture(ImagenQr), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
'        Clipboard.Clear
'        Clipboard.SetData Picture2.Image
'    '    Exel.Application.Range("a2").Select
'    '    Exel.Application.ActiveSheet.Paste
    
        Dim iResult As Variant  ', i%, y%
        sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
        
        If sino = vbYes Then
            If VAR_COD4 = "DNMAN" Then
                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man.rpt"
            Else
                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
            End If
        Else
            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
        End If
        CryQ01.WindowShowRefreshBtn = True
        CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        var_literal = Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
  Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnModificar_Click()
  codigo_doc = lbl_fac.Caption
  If Ado_datos.Recordset.RecordCount > 0 Then
    If codigo_doc <> "R-101" Then
         Call cambiarEtiquetaFactura
         Dim Cmd1 As ADODB.Command
         Dim rs  As ADODB.Recordset
         Set Cmd1 = New ADODB.Command
         Set rs = New ADODB.Recordset
        
         Cmd1.ActiveConnection = db 'sqlServer
         Cmd1.CommandType = adCmdStoredProc
         Cmd1.CommandText = "ap_genera_codigoregistro"
         Set Parm1 = Cmd1.CreateParameter("@codigo_doc", adVarChar, adParamInput, 200, codigo_doc)
         Cmd1.Parameters.Append Parm1
         rs.Open Cmd1
         rs.MoveFirst
         TxtCmpbte.Text = rs!Codigo
         rs.Close
    Else
        Call cambiarEtiquetaFactura
    End If
    If (Ado_datos.Recordset!estado_codigo_sol = "APR" And Ado_datos.Recordset!estado_codigo_fac = "REG") And (Ado_datos16.Recordset!venta_tipo = "E" Or Ado_datos16.Recordset!venta_tipo = "V" Or Ado_datos16.Recordset!venta_tipo = "C" Or Ado_datos16.Recordset!venta_tipo = "L") Then
      FraNavega.Enabled = False
'      fraOpciones.Visible = False
      FraGrabarCancelar.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      'swgrabar = 0
      swnuevo = 2
      Txt_tdc.Text = GlTipoCambioMercado    'GlTipoCambioOficial
      SSTab1.Tab = 1
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = False
      FrmCobros.Visible = True
      FrmCobros.Enabled = True
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      
'      BtnImprimir2.Visible = False
      BtnImprimir3.Visible = False
'      CmdFoto.Visible = False
      If Ado_datos.Recordset!factura_impresa = "N" Then
      '    sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
      '    If sino = vbYes Then
              'DTPFechaProg.Visible = True
'              DTPFechaCobro.Visible = True
              DTPFechaCobro.Value = Date
'              Lbl_nombre_fac.Caption = "Factura a Nombre de:"
'              lbl_fechas.Caption = "Fecha de Cobranza"
              TxtCmpbte.Text = "0"
      '        Txt_parche.Visible = False      '&H80000013&
      '        'dtc_desc2A.BackColor = &H80000013
      '    Else
      '        DTPFechaProg.Visible = True
      '        DTPFechaCobro.Visible = False
      '        Lbl_nombre_fac.Caption = "Cliente :"
      '        lbl_fechas.Caption = "Fecha Programada de Cobranza"
      '        Txt_parche.Visible = True       '&H80000005&
      '        'dtc_desc2A.BackColor = &H80000005
      '    End If
      Else
      '    DTPFechaProg.Visible = True
      '    DTPFechaCobro.Visible = False
      '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
      End If
      'TxtMonto.Text = CDbl(TxtDsctoTot)
      TxtMonto.SetFocus
    Else
      MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModificar1_Click()
  If Ado_datos01.Recordset.RecordCount > 0 Then
    If Ado_datos01.Recordset!estado_codigo_sol = "REG" Then
      SSTab1.Tab = 0
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega1.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar1.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros1.Visible = True
      FrmCobros1.Enabled = True
      swnuevo = 2
      DTPfechasol.Value = Date
      If TxtMonto1.Text = "0" Or TxtMonto1.Text = "0.00" Then
        TxtMonto1.Text = TxtDsctoTot1.Caption
      End If
      If TxtDscto1.Text = "0" Or TxtDscto1.Text = "0.00" Then
        TxtDscto1.Text = txtProg_Bs.Caption
      End If
      TxtMonto1.SetFocus
    Else
      MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

Private Sub BtnModificar2_Click()
  If Ado_datos02.Recordset.RecordCount > 0 Then
    If Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG" And (glusuario = "CNUÑEZ" Or glusuario = "ADMIN") Then
        DataCombo2.Locked = False
    End If
    If Ado_datos02.Recordset!estado_codigo_fac = "APR" And Ado_datos02.Recordset!estado_codigo_bco = "REG" Then
      SSTab1.Tab = 2
      SSTab1.TabEnabled(0) = False
      SSTab1.TabEnabled(1) = False
      SSTab1.TabEnabled(2) = False
      'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
      FraNavega2.Enabled = False
'      fraOpciones1.Visible = False
      FraGrabarCancelar2.Visible = True
      FrmDetalle.Enabled = False
      FrmCobranza.Enabled = False
      FrmABMDet.Visible = False
      FrmABMDet2.Visible = False
      FrmCobros2.Visible = True
      FrmCobros2.Enabled = True
      swnuevo = 2
'      BtnAprobar3.Visible = True
      DTPFechaCobro2A.Visible = True
      If DTPFechaCobro2.Value = "01/01/1900" Then
        DTPFechaCobro2A.Value = Date
      Else
        DTPFechaCobro2A.Value = DTPFechaCobro2.Value
      End If
      'DTPFechaCobro02.Value = Date
      'Txt_deposito.Text = "0"
      TxtMonto02.SetFocus
    Else
      If Ado_datos02.Recordset!estado_codigo_bco = "APR" And Ado_datos02.Recordset!estado_codigo = "REG" And (glusuario = "CNUÑEZ" Or glusuario = "ADMIN") Then
          SSTab1.Tab = 2
          SSTab1.TabEnabled(0) = False
          SSTab1.TabEnabled(1) = False
          SSTab1.TabEnabled(2) = False
          'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
          FraNavega2.Enabled = False
    '      fraOpciones1.Visible = False
          FraGrabarCancelar2.Visible = True
          FrmDetalle.Enabled = False
          FrmCobranza.Enabled = False
          FrmABMDet.Visible = False
          FrmABMDet2.Visible = False
          FrmCobros2.Visible = True
          FrmCobros2.Enabled = True
          swnuevo = 2
          
          'DTPFechaCobro2.Value = Date
          'DTPFechaCobro02.Value = Date
          'Txt_deposito.Text = "0"
          TxtMonto02.SetFocus
      Else
            MsgBox "No se puede editar, porque el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
      End If
    End If
  Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

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

'Private Sub Cmd_Cliente_Click()
'    glPersNew = "P"
'    frmBeneficiario.Show 'vbModal
'End Sub

Private Sub CmdCancelaCobro_Click()
End Sub

Private Sub BtnModDetalle2_Click()
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

    'If Ado_datos.Recordset.RecordCount > 0 Then
'            Dim iResult As Variant  ', i%, y%
            'CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_cobranzas_facturadas.rpt"
            CryF02.ReportFileName = App.Path & "\reportes\ventas\ar_lista_diaria_facturas.rpt"
            CryF02.WindowShowRefreshBtn = True
            'CryF02.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
            'CryF02.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
            'CryF02.Formulas(1) = "literalcobro = '" & Ado_datos.Recordset!Literal & "' "
            'CryF02.Formulas(2) = "correlcobro = '" & Ado_datos.Recordset!cobranza_codigo & "' "
        CryF02.Formulas(1) = "titulo = 'LISTADO DE FACTURACION' "
        CryF02.Formulas(2) = "subtitulo = 'MODULO COBRANZAS' "
            iResult = CryF02.PrintReport
            If iResult <> 0 Then MsgBox CryF02.LastErrorNumber & " : " & CryF02.LastErrorString, vbCritical, "Error de impresión"
'          Else
'            MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
     'End If

End Sub

Private Sub BtnDesAprobar_Click()

 If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_codigo_fac = "REG" And Ado_datos.Recordset!estado_codigo1 = "APR" Then
    'If Ado_datos.Recordset!estado_codigo_fac = "REG" And Ado_datos.Recordset!factura_impresa = "N" Then
'        Ado_datos.Recordset!estado_codigo_sol = "REG"
'        Ado_datos.Recordset!estado_codigo_fac = "REG"
'        Ado_datos.Recordset.Update
          db.Execute "update ao_ventas_cobranza set estado_codigo1 = 'REG' Where cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
    Else
        MsgBox "No se puede DEVOLVER, el registro ya fue FACTURADO, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
 End If
End Sub

'Private Sub CmdDetallePoa_Click()
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'   MsgBox "No Existen Registros ", vbInformation, "Formulario 11"
'  Else
'    marca1 = Ado_datos.Recordset.BookMark
'    FrmPoasCapturaALB.Lblformulario = "F11"
'    FrmPoasCapturaALB.lblges_gestion = Ado_datos.Recordset!ges_gestion
'    FrmPoasCapturaALB.lblcodigo_unidad = Ado_datos.Recordset!codigo_unidad
'    FrmPoasCapturaALB.lblcodigo_solicitud = Ado_datos.Recordset!codigo_solicitud
'    FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'Ado_datos.Recordset!tipoben_codigo
'    FrmPoasCapturaALB.Show vbModal
'  If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
'    '
'  Else
'    Ado_datos.Refresh
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  End If
'End Sub

'Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            txtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
'End Sub

Private Sub Contabiliza_venta()
'    Call graba_proyecto
    If VAR_SW = 1 Then
        Call graba_ingreso
    End If
    'If VAR_SW = 1 Then
        Set rstdestino = New ADODB.Recordset
        If VAR_TIPOV = "L" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
        Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
        End If
        If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo

            'Modificar con CASE WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW MAY-2015
            If VAR_SW = 1 Then
                VAR_TSOL = VAR_TIPOS
            Else
                VAR_TSOL = rstdestino!solicitud_tipo
                VAR_TIPOS = rstdestino!solicitud_tipo
                VAR_PARTIDA = rstdestino!rubro_codigo
            End If
        Else
        
        End If
    'End If
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    fte_codigo1 = VAR_FTE
    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI", "DEY"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
        Case "REC"
            If VAR_MONEDA = "BOB" Then
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and SubCta_Deb1 = '01' ", db, adOpenKeyset, adLockReadOnly
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            Else
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  and SubCta_Deb1 = '02' ", db, adOpenKeyset, adLockReadOnly
            End If
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
                        
            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                    'JQA FEB-2016
                    'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "REF"
            If VAR_VTIPO = "L" Then     'Importación Directa
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " AND correlativo <> '6' ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO = "V" Then     'Facturacion Local
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_VTIPO <> "L" And VAR_VTIPO <> "V" Then       'Mant, Rep, Inst, etc.
                If rstdestino.State = 1 Then rstdestino.Close
                rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = '" & VAR_CODTIPO & "' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "  ", db, adOpenKeyset, adLockReadOnly
                If rstdestino.RecordCount > 0 Then
                    j = rstdestino.RecordCount
                Else
                  MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
                  Exit Sub
                End If
            End If
            If VAR_JQ = "" Then
                If rs_aux1.State = 1 Then rs_aux1.Close
                rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
                'rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
                If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
                  If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                    MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                     'Exit Sub
                  End If
                End If
                If rs_aux1.State = 1 Then rs_aux1.Close
            End If
        Case "DYR"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
        Case "DES"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "ANI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "DVI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If
            
            '' 02/07/2014 VERIFICAR
            'If rstdestino.State = 1 Then rstdestino.Close
            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
            'If rstdestino2.State = 1 Then rstdestino2.Close
            'rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            'If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
            '  MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
            '  Exit Sub
            'End If
        Case Else
            MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
            If rstdestino.State = 1 Then rstdestino.Close
            Exit Sub
    End Select
    'If rstdestino.State = 1 Then rstdestino.Close
    '**** FIN VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************

    Dim cta_deb1 As String
    Dim Subcta_deb11 As String
    Dim Subcta_deb21 As String

    Dim cta_credito1 As String
    Dim Subcta_cred11 As String
    Dim Subcta_cred21 As String

    Dim cod_ant As Integer
    Dim org_ant As String

    v_Tipo_Comp(1, 1) = VAR_CODTIPO
    
    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
    '========================================
    '==== verifica si ya fue contabilizado
      yacontabilizo = 0
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '" & VAR_CODANT & "' and org_codigo = '" & VAR_ORG & "' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      'rs_aux2.Open "select * from co_comprobante_m where Cod_trans = '2' and org_codigo = '111' and tipo_comp = '" & VAR_CODTIPO & "' AND estado_codigo = 'APR'", db, adOpenKeyset, adLockOptimistic
      If rs_aux2.RecordCount > 0 Then
        ' revisar para validar mejor si YA contabilizo !!
        'yacontabilizo = 1
        yacontabilizo = 0
      Else
        yacontabilizo = 0
      End If
      If yacontabilizo = 1 Then
        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
        Var_Comp = rs_aux2!Cod_Comp
      Else
        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
        Set rstCodComp = New ADODB.Recordset
        rstCodComp.CursorLocation = adUseClient
        If rstCodComp.State = 1 Then rstCodComp.Close
        rstCodComp.Open "select * from fc_Correl where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
        
        'Correlativo por Mes y Tipo de Comprobante
        Set rs_aux14 = New ADODB.Recordset
        SQL_FOR = "select numero_correlativo, tipo_tramite FROM fc_correl WHERE (cta_codigo1 = '" & Trim(VAR_MES) & "' and cta_codigo2 = '" & VAR_DOC & "' ) "
        rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux14.RecordCount > 0 Then
              rs_aux14!numero_correlativo = rs_aux14!numero_correlativo + 1
              VAR_COMPM = rs_aux14!numero_correlativo    'VAR_DOCR
              rs_aux14.Update
        End If
        'R-112, R-110, R-111
         Set rs_aux14 = New ADODB.Recordset
          If rs_aux14.State = 1 Then rs_aux14.Close
          SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "'  "  ''R-112' "          '  '" & txt_codigo1 & "' "
          rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
          If rs_aux14.RecordCount > 0 Then
                rs_aux14!correl_doc = rs_aux14!correl_doc + 1
                'VAR_COMPM = rs_aux14!correl_doc
                rs_aux14.Update
          End If
        '===== fin TERMINA GENERACION DE COMPROBANTE =====

      '==== ini registro co_comprobante_m

        rs_aux2.AddNew
        rs_aux2("cod_comp") = Var_Comp
      End If
    '========================================
    'anterior
    '      If rstdestino.State = 1 Then rstdestino.Close
    '      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
    '      If rstdestino.RecordCount > 0 Then
    '      End If
    '      rstdestino.AddNew
    
    '      rstdestino("cod_comp") = Var_Comp
    'anterior
      rs_aux2("Tipo_Comp") = VAR_CODTIPO        'v_Tipo_Comp(1, i)
      rs_aux2("cod_trans") = VAR_CODANT
      rs_aux2("org_codigo") = VAR_ORG
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_transacion") = Date
      'Else
        rs_aux2!Fecha_transacion = IIf(VAR_FFAC = "", Date, VAR_FFAC)
      'End If
      rs_aux2("mes_trasaccion") = UCase(MonthName(Month(Date)))
      rs_aux2("ges_gestion") = IIf(gestion0 = "", Year(Date), gestion0)  'glGestion
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = VAR_TCOMP + "- " + VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4           'Ado_datos16.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = VAR_SOL         'Ado_datos16.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE
      'rs_aux2!Cobranza_aux = NRO_COBR
      rs_aux2("proceso_codigo") = Left(VAR_ETAPA, 3)        '"FIN"
      rs_aux2("subproceso_codigo") = Left(VAR_ETAPA, 6)     '"FIN-02"
      rs_aux2("etapa_codigo") = VAR_ETAPA
      
      rs_aux2("clasif_codigo") = "ADM"
      'rs_aux2("doc_codigo") = "R-112"
      rs_aux2("doc_codigo") = VAR_DOC       '"R-110" o "R-112"
      rs_aux2("doc_numero") = VAR_COMPM         'Var_Comp   VAR_COMPM
      rs_aux2("pro_codigo_det") = VAR_PROY2
    
      rs_aux2("estado_codigo") = "APR"
      rs_aux2("usr_codigo") = glusuario
      rs_aux2!venta_compra = nroventa
      rs_aux2!cobranza_pago = NRO_COBR
      'rs_aux2!Factura_cheque= NroFactura
      'If yacontabilizo = 0 Then
      '  rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
      '  'rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
      If VAR_FFAC = "" Then
        VAR_FFAC = Date
      End If
        rs_aux2("Fecha_registro") = Format(VAR_FFAC, "dd/mm/yyyy")
      'End If
      rs_aux2.Update
      db.Execute "UPDATE co_comprobante_m SET edificio = gc_edificaciones.edif_descripcion FROM co_comprobante_m INNER JOIN gc_edificaciones ON co_comprobante_m.pro_codigo_det =gc_edificaciones.edif_codigo where co_comprobante_m.edificio Is Null "

      db.Execute "UPDATE co_comprobante_m SET cliente = gc_beneficiario.beneficiario_denominacion FROM co_comprobante_m INNER JOIN gc_beneficiario ON co_comprobante_m.beneficiario_codigo  =gc_beneficiario.beneficiario_codigo where co_comprobante_m.cliente Is Null "
    
      db.Execute "UPDATE co_comprobante_m SET departamento = gc_departamento.depto_descripcion FROM co_comprobante_m INNER JOIN gc_departamento ON LEFT(co_comprobante_m.pro_codigo_det,1)  =gc_departamento.depto_codigo where co_comprobante_m.departamento Is Null "

      If VAR_TCOMP = "REF" Then
        db.Execute "UPDATE co_comprobante_m SET glosa_contab = 'Fac: ' NroFactura + ' - '+ unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      Else
        db.Execute "UPDATE co_comprobante_m SET glosa_contab = unidad_codigo + ' -Edif: ' + rtrim(edificio) + ' - Benef: ' + rtrim(cliente) + ' - ' + departamento + ' - ' + right(glosa,50) where co_comprobante_m.glosa_contab is null "
      End If
      '==== fin registro co_comprobantre_m

    Dim d_cta_nombre_1 As String
    Dim d_aux1_1 As String
    Dim d_aux2_1 As String
    Dim d_aux3_1 As String
    Dim h_cta_nombre_1 As String
    Dim h_aux1_1 As String
    Dim h_aux2_1 As String
    Dim h_aux3_1 As String
    'If rstdestino.State = 1 Then rstdestino.Close
    
    For i = 1 To j
'    ' nuevo ini
      
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")
        
        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
        
        VAR_PORC = rstdestino!porcentaje
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2
    
        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
        
        VAR_PORC = rstdestino!porcentaje
      End If
      
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
        VAR_DCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = RTrim(rs_aux1("NombreCta"))
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
        VAR_HCORR = rs_aux1("correl")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMD = rs_aux1("NombreCta")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        VAR_NOMH = rs_aux1("NombreCta")
      End If
    ' nuevo fin
    
      '===== ini registra CO_diaRIO =========
      Set rstdestino2 = New ADODB.Recordset
      If rstdestino2.State = 1 Then rstdestino2.Close
      rstdestino2.Open "select * from co_diario where Cod_Comp = " & Var_Comp, db, adOpenKeyset, adLockOptimistic
      'If rstdestino2.RecordCount > 0 Then
      '  MsgBox "Ya Existe el asiento, se reemplazará con los nuevos datos..."
      'Else
        rstdestino2.AddNew
        rstdestino2("Cod_Comp") = Var_Comp
      'End If
        rstdestino2("Cod_Comp_Detalle") = rstdestino2.RecordCount
      'rstdestino2("Tipo_Comp") = "DEI"   'v_Tipo_Comp(1, i)
      'rstdestino2("Cod_Comp_C") = Var_Comp
      'If v_Tipo_Comp(1, i) = "DEI" Or v_Tipo_Comp(1, i) = "REC" Then
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REF") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        rstdestino2("NOMCTADEBE") = VAR_NOMD
        rstdestino2("d_Correl") = VAR_DCORR
        ' para Aux1
        ' ini PARA EL FUTURO ******** REVISAR
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        SQL_FOR = "select * from cc_tipo_auxiliar where aux = '" & d_aux1_1 & "' "
'        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount > 0 Then
'            Set rs_aux1 = New ADODB.Recordset
'            If rs_aux1.State = 1 Then rs_aux1.Close
'            SQL_FOR = "select * from " + rs_aux4!NombreTabla + " where " + rs_aux4!nombre_codigo + " = " + VAR_COD1
'            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux1.RecordCount > 0 Then
'        Else
'        End If
        ' fin PARA EL FUTURO ******** REVISAR
        Select Case d_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(d_aux1_1, CStr(VAR_BENEF))    'DESAUX =
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                Call DESCAUX(d_aux1_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(d_aux1_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux1_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux1_1, rstdestino2!D_Cta_Aux1)
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                Call DESCAUX(d_aux1_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux1 = DESAUX
        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(d_aux2_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                Call DESCAUX(d_aux2_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(d_aux2_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux2_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux2_1, rstdestino2!D_Cta_Aux2)
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                Call DESCAUX(d_aux2_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux2 = DESAUX
        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(d_aux3_1, CStr(VAR_BENEF))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                Call DESCAUX(d_aux3_1, CStr(VAR_CTA))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(d_aux3_1, CStr(VAR_PROY2))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(d_aux3_1, CStr(VAR_COD4))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(d_aux3_1, rstdestino2!D_Cta_Aux3)
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                Call DESCAUX(d_aux3_1, CStr(VAR_ORG))
                rstdestino2("D_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!D_Des_Aux3 = DESAUX
        
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR

        'VAR_PORC Definido en el Relacionador
        If VAR_PORC = "0.87" Then
            rstdestino2("D_MontoBs") = VAR_87       'VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_87 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
        End If
        If VAR_PORC = "0.13" Then
            rstdestino2("D_MontoBs") = VAR_13       'VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_13 * GlTipoCambioOficial  'VAR_DOL2 * VAR_PORC
        End If
        If VAR_PORC <> "0.87" And VAR_PORC <> "0.13" Then
            rstdestino2("D_MontoBs") = VAR_BS2 * VAR_PORC
            rstdestino2("D_MontoDl") = VAR_DOL2 * VAR_PORC
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        rstdestino2("NOMCTAHABER") = VAR_NOMH
        rstdestino2("h_Correl") = VAR_HCORR
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                Call DESCAUX(h_aux1_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                Call DESCAUX(h_aux1_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                Call DESCAUX(h_aux1_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux1_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux1_1, rstdestino2!H_Cta_Aux1)
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                Call DESCAUX(h_aux1_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux1") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux1 = DESAUX
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                Call DESCAUX(h_aux2_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                Call DESCAUX(h_aux2_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                Call DESCAUX(h_aux2_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux2_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux2_1, rstdestino2!H_Cta_Aux2)
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                Call DESCAUX(h_aux2_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux2") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux2 = DESAUX
        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                Call DESCAUX(h_aux3_1, CStr(VAR_BENEF))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                Call DESCAUX(h_aux3_1, CStr(VAR_CTA))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                Call DESCAUX(h_aux3_1, CStr(VAR_PROY2))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4            'Ado_datos.Recordset("unidad_codigo")
                Call DESCAUX(h_aux3_1, CStr(VAR_COD4))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = Left(VAR_PROY2, 1)           '"LA_PAZ"
                Call DESCAUX(h_aux3_1, rstdestino2!H_Cta_Aux3)
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                Call DESCAUX(h_aux3_1, CStr(VAR_ORG))
                rstdestino2("H_Des_Aux3") = DESAUX
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                DESAUX = ""
        End Select
        rstdestino2!H_Des_Aux3 = DESAUX
        
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = VAR_BS2 * VAR_PORC
        rstdestino2("H_MontoDl") = VAR_DOL2 * VAR_PORC
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
        'desafecta un devengado
        rstdestino2("D_Cuenta") = cta_credito1
        rstdestino2("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_cred11
        rstdestino2("D_SubCta2") = Subcta_cred21
        rstdestino2("D_Aux1") = h_aux1_1
        rstdestino2("D_Aux2") = h_aux2_1
        rstdestino2("D_Aux3") = h_aux3_1
'        rstdestino2("D_Cta_Aux1") = "VESCT"
        Select Case h_aux1_1
            Case "01"
                rstdestino2("D_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = ""
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
        End Select
        
        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = ""
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
        End Select
        
        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("D_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = ""
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("D_Cambio") = GlTipoCambioMercado

        rstdestino2("H_Cuenta") = cta_deb1
        rstdestino2("H_Nombre") = d_cta_nombre_1  ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_deb11
        rstdestino2("H_SubCta2") = Subcta_deb21
        rstdestino2("H_Aux1") = d_aux1_1
        rstdestino2("H_Aux2") = d_aux2_1
        rstdestino2("H_Aux3") = d_aux3_1
'        rstdestino2("H_Cta_Aux1") = "VESCT"
        Select Case d_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux1") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = ""
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
        End Select
        
        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux2") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = ""
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
        End Select
        
        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
            Case "04"
                rstdestino2("H_Cta_Aux3") = Ado_datos.Recordset("unidad_codigo")
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = ""
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
        rstdestino2("H_Cambio") = GlTipoCambioMercado
        rstdestino2!cobranza_pago = NRO_COBR
      End If

'      '==== INI DVI ====
'      If (VAR_CODTIPO = "DVI") Then
'        rstdestino2("D_Cuenta") = cta_deb1
''        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino2("D_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("D_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("D_Cambio") = GlTipoCambioMercado
'        rstdestino2("H_Cuenta") = cta_credito1
''        rstdestino2("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("H_SubCta1") = Subcta_cred11
'        rstdestino2("H_SubCta2") = Subcta_cred21
'        rstdestino2("H_Aux1") = h_aux1_1
'        rstdestino2("H_Aux2") = h_aux2_1
'        rstdestino2("H_Aux3") = h_aux3_1
'        'rstdestino2("H_Cta_Aux1") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
''        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino2("H_MontoBs") = IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)
'        rstdestino2("H_MontoDl") = IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)
'        rstdestino2("H_Cambio") = GlTipoCambioMercado
'      End If
'      '==== FIN DVI ====
      rstdestino2("Usr_codigo") = glusuario
      'If yacontabilizo = 0 Then
      '  rstdestino2("Fecha_registro") = Date
      ''  rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      'Else
        rstdestino2("Fecha_registro") = VAR_FFAC
      'End If
      
      rstdestino2.Update
      If rstdestino2.State = 1 Then rstdestino2.Close
      '======= fin registra co_diario ==========
      rstdestino.MoveNext
    Next i
    '======= inI Actualiza campos de estatus de ingresos ==========
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '" & correlativo1 & "' and org_codigo = '" & VAR_ORG & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
'    rstdestino.MoveFirst
'    If Not (rstdestino.EOF) Then
'      rstdestino("estado_aprobacion") = "S"
'        If VAR_CODTIPO = "DEI" Then
'          rstdestino("estado_devengado") = "S"
'        End If
'        If VAR_CODTIPO = "REC" Then
'          rstdestino("estado_recaudado") = "S"
'        End If
'        If VAR_CODTIPO = "DYR" Then
'          rstdestino("estado_devengado") = "S"
'          rstdestino("estado_recaudado") = "S"
'        End If
'
'        If VAR_CODTIPO = "DES" Then
'          rstdestino("estado_desafectado") = "S"
'        End If
'        If VAR_CODTIPO = "ANI" Then
'          rstdestino("estado_anulado") = "S"
'        End If
'        If VAR_CODTIPO = "DVI" Then
'          rstdestino!estado_desafectado = "S"
'          rstdestino!estado_anulado = "S"
'        End If
'       rstdestino.Update
'       If rstdestino.State = 1 Then rstdestino.Close
'    End If
    '======= fin Actualiza campos de estatus de ingresos ==========
    ' AAAAAAAAAQQQQQQQQQQQUUUUUUUUUUUIIIIIIIIIII
    cod_ant = 0
    org_ant = ""
    '======= ini Actualiza el monto recaudado  ==========
    If (VAR_CODTIPO = "REC") Then
      '      If rstdestino.State = 1 Then rstdestino.Close
      '      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
      '        cod_ant = rstdestino("ingreso_codigo_anterior")
      '        org_ant = rstdestino("org_codigo")
      '      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = '2' and org_codigo = '111' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + VAR_DOL2
          rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") + VAR_BS2
          rstdestino.Update
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If

    If (VAR_CODTIPO = "DES") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
'      Print VAR_CODANT
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
'        org_ant = rstdestino("org_codigo")
'      End If

      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DEI" Then 'And VAR_CODTIPO = "DES"
'          rstdestino!estado_desafectado = "S" 02/07/01
          rstdestino!estado_codigo = "DES"
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        Else
          rstdestino("estado_codigo") = "DES"
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
          cod_ant = IIf(IsNull(rstdestino("ingreso_codigo_anterior")), 0, rstdestino("ingreso_codigo_anterior"))
          org_ant = rstdestino("org_codigo")
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
          rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - VAR_DOL2
            rstdestino("monto_recaudado_bolivianos") = rstdestino("monto_recaudado_bolivianos") - VAR_BS2
          End If
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        End If
      End If
    End If

    If (VAR_CODTIPO = "ANI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "REC" Then
'          rstdestino("estado_desafectado") = ""
          rstdestino("estado_codigo") = "ANI"
'          rstdestino("estado_devengado") = "S" 02/07/01
'          rstdestino("estado_anulado") = ""
'          rstdestino("codigo_tipo") = "DEI" 02/07/01
          rstdestino("monto_recaudado_dolares") = 0
        End If
      End If
      rstdestino.Update
'      Print rstdestino!ingreso_codigo_anterior
'      Print rstdestino!monto_recaudado
      cod_ant = 0
      org_ant = ""
      
      'Call f_actual_rec(rstdestino!org_codigo, rstdestino!ingreso_codigo_anterior)
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    If (VAR_CODTIPO = "DVI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        rstdestino!estado_codigo = "DVI"
      End If
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza el monto recaudado  ==========

    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    If VAR_CODTIPO = "REC" Or VAR_CODTIPO = "DYR" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    If VAR_CODTIPO = "ANI" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & VAR_CTA & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + VAR_BS2
        rstdestino.Update
      End If
    End If
    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    'LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
    'Frmmensaje.Visible = False
    db.CommitTrans
  'End If
'  'marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

Private Function DESCAUX(VARAUX As String, VARCODIG As String)
    Set rsAuxDetalle = New ADODB.Recordset
    If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
    Select Case VARAUX
        Case "01"
            rsAuxDetalle.Open "SELECT beneficiario_denominacion AS DESAUX2 FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT beneficiario_denominacion AS DESAUX FROM gc_beneficiario where beneficiario_codigo = '" & VARCODIG & "' "
        Case "02"
            rsAuxDetalle.Open "SELECT cta_descripcion AS DESAUX2 FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT cta_descripcion AS DESAUX FROM fc_cuenta_bancaria where Cta_codigo = '" & VARCODIG & "' "
        Case "03"
            rsAuxDetalle.Open "SELECT pro_codigo_det_descripcion AS DESAUX2 FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT pro_codigo_det_descripcion AS DESAUX FROM fo_proyectos_ejecucion where pro_codigo_det = '" & VARCODIG & "' "
        Case "04"
            rsAuxDetalle.Open "SELECT unidad_descripcion AS DESAUX2 FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "05"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "06"
            rsAuxDetalle.Open "SELECT depto_descripcion AS DESAUX2 FROM gc_departamento where depto_codigo = '" & VARCODIG & "'  ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT depto_descripcion AS DESAUX FROM gc_departamento where depto_codigo = '" & VARCODIG & "' "
        Case "07"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "08"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "09"
            rsAuxDetalle.Open "SELECT Org_descripcion AS DESAUX2 FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' ", db, adOpenKeyset, adLockReadOnly
            'db.Execute "SELECT Org_descripcion AS DESAUX FROM fc_organismo_financiamiento where org_codigo = '" & VARCODIG & "' "
        Case "10"
            'db.Execute "SELECT impuesto_descripcion AS DESAUX FROM fc_impuestos where impuesto_codigo = '" & VARCODIG & "' "
        Case "11"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "12"
            DESAUX = ""
            'db.Execute "SELECT unidad_descripcion AS DESAUX FROM gc_unidad_ejecutora where unidad_codigo = '" & VARCODIG & "' "
        Case "00"
            DESAUX = ""
    End Select
    If rsAuxDetalle.RecordCount > 0 Then
      DESAUX = RTrim(rsAuxDetalle!DESAUX2)
    Else
      DESAUX = ""
    End If
End Function

'Private Sub f_actual_rec(org, codant)
'  Dim acumDl As Double
'  Dim rsrecalc As New ADODB.Recordset
'  Set rsrecalc = New ADODB.Recordset
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
'  If rsrecalc.RecordCount > 0 Then
'    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'  Else
'    acumDl = 0
'  End If
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsrecalc.RecordCount > 0 Then
'    rsrecalc!monto_recaudado_dolares = acumDl
'  End If
'  rsrecalc.Update
'  If rsrecalc.State = 1 Then rsrecalc.Close
'
'End Sub

Private Sub graba_proyecto()
'    Select Case Ado_datos.Recordset!unidad_codigo
'        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
'            VAR_PROY = 12
'        Case "UCOM"
'            VAR_PROY = 17
'        Case "DVTA"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & GlUsuario & "', '" & Date & "')"
'    End If
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
    
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False
      
      'If v_añadir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         'Call add_correl
         Set rstdestino = New ADODB.Recordset
         If VAR_TIPOV = "L" Then
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
         Else
            rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
         End If
         VAR_CODANT = 0
         If rstdestino.RecordCount > 0 Then
            VAR_CODANT = rstdestino!ingreso_codigo
            VAR_ORG = rstdestino!org_codigo
            VAR_FTE = rstdestino!fte_codigo
            VAR_TIPOS = rstdestino!solicitud_tipo
            VAR_PARTIDA = rstdestino!rubro_codigo
            
         Else
             Select Case VAR_COD4
                 Case "DVTA", "DCOMB", "DCOMC", "DCOMS"             'INI COMERCIAL
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     VAR_TIPOS = 3
                     EST_PROG = 18      'Activ=17, Proy=18
                     If VAR_TIPOV = "L" Then
                         VAR_PARTIDA = "11360"
                     Else
                         VAR_PARTIDA = "11310"
                     End If
                 Case "COMEX"            'INI COMEX
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     EST_PROG = 15      'Activ=14, Proy=15
                     VAR_PARTIDA = "11310"
                 Case "DNMAN", "DMANB", "DMANC", "DMANS"            'INI MANTENIMIENTO
                     VAR_ORG = "112"
                     VAR_FTE = "10"
                     VAR_TIPOS = 10
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11320"
                 Case "DNREP", "DNEME", "DREPB", "DREPC", "DREPS"            'INI REPARACIONES 'INI EMERGENCIAS
                     VAR_ORG = "113"
                     VAR_FTE = "10"
                     VAR_TIPOS = 7
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11330"
                 Case "DNMOD"            'INI MODERNIZACION
                     VAR_ORG = "114"
                     VAR_FTE = "10"
                     VAR_TIPOS = 9
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11340"
                 Case "DNINS", "DNAJS", "DINSB", "DINSC", "DINSS"            'INI INSTALACIONES    'INI AJUSTE
                     VAR_ORG = "111"
                     VAR_FTE = "10"
                     VAR_TIPOS = 4
                     EST_PROG = 18      'Activ=17, Proy=18
                     VAR_PARTIDA = "11350"
                 Case "DCONT"            'INI EMERGENCIAS
                     VAR_ORG = "112"
                     VAR_FTE = "10"
                     VAR_TIPOS = 10
                     EST_PROG = 12      'Activ=11, Proy=12
                     VAR_PARTIDA = "11320"
                 Case Else               'INI COMPRAS
                     VAR_ORG = "311"
                     VAR_FTE = "30"
                     VAR_TIPOS = 6
                     EST_PROG = 18      'Activ=17, Proy=18
                     VAR_PARTIDA = "11320"
            End Select
            'Call add_correl
            'EXEPCION PARA GRABAR CONTRATO EN INGRESOS
'             rstdestino.AddNew
'             rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
'             rstdestino("ingreso_codigo") = correlativo1
'             rstdestino("org_codigo") = VAR_ORG
'             If VAR_CODANT = 0 Then
'                VAR_CODANT = correlativo1
'             End If
'             rstdestino("ingreso_codigo_anterior") = VAR_CODANT
'
'             rstdestino("proceso_codigo") = "FIN"
'             rstdestino("subproceso_codigo") = "FIN-01"
'             rstdestino("etapa_codigo") = "FIN-01-02"
'             rstdestino("clasif_codigo") = "ADM"
'             rstdestino("doc_codigo") = "R-110"
'             rstdestino("doc_numero") = correlativo1
'             rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos16.Recordset("unidad_codigo")
'             rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos16.Recordset("solicitud_codigo")
'             '
'             rstdestino("solicitud_tipo") = VAR_TIPOS
'
'             If VAR_COD4 = "DVTA" Then
'                rstdestino("tipo_comp") = "DEY"
'                rstdestino("Codigo_tipo") = "DEY"
'             Else
'                rstdestino("tipo_comp") = "DEI"
'                rstdestino("Codigo_tipo") = "DEI"
'             End If
'             'OJO JQA
'             rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
'             rstdestino("fecha_ingreso") = Date
'             rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
'             rstdestino("tipo_moneda") = VAR_MONEDA
'             'VAR_MONEDA = "BOB"
'             rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA       'Ado_datos.Recordset("cobranza_observaciones")
'             'CAMBIAR FTE
'             rstdestino("fte_codigo") = VAR_FTE
'             'CAMBIAR RUBROS
'             rstdestino("rubro_codigo") = VAR_PARTIDA
'             'CAMBIAR RUBROS
'             rstdestino("cheque_o_trf") = "T"
'             'CAMBIAR CTA
'             rstdestino("cta_codigo") = VAR_CTA
'             If VAR_CTA = "NN" Then
'                rstdestino("Bco_codigo") = "BCP"
'             Else
'                rstdestino("Bco_codigo") = "BMS"
'             End If
'             'CAMBIAR CTA
'             rstdestino("numero_documento") = VAR_COD1
'             rstdestino("unidad_codigo_ant") = VAR_CITE
'             rstdestino("monto_dolares") = VAR_DOL2 * 12
'             rstdestino("monto_bolivianos") = VAR_BS2 * 12
'             rstdestino("monto_recaudado_dolares") = VAR_DOL2 * 12 'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
'             rstdestino("monto_recaudado_bolivianos") = VAR_BS2 * 12   'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
'             rstdestino("convenio_codigo") = "NN"
'             rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
'             rstdestino("estado_CODIGO") = "APR"
'             'rstdestino("estado_codigo_dr") = "DEI"
'
'             rstdestino("usr_CODIGO") = glusuario
'             rstdestino("fecha_registro") = Date
'             rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'             rstdestino.Update
'             VAR_CODANT = rstdestino!ingreso_codigo
'             VAR_ORG = rstdestino!org_codigo
'             VAR_FTE = rstdestino!fte_codigo
'             If rstdestino.State = 1 Then rstdestino.Close
'             If VAR_TIPOV = "L" Then
'                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEY' ", db, adOpenDynamic, adLockOptimistic
'             Else
'                rstdestino.Open "select * from fo_ingresos_cabecera where unidad_codigo= '" & VAR_COD4 & "' and solicitud_codigo= " & VAR_SOL & " and codigo_tipo= 'DEI' ", db, adOpenDynamic, adLockOptimistic
'             End If
         End If
         Call add_correl
         ' OJO CAMBIA FINANCIADOR WWWWWWWWWWWWWWWWWWWWW
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         rstdestino("org_codigo") = VAR_ORG
         If VAR_CODANT = 0 Then
            VAR_CODANT = correlativo1
         End If
         rstdestino("ingreso_codigo_anterior") = VAR_CODANT
         rstdestino("Codigo_tipo") = VAR_CODTIPO
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-02"
         rstdestino("etapa_codigo") = VAR_ETAPA
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = VAR_DOC
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4
         rstdestino("solicitud_codigo") = VAR_SOL
         rstdestino("solicitud_tipo") = VAR_TIPOS
         'OJO JQA
         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioMercado        'GlTipoCambioOficial
         rstdestino("tipo_moneda") = VAR_MONEDA
         'VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = VAR_TCOMP + ": " + VAR_GLOSA      'Ado_datos.Recordset("cobranza_observaciones")
         'VAR_GLOSA = "INGRESO POR: " + Ado_datos.Recordset("cobranza_observaciones")
         If VAR_TIPOV = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = VAR_CODTIPO
         End If
         rstdestino("fte_codigo") = VAR_FTE
         rstdestino("rubro_codigo") = VAR_PARTIDA
         rstdestino("cheque_o_trf") = "T"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = VAR_CTA
         If VAR_CTA = "2015046557-03-054" Then
            rstdestino("Bco_codigo") = "BCP"
         Else
            rstdestino("Bco_codigo") = "BMS"
         End If
         'CAMBIAR CTA
         NroFactura = Trim(Str(VAR_COD1))
         rstdestino("numero_documento") = NroFactura        'Ado_datos.Recordset!cobranza_nro_factura
         rstdestino("unidad_codigo_ant") = VAR_CITE
         rstdestino("monto_dolares") = VAR_DOL2
         rstdestino("monto_bolivianos") = VAR_BS2
         rstdestino("monto_recaudado_dolares") = VAR_DOL2   'Round(Ado_datos.Recordset("cobranza_total_dol"), 2)
         rstdestino("monto_recaudado_bolivianos") = VAR_BS2     'Round(Ado_datos.Recordset("cobranza_total_bs"), 2)
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = VAR_PROY2       'Ado_datos16.Recordset("edif_codigo")
         rstdestino("estado_CODIGO") = "APR"
         'rstdestino("estado_codigo_dr") = "DEI"

         rstdestino("usr_CODIGO") = glusuario
         rstdestino("fecha_registro") = Date
         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
         
         rstdestino.Update
         If rstdestino.State = 1 Then rstdestino.Close
        'db.CommitTrans
          
'          If rstIngresos.State = 1 Then rstIngresos.Close
'          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rstIngresos.Sort = "ingreso_codigo"
'          rstIngresos.Requery
          
'          rstIngresos.Requery
'          Set AdoIngresos.Recordset = rstIngresos
'          AdoIngresos.Refresh
'          AdoIngresos.Recordset.Find "ultimo = 'S'"
'          If Not (AdoIngresos.Recordset.EOF) Then
'            marca1 = AdoIngresos.Recordset.Bookmark
'            AdoIngresos.Recordset("ultimo") = "N"
'            AdoIngresos.Recordset.Update
'          End If

'          AdoIngresos.Recordset.Move marca1 - 1

'          marca1 = 0
      'End If
'   Else*
'      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
''      FraOpciones2.Visible = False
''      FraOpciones.Visible = True
''      FraIngresosNav.Enabled = True
''      FraIngresosDat.Enabled = False
''      AdoIngresos.Refresh
'   End If
'   LblAccion = ""
'AAQQQQQUIIIIIIIIII    JQA

End Sub

Private Sub add_correl()
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' ", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     VAR_ORG = "112"
     VAR_FTE = "10"
     If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
     rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "'  ", db, adOpenDynamic, adLockOptimistic
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG   'Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
'     rstcorrel_ing("fte_codigo") = "10"
'     'rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing("correlativo_ingreso") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo_ingreso")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     rstcorrel_ing("correlativo_ingreso") = rstcorrel_ing("correlativo_ingreso") + 1
     rstcorrel_ing.Update
     
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub

'Private Sub CmdGrabaCobranza()
'    If swnuevo = 1 Then
''      rstdestino.Open "select * from ao_ventas_detalle where correl_venta = " & lblcorrelVenta & " and venta_codigo = " & TxtNroVenta, db, adOpenKeyset, adLockOptimistic
''      Set Ado_datos16.Recordset = rstdestino
''      Ado_datos16.Recordset.AddNew
'      Ado_datos16.Recordset!correl_venta = Val(lblcorrelVenta.Caption)
'      Ado_datos16.Recordset!venta_codigo = Val(TxtNroVenta.Text)
'      Ado_datos16.Recordset!ges_gestion = Year(Date)    'Trim(LblGestion.Caption)
'    End If
'      Ado_datos16.Recordset!beneficiario_codigo = dtc_codigo2A.Text                                 'Codigo Beneficiario/Cliente
'      Ado_datos16.Recordset!ci = dtc_codigo4A.Text                                                     'Codigo Cobrador
'      Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text + " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
'      Ado_datos16.Recordset!deuda_cobrada = Val(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!deuda_cobrada_dol = Val(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
'      Ado_datos16.Recordset!fecha_cobranza = DTPFechaCobro.Value                                'Fecha de Cobranza
'      'Call acumulaMont(Ado_datos16.Recordset!ges_gestion, Ado_datos16.Recordset!correl_venta, Ado_datos16.Recordset!venta_codigo)
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
'
'      Ado_datos16.Recordset!obs_cobranza = TxtObs
'      Ado_datos16.Recordset!nro_cmpbte = Trim(TxtCmpbte)
'      Ado_datos16.Recordset!usr_usuario = GlUsuario
'      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      Ado_datos16.Recordset.Update
'End Sub

'Private Sub CmdModDetalle_Click()
'  FraDetalle.Visible = True
'  FraDetalle.Enabled = True
'  txtnosolicitud1.Enabled = False
'  txtcorrdet.Enabled = False
'  dtccodpar.SetFocus
'  CmdGraDetalle.Enabled = True
'  CmdAddDetalle.Enabled = False
'  CmdModDetalle.Enabled = False
'  CmdSalDetalle.Enabled = False
'  CmdCanDetalle.Enabled = True
'  swgrabar = 2
'End Sub

'Private Sub CmdGraDetalle_Click()
'    If swgrabar = 1 Then
'        Dim rstdestino As New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle_correl where formulario = '" & "F11" & "' and correl_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("correl_solicitud_detalle") = rstdestino("correl_solicitud_detalle") + 1
'        Else
'            rstdestino.AddNew
'            rstdestino("formulario") = "F11"
'            rstdestino("correl_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correl_solicitud_detalle") = 1
'        End If
'        correldetalle = rstdestino("correl_solicitud_detalle")
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correlativo_solicitud = " & Ado_datos.Recordset("codigo_solicitud"), db, adOpenDynamic, adLockOptimistic
'        rstdestino.AddNew
'        rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'        rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'        rstdestino("correlativo_detalle") = correldetalle
'        rstdestino("Par_codigo") = dtccodpar.Text
'        rstdestino("Importe_nacional") = txtsolpeso.Text
'        rstdestino("formulario") = "F11"
'        rstdestino.Update
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    If swgrabar = 2 Then
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoDetalleSolicitud.Recordset("ges_gestion") & "' and correlativo_solicitud = " & adoDetalleSolicitud.Recordset("correlativo_solicitud") & " and correlativo_detalle =" & adoDetalleSolicitud.Recordset("correlativo_detalle"), db, adOpenDynamic, adLockOptimistic
'        If Not (rstdestino.EOF) Then
'            rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'            rstdestino("correlativo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
'            rstdestino("correlativo_detalle") = correldetalle
'            rstdestino("Par_codigo") = dtccodpar.Text
'            rstdestino("Importe_nacional") = txtsolpeso.Text
'            rstdestino("formulario") = "F11"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
'        Set rs_datos14 = New ADODB.Recordset
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        rs_datos14.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(Ado_datos.Recordset("ges_gestion")) & "' and correlativo_solicitud = " & Trim(Ado_datos.Recordset("codigo_solicitud")) & " and formulario = 'F11'", db, ad0OpenKeyset, adLockOptimistic
'        Set adoDetalleSolicitud.Recordset = rs_datos14
'        adoDetalleSolicitud.Refresh
'    End If
'    CmdGraDetalle.Enabled = False
'    CmdAddDetalle.Enabled = True
'    CmdModDetalle.Enabled = True
'    CmdSalDetalle.Enabled = True
'    CmdCanDetalle.Enabled = False
'    FraDetalle.Enabled = False
'    swgrabar = 0
'End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdOKunidad_Click()
    swunidad = 1
        If swunidad = 1 Then
            Dim rstpagos As New ADODB.Recordset
            Set rstpagos = New ADODB.Recordset
            If rstpagos.State = 1 Then rstpagos.Close
            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
            rstpagos.AddNew
                rstpagos("ges_gestion") = Ado_datos.Recordset("ges_gestion")
                rstpagos("org_codigo") = DataCombo1.Text   'Ado_datos.Recordset("formulario")
                rstpagos("codigo_pago") = "" 'genera jorge
                rstpagos("codigo_solicitud") = Ado_datos.Recordset("codigo_solicitud")
                rstpagos("formulario") = Ado_datos.Recordset("formulario")
                rstpagos("codigo_unidad") = Ado_datos.Recordset("codigo_unidad")
                rstpagos("monto_bolivianos") = Ado_datos.Recordset("monto_bolivianos")
                rstpagos("estado_compromiso") = "N"
                rstpagos("justificacion") = Ado_datos.Recordset("justificacion_solicitud")
             rstpagos.Update
        End If
End Sub

Private Sub CmdGrabaCobro_Click()
End Sub

'Private Sub CmdGrabaDet_Click()
''If dtc_desc12 = "" Then
''    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'  If dtc_codigo15 = "" Then
'     MsgBox "Debe Elejir un Producto para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
''  If dtc_desc13 = "" Then
''    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
''    Exit Sub
''  End If
'    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
'    '    VAR_PARTIDA = "OK"
'    If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Then
'          'fraOpciones.Visible = True
'          'FraGrabarCancelar.Visible = False
'          'TxtNroVenta.Enabled = True
'          FrmEdita.Enabled = False
'        '  DtGListaN.Enabled = True
'          'cmdElige.Enabled = False
'        '  dtc_codigo15.Visible = False
'        '  dtc_desc15.Visible = False
'          'txt_descripcion_venta.Enabled = False
'        If swnuevo = 1 Then
'          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
'          ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset("venta_codigo")
'          ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
'        End If
'          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
'          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
'          ado_datos14.Recordset!bien_codigo = Trim(dtc_codigo15.Text)                       'Codigo Bien (Equipo, Producto, etc)
'          ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
'          ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
'          ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
'          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
'          ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
'          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
'          If TxtCantidad.Text = "" Then
'            TxtCantidad.Text = "1"
'          End If
'          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
'          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
'          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
'          If CDbl(TxtDescuento) > 0 Then
'            ado_datos14.Recordset!venta_descuento_bs = CDbl(TxtDescuento.Text)      'Dcto por producto CON DESCUENTO
'            ado_datos14.Recordset!venta_descuento_dol = Val(TxtDescuento) / GlTipoCambioMercado
'          Else
'            'ado_datos14.Recordset!descuento_venta = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) * (CDbl(Dtc_aux12)) 'Dcto por producto DE LA TABLA
'            TxtDescuento.Text = "0"
'            ado_datos14.Recordset!venta_descuento_bs = 0
'            ado_datos14.Recordset!venta_descuento_dol = 0
'          End If
'          ado_datos14.Recordset!venta_precio_total_bs = (Val(TxtCantidad) * CDbl(TxtPrecioU.Text)) - (CDbl(TxtDescuento)) 'Precio Total Producto
'          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
'          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
'          ado_datos14.Recordset!venta_precio_total_dol = (ado_datos14.Recordset!venta_precio_total_bs) / GlTipoCambioMercado
'          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
'          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
'          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
'          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
'          If OpMod1.Value = True Then
'            ado_datos14.Recordset!modelo_elegido = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido = "N"
'          End If
'          If OpMod2.Value = True Then
'            ado_datos14.Recordset!modelo_elegido_h = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_h = "N"
'          End If
'          If OpMod2.Value = True Then
'            ado_datos14.Recordset!modelo_elegido_x = "S"
'          Else
'            ado_datos14.Recordset!modelo_elegido_x = "N"
'          End If
'          ado_datos14.Recordset!estado_codigo = "REG"
'          ado_datos14.Recordset!usr_codigo = GlUsuario
'          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'          ado_datos14.Recordset.Update
'        'db.CommitTrans
'
'        'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
'        Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"))
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        FraNavega.Enabled = True
'        FrmDetalle.Enabled = True
'        'FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
'        FrmABMDet.Visible = True
'        FrmABMDet2.Visible = True
'        Call OptFilGral1_Click
'        If swnuevo = 1 Then
'          'Call abre_ventas_det
'          'rs_datos14.Requery
'          'ado_datos14.Refresh
'          'ado_datos14.Recordset.MoveLast
'
'        End If
'        swnuevo = 0
'    Else
'        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
'    End If
'  'Else
'  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
'  'End If
'End Sub

'Private Sub BtnImprimir2_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    Dim iResult As Variant  ', i%, y%
'    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R103_recibo_cobranza_grp.rpt"
'    CryR01.WindowShowRefreshBtn = True
''    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
''    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
''    CryR01.StoredProcParam(2) = Me.Ado_datos.Recordset!cobranza_codigo
'    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!venta_codigo
'    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!cobranza_codigo
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

Private Sub BtnVer_Click()
'CONTABILIZA FAC
VAR_JQ = "FAC"
Set rs_aux20 = New Recordset
If rs_aux20.State = 1 Then rs_aux20.Close
'queryinicial2 = "select * From av_ventas_cobranzas_REC where venta_codigo_new = '1' AND unidad_codigo ='DVTA' AND SOLICITUD_CODIGO = '975' "
queryinicial2 = "select * From av_ventas_cobranzas_REC WHERE unidad_codigo ='DNMAN' "
rs_aux20.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
rs_aux20.Sort = "cobranza_fecha_fac, cobranza_codigo"
If rs_aux20.RecordCount > 0 Then
rs_aux20.MoveFirst

While Not rs_aux20.EOF
    'Set rs_aux7 = New ADODB.Recordset
    'rs_aux7.CursorLocation = adUseClient
    'If rs_aux7.State = 1 Then rs_aux7.Close
    ''rs_aux7.Open "select * from cv_comprobante_y_diario where cobranza_pago = '" & rs_aux20!cobranza_codigo & "' AND Tipo_Comp = 'REF' ", db, adOpenDynamic, adLockOptimistic
    'rs_aux7.Open "Select * from av_ventas_cobranza where cobranza_codigo = " & rs_aux20!cobranza_codigo & " ", db, adOpenStatic
    'If rs_aux7.RecordCount > 0 Then
    ''If (Ado_datos.Recordset!factura_impresa = "N") And (Ado_datos.Recordset!cobranza_deuda_bs <> "0.00") Then
    '  If rs_aux20!doc_codigo_fac = "R-101" Then
        '===== ini GENERA CONTABILIZA FACTURADO ====
            correlv = rs_aux20("venta_codigo")
            nroventa = rs_aux20("venta_codigo")
            NRO_COBR = rs_aux20!cobranza_codigo
            VAR_BENEF = rs_aux20!beneficiario_codigo_fac
            VAR_CITE = rs_aux20!unidad_codigo_ant
            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
            VAR_GLOSA = Trim(rs_aux20!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
            NRO_COBR = rs_aux20!cobranza_codigo
            VAR_DOL2 = Round(rs_aux20!cobranza_total_dol, 2)
            VAR_BS2 = Round(rs_aux20!cobranza_total_bs, 2)
            VAR_13 = Round(rs_aux20!cobranza_descuento_bs, 2)
            VAR_87 = Round(rs_aux20!cobranza_descuento_dol, 2)
            
            VAR_CTA = IIf(rs_aux20!cta_codigo = "", "NN", rs_aux20!cta_codigo)
            var_literal = IIf(IsNull(rs_aux20!Literal), "-", rs_aux20!Literal)
            VAR_FFAC = rs_aux20!cobranza_fecha_fac
            VAR_VTIPO = rs_aux20!venta_tipo
            VAR_ANIO = Year(VAR_FFAC)
            gestion0 = Year(VAR_FFAC)
            VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
            VAR_CODTIPO = "REF"     'Tipo Comprobante (paralelo VAR_DOC)
            VAR_DOC = "R-112"       'Doc. Respaldo
            VAR_ETAPA = "FIN-02-02"
            NitCi = IIf(IsNull(rs_aux20!beneficiario_codigo_fac), "0", rs_aux20!beneficiario_codigo_fac)    'VAR_BENEF
            'Llave = Trim(rs_aux1!dosifica_llave)
            Autorizacion = rs_aux20!cobranza_nro_autorizacion
            VAR_PROY2 = rs_aux20!edif_codigo
            VAR_COD4 = rs_aux20!unidad_codigo
            VAR_TIPOV = rs_aux20!venta_tipo
            VAR_SOL = rs_aux20!solicitud_codigo
            VAR_MONEDA = rs_aux20!tipo_moneda
            VAR_COD1 = rs_aux20!cobranza_nro_factura
            VAR_TCOMP = "FAC: " + Trim(VAR_COD1)
'            sino = MsgBox("Esta seguro(a) de Re-Contabilizar la Factura Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
'            If sino = vbYes Then
                VAR_COD2 = rs_aux20!cobranza_nro_autorizacion
                NroFactura = Trim(Str(VAR_COD1))
                Fecha = Val(Format((VAR_FFAC), "YYYYMMDD"))
                Monto = Redondeo((VAR_BS2), 0)
                'If VAR_TIPOV = "C" Then
                    VAR_SW = 0
                    Call Contabiliza_venta
                'End If
'            Else
'                VAR_COD1 = "0"
'                If rs_aux1.State = 1 Then rs_aux1.Close
'                Exit Sub
'            End If
        '===== fin TERMINA CONTABILIZA FACTURADO =====
    '  End If
    'End If
    rs_aux20.MoveNext
Wend
End If
MsgBox "OK"
                
'
'
'        TxtCmpbte = VAR_COD1
'        If (rs_aux20("estado_codigo_sol") = "APR" And rs_aux20("estado_codigo_fac") = "REG") Then          'REG
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'      Else
'        'Call generarRepRecibo
'      End If
'      If rs_aux20!doc_codigo_fac = "R-103" Then
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'        '===== ini GENERA EL CODIGO DE RECIBO ====
''        Set rs_aux1 = New ADODB.Recordset
''        rs_aux1.CursorLocation = adUseClient
''        If rs_aux1.State = 1 Then rs_aux1.Close
''        rs_aux1.Open "select * from gc_documentos_respaldo where doc_codigo = 'R-103' AND estado_codigo = 'APR' ", db, adOpenDynamic, adLockOptimistic
''        If rs_aux1.RecordCount > 0 Then
''            gestion0 = glGestion        'Ado_datos.Recordset("ges_gestion")
''            correlv = rs_aux20("venta_codigo")
''            nroventa = rs_aux20("venta_codigo")
''            NRO_COBR = Me.rs_aux20!cobranza_codigo
''            VAR_BENEF = rs_aux20!beneficiario_codigo
''            VAR_CITE = rs_aux20!unidad_codigo_ant
''            'VAR_GLOSA = Ado_datos.Recordset!cobranza_observaciones
''            VAR_GLOSA = Trim(rs_aux20!cobranza_observaciones) + " - Tram.: " + Trim(VAR_CITE)
''            VAR_DOL2 = Round(rs_aux20!cobranza_deuda_dol, 2)
''            VAR_BS2 = Round(rs_aux20!cobranza_deuda_bs, 2)
''            'VAR_CTA = IIf(rs_aux20!Cta_Codigo = "", "NN", rs_aux20!Cta_Codigo)
''            var_literal = rs_aux20!Literal
''            'Llave = Trim(rs_aux1!dosifica_llave)
''            NitCi = IIf(dtc_aux5.Text = "", rs_aux20!beneficiario_codigo_fac, dtc_aux5.Text)    'VAR_BENEF
''            'Autorizacion = rs_aux1!dosifica_autorizacion
''            VAR_PROY2 = rs_aux20!edif_codigo
''            VAR_COD4 = rs_aux20!unidad_codigo
''            VAR_TIPOV = rs_aux20!venta_tipo
''            VAR_SOL = rs_aux20!solicitud_codigo
''            VAR_MONEDA = Ado_datos.Recordset!tipo_moneda
''
''            VAR_COD1 = CDbl(rs_aux1!correl_doc) + 1
''            sino = MsgBox("Esta seguro(a) de IMPRIMIR la Recibo Nro. " + Str(VAR_COD1) + " ?", vbYesNo, "Confirmando")
''            If sino = vbYes Then
''                rs_aux1!correl_doc = Trim(Str(VAR_COD1))
''                rs_aux1.Update
''                'GENERA CORREL NOTA DEBITO POR DEPTO INI
''                VAR_GLOSA = TxtObs.Text
''                db.Execute "update ao_ventas_cobranza set cobranza_observaciones = '" & VAR_GLOSA & "' Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
'''                'GENERA CORREL NOTA DEBITO POR DEPTO FIN
''
''                VAR_COD2 = "0"  'rs_aux1!dosifica_autorizacion
''                NroFactura = Trim(Str(VAR_COD1))
''                '===== ini nombre archivo de la FACTURA ====
''                db.Execute "update ao_ventas_cobranza set archivo_foto = 'R103-' + '" & Str(VAR_COD1) & "' + '.JPG' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                db.Execute "update ao_ventas_cobranza set archivo_foto_cargado = 'N' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                '===== fin nombre archivo de la FACTURA ====
''                ' ACTUALIZA NRO FAC. EN ao_ventas_cobranza
''                'db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac = '" & Date & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                'db.Execute "update ao_ventas_cobranza set cobranza_nro_factura = " & VAR_COD1 & " Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                'db.Execute "update ao_ventas_cobranza set cobranza_nro_autorizacion = " & VAR_COD2 & " Where ao_ventas_cobranza.venta_codigo = " & nroventa & "  And ao_ventas_cobranza.cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                'IMPRIMIR FACTURA
''                Fecha = Val(Format((Date), "YYYYMMDD"))
''                Monto = Redondeo((VAR_BS2), 0)
''                db.Execute "update ao_ventas_cobranza set cobranza_fecha_fac2 = '" & Fecha & "' Where venta_codigo = " & nroventa & "  And cobranza_codigo = " & NRO_COBR & " "
''                'Dim F1
''                'FI = rs_aux20!cobranza_fecha_cobro
''                'frm_qr.txt_texto = GlParametro + "|" + GlParametroDes + "|" + Trim(str(VAR_COD1)) + "|" + TrimSTR((VAR_COD2)) + "|" + '" & rs_aux20!cobranza_fecha_cobro & "' + "|" + " & rs_aux20!cobranza_deuda_bs & " + "|" + '" & rs_aux1!dosifica_codigo_control & "' + "|" + '" & rs_aux1!dosifica_fecha_limite & "' + "|" + "0" + "|" + "0" + "|" + '" & rs_aux20!beneficiario_codigo & "' + "|" + '" & dtc_desc2A.Text & "'
''                'frm_qr.Show vbModal
''                'NIT del emisor, Nombre o Razón Social del emisor, Número correlativo de Factura, Número de Autorización, Fecha de emisión, Importe de la compra, Código de Control, Fecha Límite de Emisión, 0, 0, NIT / NDI Comprador, Nombre o Razón Social del comprador
''
''                'MsgBox "Se está Imprimiendo la Factura Nro. " + Str(VAR_COD1), , "Atención"
''                db.Execute "update ao_ventas_cobranza set factura_impresa = 'S' Where venta_codigo = " & rs_aux20!venta_codigo & "  And cobranza_codigo = " & rs_aux20!cobranza_codigo & " "
''                db.Execute "update ao_ventas_cobranza set estado_codigo_fac = 'APR' Where cobranza_codigo = " & rs_aux20("cobranza_codigo") & " "
''                db.Execute "update ao_ventas_cobranza set estado_codigo_bco = 'REG' Where cobranza_codigo = " & rs_aux20("cobranza_codigo") & " "
''
''                VAR_SW = 1
''                'CodigoContro = CodigoControl(Autorizacion, NroFactura, NitCi, Fecha, Monto, Llave)
''                'db.Execute "update ao_ventas_cobranza set cobranza_codigo_control = '" & CodigoContro & "' Where cobranza_codigo = " & rs_aux20("cobranza_codigo") & " "
''                Call IMPRIME_RECIBO
''                'If VAR_TIPOV = "C" Then
''                    Call Contabiliza_venta
''                'End If
''            Else
''                VAR_COD1 = "0"
''                If rs_aux1.State = 1 Then rs_aux1.Close
''                Exit Sub
''            End If
''        End If
''        If rs_aux1.State = 1 Then rs_aux1.Close
''        '===== fin TERMINA GENERACION DE FACTURA =====
''        TxtCmpbte = VAR_COD1
''        If (rs_aux20("estado_codigo_sol") = "APR" And rs_aux20("estado_codigo_fac") = "REG") Then          'REG
''          Call OptFilGral1_Click
''        Else
''          Call OptFilGral2_Click
''        End If
'      'WWWWWWWWWWWWWWWWWWWWWWWWW
'      End If
'    Else
'        MsgBox "Error: La Factura Nro. " + rs_aux20!cobranza_nro_factura + " ya fue Impresa y contabilizada. Elija otro Registro a procesar...", , "Atención"
'        'Call IMPRIME_FACTURA
'        'If (rs_aux20("estado_codigo_sol") = "APR" And rs_aux20("estado_codigo_fac") = "REG") Then          'REG
'        '  Call OptFilGral1_Click
'        'Else
'        '  Call OptFilGral2_Click
'        'End If
'    End If
'  Else
'    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'  End If

End Sub

Private Sub cmd_benef_Click()
    Set rs_datos8 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_beneficiario where tipoben_codigo <> '0' and tipoben_codigo <> '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    If Ado_datos8.Recordset.RecordCount > 0 Then
        dtc_desc8.BoundText = dtc_codigo8.BoundText
        FraGrabarCancelar.Enabled = False
        frm_benef.Visible = True
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
  'IMPRIME ENVIAR A FACTURAR
  If Ado_datos.Recordset.RecordCount > 0 And (Ado_datos.Recordset!estado_codigo1 = "APR") Then         'And (dtc_aux5.Text <> "")
        gestion0 = Ado_datos.Recordset!ges_gestion
        nroventa = Ado_datos.Recordset!venta_codigo
        NRO_COBR = Ado_datos.Recordset!cobranza_codigo
        VAR_COD1 = Ado_datos.Recordset!cobranza_nro_factura
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
    
        Dim iResult As Variant  ', i%, y%
'        sino = MsgBox("Imprimirá con el detalle de Bienes ? ", vbYesNo, "Confirmando")
'
'        If sino = vbYes Then
'            If VAR_COD4 = "DNMAN" Then
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_man.rpt"
'            Else
'                CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura_rep.rpt"
'            End If
'        Else
'            CryQ01.ReportFileName = App.Path & "\reportes\ventas\ar_R101_factura.rpt"
'        End If
        CryQ01.ReportFileName = App.Path & "\reportes\ventas\fr_envia_facturar.rpt"
        CryQ01.WindowShowRefreshBtn = True
        'CryQ01.StoredProcParam(0) = gestion0       'Me.Ado_datos.Recordset!ges_gestion
        'CryQ01.StoredProcParam(1) = nroventa        'Me.Ado_datos.Recordset!venta_codigo
        'CryQ01.StoredProcParam(2) = NRO_COBR        'Me.Ado_datos.Recordset!cobranza_codigo
        var_literal = Ado_datos.Recordset!Literal
        CryQ01.Formulas(1) = "literalcobro = '" & var_literal & "' "
        CryQ01.Formulas(2) = "correlcobro = '" & NRO_COBR & "' "
        ''" & Ado_datos.Recordset!cobranza_codigo & "' "
        '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
        iResult = CryQ01.PrintReport
        If iResult <> 0 Then MsgBox CryQ01.LastErrorNumber & " : " & CryQ01.LastErrorString, vbCritical, "Error de impresión"
  Else
      MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
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

Private Sub Command1_Click()
'CONTABILIZA REC
VAR_JQ = "REC"
'Set rs_aux20 = New Recordset
'If rs_aux20.State = 1 Then rs_aux20.Close
''queryinicial2 = "select * From av_ventas_cobranzas_REC where venta_codigo_new = '1' AND unidad_codigo ='DVTA' AND SOLICITUD_CODIGO = '975' "
'queryinicial2 = "select * From av_ventas_cobranzas_REC WHERE unidad_codigo ='DREPB' "
'rs_aux20.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'rs_aux20.Sort = "cobranza_fecha_fac, cobranza_codigo"
'If rs_aux20.RecordCount > 0 Then
'rs_aux20.MoveFirst
'While Not rs_aux20.EOF
'
'    rs_aux20.MoveNext
'Wend
'End If
 Set rs_datos1 = New ADODB.Recordset
 If rs_datos1.State = 1 Then rs_datos1.Close
 rs_datos1.Open "Select * from ao_ventas_cobranza_det order by cobranza_fecha, cobranza_codigo", db, adOpenStatic
 Set AdoAux.Recordset = rs_datos1
 If AdoAux.Recordset.RecordCount > 0 Then
   rs_datos1.MoveFirst
   While Not rs_datos1.EOF
     If (AdoAux.Recordset("cobranza_codigo") = 0) Then
        MsgBox "No se puede CONTABILIZAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        'Exit Sub
     Else
        If AdoAux.Recordset("estado_codigo") <> "ANL" Then
           'sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
           'If sino = vbYes Then
               'If AdoAux.Recordset("venta_tipo") = "C" Or AdoAux.Recordset("venta_tipo") = "V" Then
               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               'End If
               Set rs_datos2 = New ADODB.Recordset
               If rs_datos2.State = 1 Then rs_datos2.Close
               rs_datos2.Open "Select * from av_ventas_cobranza where cobranza_codigo = " & AdoAux.Recordset!cobranza_codigo & " ", db, adOpenStatic
               'Set AdoAux.Recordset = rs_datos1
                
               correlv = rs_datos2!venta_codigo
               nroventa = rs_datos2!venta_codigo
               
               VAR_BENEF = rs_datos2!beneficiario_codigo
               VAR_CITE = rs_datos2!unidad_codigo_ant
               VAR_GLOSA = Trim(AdoAux.Recordset!cobranza_observaciones) '+ " - Nro.: " + Trim(VAR_CITE)
               VAR_DOL2 = Round(AdoAux.Recordset!cobranza_dol, 2)
               VAR_BS2 = Round(AdoAux.Recordset!cobranza_bs, 2)
               VAR_CTA = IIf(AdoAux.Recordset!cta_codigo = "", "NN", AdoAux.Recordset!cta_codigo)
               VAR_PROY2 = rs_datos2!edif_codigo
               VAR_COD4 = rs_datos2!unidad_codigo
               VAR_TIPOV = ""       'rs_datos2!venta_tipo
               VAR_SOL = rs_datos2!solicitud_codigo
               'If AdoAux.Recordset!cta_codigo <> "NN" Then
               VAR_FFAC = IIf(IsNull(AdoAux.Recordset!cobranza_fecha), Date, AdoAux.Recordset!cobranza_fecha)   '"17/11/2016"  '
               'End If
               NroFactura = rs_datos2!cobranza_nro_factura
               NRO_COBR = Me.AdoAux.Recordset!cobranza_codigo
               var_literal = "-"        'IIf(IsNull(AdoAux.Recordset!Literal), "-", AdoAux.Recordset!Literal)
               VAR_MONEDA = "BOB"   'AdoAux.Recordset!tipo_moneda
               VAR_CODTIPO = "REC"
               VAR_DOC = "R-110"
               VAR_ETAPA = "FIN-02-03"
               VAR_TCOMP = "REC"        '"RECAUDADO (INGRESOS)"
               VAR_ANIO = Year(VAR_FFAC)
               gestion0 = Year(VAR_FFAC)
               VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
               
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "'  "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                    'VAR_DOCR = rs_aux2!correl_doc
'                    rs_aux2.Update
'                End If
                
                ' GRABA Nombre de Archivo en ao_ventas_cabecera
               VAR_SW = 2
               Call Contabiliza_venta
               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
               'Call OptFilGral1_Click
           'End If
        End If
     End If
     rs_datos1.MoveNext
   Wend
   MsgBox "OK"
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

Private Sub Command2_Click()
'CONTABILIZA DEI-DEY
 VAR_JQ = "DEI"
 Set rs_datos1 = New ADODB.Recordset
 If rs_datos1.State = 1 Then rs_datos1.Close
 rs_datos1.Open "Select * from av_ventas_cabecera_apr WHERE unidad_codigo = 'DVTA' AND solicitud_codigo = '951' order by venta_fecha, venta_codigo", db, adOpenStatic
 Set AdoAux.Recordset = rs_datos1
 If AdoAux.Recordset.RecordCount > 0 Then
   rs_datos1.MoveFirst
   While Not rs_datos1.EOF
     If (AdoAux.Recordset("venta_codigo") = 0) Then
        MsgBox "No se puede CONTABILIZAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        'Exit Sub
     Else
        If AdoAux.Recordset("estado_codigo") <> "ANL" Then
           'sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
           'If sino = vbYes Then
               ''If AdoAux.Recordset("venta_tipo") = "C" Or AdoAux.Recordset("venta_tipo") = "V" Then
               ''     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               ''End If
               'Set rs_datos2 = New ADODB.Recordset
               'If rs_datos2.State = 1 Then rs_datos2.Close
               'rs_datos2.Open "Select * from av_ventas_cobranza where cobranza_codigo = " & AdoAux.Recordset!cobranza_codigo & " ", db, adOpenStatic
               ''Set AdoAux.Recordset = rs_datos1
               correlv = rs_datos1!venta_codigo
               nroventa = rs_datos1!venta_codigo
               
               VAR_BENEF = rs_datos1!beneficiario_codigo
               VAR_CITE = rs_datos1!unidad_codigo_ant
               VAR_GLOSA = Trim(AdoAux.Recordset!venta_descripcion) '+ " - Nro.: " + Trim(VAR_CITE)
               VAR_DOL2 = Round(AdoAux.Recordset!venta_monto_total_dol, 2)
               VAR_BS2 = Round(AdoAux.Recordset!venta_monto_total_bs, 2)
               VAR_CTA = "NN"       'IIf(AdoAux.Recordset!Cta_Codigo = "", "NN", AdoAux.Recordset!Cta_Codigo)
               VAR_PROY2 = rs_datos1!edif_codigo
               VAR_COD4 = rs_datos1!unidad_codigo
               VAR_TIPOV = rs_datos1!venta_tipo
               VAR_SOL = rs_datos1!solicitud_codigo
               'If AdoAux.Recordset!cta_codigo <> "NN" Then
               VAR_FFAC = AdoAux.Recordset!venta_fecha   '"17/11/2016"  '
               'End If
               NroFactura = "0"      'rs_datos1!cobranza_nro_factura
               NRO_COBR = "0"        'Me.AdoAux.Recordset!cobranza_codigo
               var_literal = Literal(CStr(VAR_BS2)) + " BOLIVIANOS"     'IIf(IsNull(AdoAux.Recordset!Literal), "-", AdoAux.Recordset!Literal)
               VAR_MONEDA = "BOB"   'AdoAux.Recordset!tipo_moneda
               'If VAR_TIPOV = "L" Then
               '     VAR_CODTIPO = "DEY"
               'Else
                    VAR_CODTIPO = "DEI"
               'End If
               VAR_DOC = "R-112"
               VAR_ETAPA = "FIN-01-01"
               VAR_TCOMP = VAR_CODTIPO      '"DE"        '"RECAUDADO (INGRESOS)"
               VAR_ANIO = Year(VAR_FFAC)
               gestion0 = Year(VAR_FFAC)
               VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
               
'                Set rs_aux2 = New ADODB.Recordset
'                If rs_aux2.State = 1 Then rs_aux2.Close
'                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "'  "
'                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux2.RecordCount > 0 Then
'                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                    'VAR_DOCR = rs_aux2!correl_doc
'                    rs_aux2.Update
'                End If
                
                ' GRABA Nombre de Archivo en ao_ventas_cabecera
               VAR_SW = 1
               Call Contabiliza_venta
               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
               'Call OptFilGral1_Click
           'End If
        End If
     End If
     rs_datos1.MoveNext
   Wend
   MsgBox "OK"
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

Private Sub Command3_Click()
'RECODIFICA
VAR_JQ = "REC"
'Set rs_aux20 = New Recordset
'If rs_aux20.State = 1 Then rs_aux20.Close
''queryinicial2 = "select * From av_ventas_cobranzas_REC where venta_codigo_new = '1' AND unidad_codigo ='DVTA' AND SOLICITUD_CODIGO = '975' "
'queryinicial2 = "select * From av_ventas_cobranzas_REC WHERE unidad_codigo ='DREPB' "
'rs_aux20.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
'rs_aux20.Sort = "cobranza_fecha_fac, cobranza_codigo"
'If rs_aux20.RecordCount > 0 Then
'rs_aux20.MoveFirst
'While Not rs_aux20.EOF
'
'    rs_aux20.MoveNext
'Wend
'End If
 Set rs_datos1 = New ADODB.Recordset
 If rs_datos1.State = 1 Then rs_datos1.Close
 rs_datos1.Open "Select * from cv_comprobante_gestion_mes order by ges_gestion, mes", db, adOpenStatic
 Set AdoAux.Recordset = rs_datos1
 If AdoAux.Recordset.RecordCount > 0 Then
   rs_datos1.MoveFirst
   db.Execute "DELETE co_correl_mes"
   correlv = 0
   While Not rs_datos1.EOF
     If (AdoAux.Recordset("Cod_Comp") = 0) Then
        MsgBox "No se puede CONTABILIZAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        'Exit Sub
     Else
        'loop por Gestion y mes
        VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
        If Trim(mes_trasaccion) = Trim(VAR_MES) Then
           'sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
           'If sino = vbYes Then
               'If AdoAux.Recordset("venta_tipo") = "C" Or AdoAux.Recordset("venta_tipo") = "V" Then
               '     db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               'End If
                Set rs_datos2 = New ADODB.Recordset
                If rs_datos2.State = 1 Then rs_datos2.Close
                rs_datos2.Open "Select * from av_ventas_cobranza where cobranza_codigo = " & AdoAux.Recordset!cobranza_codigo & " ", db, adOpenStatic
                'Set AdoAux.Recordset = rs_datos1
                
               correlv = rs_datos2!venta_codigo
               nroventa = rs_datos2!venta_codigo
               
               VAR_BENEF = rs_datos2!beneficiario_codigo
               VAR_CITE = rs_datos2!unidad_codigo_ant
               VAR_GLOSA = Trim(AdoAux.Recordset!cobranza_observaciones) '+ " - Nro.: " + Trim(VAR_CITE)
               VAR_DOL2 = Round(AdoAux.Recordset!cobranza_dol, 2)
               VAR_BS2 = Round(AdoAux.Recordset!cobranza_bs, 2)
               VAR_CTA = IIf(AdoAux.Recordset!cta_codigo = "", "NN", AdoAux.Recordset!cta_codigo)
               VAR_PROY2 = rs_datos2!edif_codigo
               VAR_COD4 = rs_datos2!unidad_codigo
               VAR_TIPOV = ""       'rs_datos2!venta_tipo
               VAR_SOL = rs_datos2!solicitud_codigo
               'If AdoAux.Recordset!cta_codigo <> "NN" Then
               VAR_FFAC = IIf(IsNull(AdoAux.Recordset!cobranza_fecha), Date, AdoAux.Recordset!cobranza_fecha)   '"17/11/2016"  '
               'End If
               NroFactura = rs_datos2!cobranza_nro_factura
               NRO_COBR = Me.AdoAux.Recordset!cobranza_codigo
               var_literal = "-"        'IIf(IsNull(AdoAux.Recordset!Literal), "-", AdoAux.Recordset!Literal)
               VAR_MONEDA = "BOB"   'AdoAux.Recordset!tipo_moneda
               VAR_CODTIPO = "REC"
               VAR_DOC = "R-110"
               VAR_ETAPA = "FIN-02-03"
               VAR_TCOMP = "REC"        '"RECAUDADO (INGRESOS)"
               VAR_ANIO = Year(VAR_FFAC)
               gestion0 = Year(VAR_FFAC)
               VAR_MES = UCase(MonthName(Month(VAR_FFAC)))
               
               
'INSERT INTO co_correl_mes (doc_numero, Cod_Comp)
'(SELECT ROW_NUMBER() OVER (ORDER BY Fecha_transacion) AS doc_numero, Cod_Comp FROM co_comprobante_m
'WHERE (Fecha_transacion >= CONVERT(DATETIME, '2016-01-01 00:00:00', 102)) AND (Fecha_transacion <= CONVERT(DATETIME, '2016-01-31 00:00:00', 102)) AND
'(doc_codigo = 'R-110') AND (estado_codigo = 'APR') )
'
'update co_comprobante_m set co_comprobante_m.doc_numero = co_correl_mes.doc_numero
'from co_comprobante_m inner join co_correl_mes
'on co_comprobante_m.Cod_Comp = co_correl_mes.Cod_Comp
'
                ' GRABA Nombre de Archivo en ao_ventas_cabecera
               VAR_SW = 2
               Call Contabiliza_venta
               'db.Execute "update ao_ventas_cobranza set estado_codigo = 'APR' Where cobranza_codigo = " & NRO_COBR & " "
               'Call OptFilGral1_Click
           'End If
           
        Else
            db.Execute "DELETE co_correl_mes"
            correlv = 0
        End If
     End If
     rs_datos1.MoveNext
   Wend
   MsgBox "OK"
 Else
    MsgBox "NO existen registros para procesar !! ", vbExclamation, "Atención!"
  End If

End Sub

Private Sub DataCombo3_Click(Area As Integer)
'    dtc_cta2.BoundText
End Sub

Private Sub DataCombo8_Click(Area As Integer)
    DataCombo9.BoundText = DataCombo8.BoundText
End Sub

Private Sub DataCombo9_Click(Area As Integer)
    DataCombo8.BoundText = DataCombo9.BoundText
End Sub

Private Sub dtc_aux8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_aux8.BoundText
    dtc_codigo8.BoundText = dtc_aux8.BoundText
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

'Private Sub dtc_codigo61_Click(Area As Integer)
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
'End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    dtc_aux8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_cta_Click(Area As Integer)
    dtc_ctades.BoundText = dtc_cta.BoundText
End Sub

Private Sub dtc_cta2_LostFocus()
    If Ado_datos02.Recordset!cobranza_deuda_bs = "0" And dtc_cta2.Text <> "NN" Then
        TxtMonto02.Text = Ado_datos02.Recordset!cobranza_total_bs
        TxtMonto02D.Text = Ado_datos02.Recordset!cobranza_total_dol
    Else
        TxtMonto02.Text = "0"
        TxtMonto02D.Text = "0"
    End If
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

'Private Sub dtc_desc61_Click(Area As Integer)
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
'End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
    dtc_aux8.BoundText = dtc_desc8.BoundText
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
'    If (CDate(DTPFechaCobro2.Value) > CDate(DTPFechaCobro02.Value)) Then
'        MsgBox "La <<Fecha Cobranza2>> No puede ser MENOR a la <<Fecha Cobranza1>>, Vuelva a Intentar !! ", vbExclamation, "Atención!"
'        DTPFechaCobro02.SetFocus
'    End If
    VAR_DDIF = DateDiff("y", DTPFechaProg2, DTPFechaCobro02)
    If Val(VAR_DDIF) < 0 Then
        MsgBox "La <<Fecha de Cobro2>> NO puede ser MENOR a la Fecha de Facturación, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        DTPFechaCobro2A.SetFocus
    End If
    VAR_DDIF = DateDiff("y", DTPFechaCobro2, DTPFechaCobro02)
    If Val(VAR_DDIF) < 0 Then
        MsgBox "La <<Fecha de Cobro2>> NO puede ser MENOR a la <<Fecha de Cobro1>>, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        DTPFechaCobro2A.SetFocus
    End If
End Sub

Private Sub DTPFechaCobro2A_LostFocus()
    DTPFechaCobro2.Value = DTPFechaCobro2A.Value
'    VAR_DDIF = DateDiff("y", DTPFechaProg2, DTPFechaCobro2A)
'    If Val(VAR_DDIF) < 0 Then
'        MsgBox "La <<Fecha de Cobro1>> NO puede ser MENOR a la Fecha de Facturación, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
'        DTPFechaCobro2A.SetFocus
'    End If
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = 0
    parametro = Aux
    VAR_JQ = ""
    'BtnVer
    Call ABRIR_TABLAS_AUX
    'Call OptFilGral01_Click
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
'    FrmCabecera.Enabled = False
    FrmCobros.Enabled = False
    FrmCobros1.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    FraGrabarCancelar1.Visible = False
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = True
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
    Else
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
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
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "MCOLLAO" Or glusuario = "RTORREZ" Then
            queryinicial = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'REG' AND (unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC')) "
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND  estado_codigo_fac = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "') "
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
        queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  "
        'queryinicial = "SELECT ao_ventas_cobranza.*, ao_ventas_cabecera.* FROM ao_ventas_cobranza INNER JOIN ao_ventas_cabecera ON ao_ventas_cobranza.venta_codigo = ao_ventas_cabecera.venta_codigo"
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "MCOLLAO" Or glusuario = "RTORREZ" Then
            queryinicial = "select * From av_ventas_cobranza WHERE ((unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC') and estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG') "
        Else
            queryinicial = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
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
    If glusuario = "ADMIN" Or glusuario = "RVALDIVIEZO" Or glusuario = "VPAREDES" Or glusuario = "CNUÑEZ" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        queryinicial2 = "select * From av_ventas_cobranza where estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG'  "
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
            'queryinicial2 = "select * From av_ventas_cobranza WHERE unidad_codigo = 'DVTA' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND (unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC')  "
        Else
            'queryinicial2 = "select * From av_ventas_cobranza WHERE beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
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
    If glusuario = "RVALDIVIEZO" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR'  "
    Else
        If glusuario = "ADMIN" Then
            queryinicial2 = "select * From av_ventas_cobranza estado_codigo_fac = 'APR' AND estado_codigo_bco = 'REG'  "
        Else
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_fac = 'APR' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
    End If
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
    If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CNUÑEZ" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' "
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC') "
        Else
            queryinicial2 = "select * From av_ventas_cobranza WHERE estado_codigo_bco = 'APR' AND estado_codigo = 'REG' AND beneficiario_codigo_resp = '" & rs_datos9!beneficiario_codigo & "' "
        End If
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
        If glusuario = "RVALDIVIEZO" Or glusuario = "VPAREDES" Or glusuario = "CNUÑEZ" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
            queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and estado_codigo1 = 'REG' and doc_codigo_fac <> 'R-103' "      'ORDER BY cobranza_fecha_prog
        Else
            If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Then
                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Then
                    queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and estado_codigo1 = 'REG' "      'ORDER BY cobranza_fecha_prog
                End If
            End If
        End If
        
        rs_datos.Open queryinicial1, db, adOpenKeyset, adLockOptimistic
        rs_datos.Sort = "cobranza_fecha_sol"
        Set Ado_datos.Recordset = rs_datos.DataSource
        Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If glusuario = "RVALDIVIEZO" Or glusuario = "FACTURACION" Or glusuario = "APALACIOS" Then
        queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR'  and estado_codigo1 = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo_bco = 'REG' and doc_codigo_fac <> 'R-103' "     'ORDER BY cobranza_fecha_prog
    Else
        If glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "MCOLLAO" Then
                queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' and estado_codigo1 = 'APR'   and doc_codigo_fac = 'R-103' "      'ORDER BY cobranza_fecha_prog
            Else
                If glusuario = "ADMIN" Or glusuario = "VPAREDES" Then
                    queryinicial1 = "select * From av_ventas_cobranza WHERE estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG'  and estado_codigo1 = 'APR'  "      'ORDER BY cobranza_fecha_prog
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
  TxtTDC.Text = GlTipoCambioMercado ' GlTipoCambioOficial
  
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
            FraGrabarCancelar1.Visible = False
            OptFilGral01.Value = True
            Call OptFilGral01_Click
        Case 1
            If glusuario = "RVALDIVIEZO" Or glusuario = "ADMIN" Or glusuario = "HBUSTILLOS" Or glusuario = "MVALDIVIA" Or glusuario = "SPAREDES" Or glusuario = "VPAREDES" Or glusuario = "DLAURA" Or glusuario = "APALACIOS" Or glusuario = "FACTURACION" Then
                lbl_titulo = SSTab1.Caption
'                lbl_titulo2 = SSTab1.Caption
                FraNavega.Caption = SSTab1.Caption
                FraGrabarCancelar.Visible = False
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
            FraGrabarCancelar2.Visible = False
            OptFilGral03.Value = True
            Call OptFilGral03_Click
'            If glusuario = "ADMIN" Then
'                BtnVer.Visible = True
'            Else
'                BtnVer.Visible = False
'            End If
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

Private Sub TxtDscto1_LostFocus()
    TxtMonto1.Text = Round(CDbl(TxtDscto1.Text) * GlTipoCambioMercado, 2)
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
        TxtMontoDol = Round(CDbl(TxtMonto.Text) / CDbl(Txt_tdc), 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub
'adelante
Private Function CodigoControl(NAuto As String, NFactura As String, Nit As String, Fecha As String, Monto As String, Key As String) As String
Dim Suma As Currency
'Dim Suma As String
'
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
  'se cambio de & a + en 13/04/2017
  Suma = Suma + CDbl(Cadena)
  'Suma = Suma & Cadena
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
      Case 1: Cadena = LTrim(Cadena) & LTrim(NAuto) & Mid(Key, Pos, Nro)
      Case 2: Cadena = LTrim(Cadena) & LTrim(NFactura) & Mid(Key, Pos, Nro)
      Case 3: Cadena = LTrim(Cadena) & LTrim(Nit) & Mid(Key, Pos, Nro)
      Case 4: Cadena = LTrim(Cadena) & LTrim(Fecha) & Mid(Key, Pos, Nro)
      Case 5: Cadena = LTrim(Cadena) & LTrim(Monto) & Mid(Key, Pos, Nro)
    End Select
    Pos = Pos + Nro
  Next i

  Cadena = AllegedRC4(Cadena, (Key & NroVer))

  
  SumTot = 0
  i = 0
  Do While i < Len(Cadena)
    i = i + 1
    SumTot = SumTot + Asc(Mid(Trim(Cadena), i, 1))
    sino = Mid(Trim(Cadena), i, 1)
  Loop


  
  For i = 1 To 5
    SumPar(i) = 0
    sino = ""
    
    j = i
    Do While j <= Len(Cadena)
    
      SumPar(i) = SumPar(i) + Asc(Mid(Cadena, j, 1))
      sino = Asc(Mid(Cadena, j, 1))
      Caracter = Mid(Cadena, j, 1)
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

Private Sub TxtMonto1_LostFocus()
    TxtDscto1.Text = Round(CDbl(TxtMonto1.Text) / GlTipoCambioMercado, 2)
End Sub

Private Sub TxtMontoDol_Change()
    'TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(txt_tdc.Text)
End Sub

Private Sub TxtMontoDol_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtMontoDol_LostFocus()
    TxtMonto.Text = CDbl(TxtMontoDol.Text) * CDbl(Txt_tdc.Text)
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form mw_ventas_alcance_acta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Administrativos - Instalaciones - Acta de Entrega Definitiva"
   ClientHeight    =   10740
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   16845
   Icon            =   "mw_ventas_alcance_acta.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   5.66013e6
   ScaleMode       =   0  'User
   ScaleWidth      =   4.8214e8
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PLAN DE ITEMS PARA CRONOGRAMA GRATUITO"
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
      Height          =   2145
      Left            =   120
      TabIndex        =   89
      Top             =   7560
      Width           =   16455
      Begin MSDataGridLib.DataGrid DtgCobro 
         Bindings        =   "mw_ventas_alcance_acta.frx":058A
         Height          =   1860
         Left            =   120
         TabIndex        =   90
         Top             =   240
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   3281
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "No.Cuota"
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Mes.Programado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "mmm-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto Programado Bs."
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
         BeginProperty Column03 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Beneficiario"
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Resp."
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
            DataField       =   "cobranza_fecha_conformidad"
            Caption         =   "Fecha.Certif."
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "cobranza_concepto_plazo"
            Caption         =   "Plazo a Cumplir"
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
         BeginProperty Column09 
            DataField       =   "estado_ac"
            Caption         =   "Aviso Cob."
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
            DataField       =   "correl_ac"
            Caption         =   "Nro. Aviso Cob"
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
            DataField       =   "cobranza_programada_dol"
            Caption         =   "Monto a Pagar Dol."
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
         BeginProperty Column12 
            DataField       =   "cobranza_codigo"
            Caption         =   "Cod.Cobranza"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   5955.024
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   120
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   63
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnVer2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7680
         Picture         =   "mw_ventas_alcance_acta.frx":05A4
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   88
         ToolTipText     =   "Registra Adenda o Modificación al Contrato"
         Top             =   40
         Width           =   1455
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18000
         Picture         =   "mw_ventas_alcance_acta.frx":12F1
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   69
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "mw_ventas_alcance_acta.frx":1AB3
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   68
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         Picture         =   "mw_ventas_alcance_acta.frx":2268
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   67
         ToolTipText     =   "Aprueba el Registro Elegido"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "mw_ventas_alcance_acta.frx":2A9B
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   66
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -15
         Picture         =   "mw_ventas_alcance_acta.frx":31E7
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   65
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         Picture         =   "mw_ventas_alcance_acta.frx":3AFC
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   64
         ToolTipText     =   "Imprimir el Listado de Actas de Entrega Definitiva"
         Top             =   0
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13800
         TabIndex        =   70
         Top             =   180
         Width           =   885
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
      ScaleWidth      =   20280
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "mw_ventas_alcance_acta.frx":43C9
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   61
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "mw_ventas_alcance_acta.frx":4CB5
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   60
         Top             =   0
         Width           =   1280
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13755
         TabIndex        =   62
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   7.5
      ScaleMode       =   4  'Character
      ScaleWidth      =   15.625
      TabIndex        =   49
      Top             =   5655
      Width           =   1935
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modifica Equipo"
         Height          =   720
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Modifica Detalle del Equipo"
         Top             =   120
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Anular-->"
         Height          =   640
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   1485
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Codificar"
         Height          =   640
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Codifica Equipos"
         Top             =   795
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4770
      Left            =   5880
      TabIndex        =   6
      Top             =   765
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8414
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   12632256
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
      TabCaption(0)   =   "REGISTRO DE ACTA DEFINITIVA DE ENTREGA"
      TabPicture(0)   =   "mw_ventas_alcance_acta.frx":548B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame FrmCabecera 
         BackColor       =   &H00C0C0C0&
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
         Height          =   4350
         Left            =   40
         TabIndex        =   9
         Top             =   360
         Width           =   11055
         Begin VB.TextBox Txt_campo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "unidad_codigo_ant"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
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
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   9000
            TabIndex        =   71
            Text            =   "0"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   6840
            TabIndex        =   48
            Top             =   370
            Width           =   350
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   10600
            TabIndex        =   45
            Top             =   1035
            Width           =   330
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   6000
            TabIndex        =   44
            Top             =   1030
            Width           =   330
         End
         Begin MSDataListLib.DataCombo Dtc_deudor2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7845
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
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
            Bindings        =   "mw_ventas_alcance_acta.frx":54A7
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9660
            TabIndex        =   36
            Top             =   1380
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Frame Fra_datos 
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
            ForeColor       =   &H00C00000&
            Height          =   1845
            Left            =   120
            TabIndex        =   20
            Top             =   1395
            Width           =   10815
            Begin MSComCtl2.DTPicker DTPfechaFin 
               DataField       =   "fecha_fin_real"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   4920
               TabIndex        =   87
               Top             =   1440
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               Format          =   108855297
               CurrentDate     =   44334
            End
            Begin MSComCtl2.DTPicker DTPfechasol 
               DataField       =   "fecha_inicio_real"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   1440
               TabIndex        =   86
               Top             =   1440
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               Format          =   108855297
               CurrentDate     =   44334
            End
            Begin VB.TextBox Txt_Campo1 
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   9360
               TabIndex        =   84
               Text            =   "0"
               Top             =   1440
               Width           =   975
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   7800
               TabIndex        =   1
               Top             =   240
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin VB.TextBox TxtPlazo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_tiempo_dias"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   8640
               TabIndex        =   0
               Text            =   "0"
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtConcepto 
               DataField       =   "venta_descripcion"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   10200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Top             =   585
               Visible         =   0   'False
               Width           =   495
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   6840
               TabIndex        =   46
               Top             =   240
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_aux4 
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3240
               TabIndex        =   47
               Top             =   720
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "tipoben_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "DataCombo1"
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               DataField       =   "doc_codigo_alcance"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   7800
               TabIndex        =   85
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Doc.ISO y Nro.Acta Entrega"
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
               Index           =   6
               Left            =   7800
               TabIndex        =   83
               Top             =   1080
               Width           =   2760
               WordWrap        =   -1  'True
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Tiempo en Días"
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
               Index           =   5
               Left            =   7080
               TabIndex        =   82
               Top             =   600
               Width           =   1620
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               DataField       =   "fecha_fin_alcance"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   4920
               TabIndex        =   78
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Fin"
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
               Index           =   4
               Left            =   3960
               TabIndex        =   80
               Top             =   600
               Width           =   1020
               WordWrap        =   -1  'True
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Inicio"
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
               Index           =   3
               Left            =   240
               TabIndex        =   79
               Top             =   600
               Width           =   1140
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label1"
               DataField       =   "fecha_inicio_alcance"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   1440
               TabIndex        =   77
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Fin"
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
               Index           =   2
               Left            =   3960
               TabIndex        =   76
               Top             =   1440
               Width           =   1020
               WordWrap        =   -1  'True
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Inicio"
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
               Index           =   1
               Left            =   240
               TabIndex        =   75
               Top             =   1440
               Width           =   1140
               WordWrap        =   -1  'True
            End
            Begin VB.Label lbl_campo4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fechas Estimadas del Contrato para Mantenimiento Gratuito:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Left            =   240
               TabIndex        =   29
               Top             =   195
               Width           =   6285
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Fechas Reales para Mantenimiento Gratuito:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   240
               Index           =   0
               Left            =   240
               TabIndex        =   21
               Top             =   1035
               Width           =   4980
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Fra_Total 
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
            ForeColor       =   &H00800000&
            Height          =   1095
            Left            =   120
            TabIndex        =   11
            Top             =   3180
            Width           =   10815
            Begin VB.TextBox Text13 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   4700
               TabIndex        =   74
               Top             =   620
               Width           =   350
            End
            Begin VB.TextBox TxtBstotalUsd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_saldo_p_cobrar_dol"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8325
               Locked          =   -1  'True
               TabIndex        =   72
               Text            =   "0"
               Top             =   280
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.TextBox TxtCobradoUsd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_cobrado_dol"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6240
               TabIndex        =   55
               Text            =   "0"
               Top             =   280
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.TextBox TxtMontoUsd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_total_dol"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               TabIndex        =   54
               Text            =   "0"
               Top             =   285
               Width           =   1545
            End
            Begin VB.TextBox txtTDC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               DataField       =   "venta_tipo_cambio"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   8760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   28
               Top             =   420
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox TxtCobrado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_cobrado_bs"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   15
               Text            =   "0"
               Top             =   675
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.TextBox txtCantTotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   2760
               TabIndex        =   14
               Text            =   "0"
               Top             =   300
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox TxtMontoBs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_total_bs"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8880
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "0"
               Top             =   675
               Width           =   1545
            End
            Begin VB.TextBox TxtBstotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_saldo_p_cobrar_bs"
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   8325
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "0"
               Top             =   675
               Visible         =   0   'False
               Width           =   1545
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               Bindings        =   "mw_ventas_alcance_acta.frx":54C0
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   240
               TabIndex        =   73
               Top             =   600
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "venta_tipo_descripcion"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               Bindings        =   "mw_ventas_alcance_acta.frx":54DA
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3720
               TabIndex        =   81
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Locked          =   -1  'True
               Appearance      =   0
               BackColor       =   12632256
               ListField       =   "venta_tipo"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7845
               TabIndex        =   57
               Top             =   315
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5760
               TabIndex        =   56
               Top             =   315
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Contato Moneda Nacional (Bs.) :"
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
               Height          =   285
               Left            =   5520
               TabIndex        =   53
               Top             =   690
               Width           =   3255
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00400000&
               X1              =   5355
               X2              =   5355
               Y1              =   1080
               Y2              =   120
            End
            Begin VB.Label Label21 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Modalidad del Contrato"
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
               Height          =   285
               Left            =   240
               TabIndex        =   19
               Top             =   195
               Width           =   2415
            End
            Begin VB.Label lbl_totalBs 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Contrato Moneda Extranjera (USD):"
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
               Height          =   285
               Left            =   5520
               TabIndex        =   18
               Top             =   285
               Width           =   3255
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   5775
               TabIndex        =   17
               Top             =   645
               Visible         =   0   'False
               Width           =   405
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
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
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   7845
               TabIndex        =   16
               Top             =   645
               Visible         =   0   'False
               Width           =   405
            End
         End
         Begin VB.TextBox txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "venta_codigo"
            DataSource      =   "Ado_datos"
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
            Left            =   7425
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   345
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "mw_ventas_alcance_acta.frx":54F4
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6480
            TabIndex        =   30
            Top             =   1020
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "mw_ventas_alcance_acta.frx":550D
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4440
            TabIndex        =   33
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
            Bindings        =   "mw_ventas_alcance_acta.frx":5526
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1755
            TabIndex        =   34
            Top             =   360
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo Dtc_aux2 
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   39
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
            Bindings        =   "mw_ventas_alcance_acta.frx":553F
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4860
            TabIndex        =   41
            Top             =   1200
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "mw_ventas_alcance_acta.frx":5558
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4980
            TabIndex        =   42
            Top             =   1020
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "edif_codigo_corto"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "mw_ventas_alcance_acta.frx":5571
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   180
            TabIndex        =   43
            Top             =   1020
            Width           =   5085
            _ExtentX        =   8969
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lbl_cerrado 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "TRAMITE CERRADO !!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   2040
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   7395
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   8865
            X2              =   8865
            Y1              =   0
            Y2              =   1695
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Edificio:"
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
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cite Contrato"
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
            Height          =   240
            Left            =   9255
            TabIndex        =   38
            Top             =   75
            Width           =   1365
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   180
            TabIndex        =   35
            Top             =   345
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tramite"
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
            Left            =   180
            TabIndex        =   32
            Top             =   75
            Width           =   690
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
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
            Left            =   1785
            TabIndex        =   31
            Top             =   120
            Width           =   1680
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Venta"
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
            Height          =   285
            Left            =   7500
            TabIndex        =   23
            Top             =   75
            Width           =   1125
         End
         Begin VB.Label lbl_campo2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
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
            Left            =   6465
            TabIndex        =   22
            Top             =   795
            Width           =   660
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   4800
      Left            =   135
      TabIndex        =   24
      Top             =   720
      Width           =   5745
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
         Left            =   3600
         TabIndex        =   27
         Top             =   4520
         Width           =   915
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
         Left            =   1320
         TabIndex        =   26
         Top             =   4520
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "mw_ventas_alcance_acta.frx":558A
         Height          =   4170
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   7355
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "#Tramite"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre de Edificio"
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
            Caption         =   "Fecha.Venta"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Contrato"
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
            DataField       =   "estado_acta"
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
         BeginProperty Column06 
            DataField       =   "edif_codigo"
            Caption         =   "Cod.Edificio"
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
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4440
         Width           =   5520
         _ExtentX        =   9737
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
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE EQUIPOS"
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
      Height          =   1935
      Left            =   2160
      TabIndex        =   7
      Top             =   5580
      Width           =   14895
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "mw_ventas_alcance_acta.frx":55A2
         Height          =   1665
         Left            =   240
         TabIndex        =   8
         Top             =   225
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
            Caption         =   "Codigo.Equipo"
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
            Caption         =   "Descripcion y Características del Equipo"
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
            DataField       =   "venta_precio_unitario_dol"
            Caption         =   "Prec.Unitario.Usd"
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
            DataField       =   "venta_precio_total_bs"
            Caption         =   "Precio Total.Bs"
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
            DataField       =   "venta_precio_total_dol"
            Caption         =   "Precio.Total.USD"
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
            Caption         =   "Modelo.Equipo"
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
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   4185.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   585.071
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
      Top             =   10200
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
      Top             =   10200
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
      Top             =   10920
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
      Top             =   10560
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
      Top             =   10560
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
      Top             =   10920
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
      Top             =   10560
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
      Top             =   10200
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
      Top             =   10560
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
      Top             =   10560
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
      Top             =   10200
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
      Top             =   10200
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
      Left            =   -120
      Top             =   12960
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
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   4560
      Top             =   10920
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   11400
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   13800
      Top             =   10560
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "mw_ventas_alcance_acta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'Ventas
Dim rs_datos As New ADODB.Recordset     'av_ventas_cabecera - VENTAS
Dim rs_datos1 As New ADODB.Recordset    'gp_listar_apr_gc_unidad_ejecutora  - UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'gp_listar_gc_beneficiario_personas - Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'gp_listar_apr_gc_edificaciones - Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'gp_listar_gc_beneficiario_funcionario  - Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos5 As New ADODB.Recordset    'Calculo de Trafico
Dim rs_datos6 As New ADODB.Recordset    'ao_ventas_alcance
Dim rs_datos7 As New ADODB.Recordset    'ao_solicitud_cotiza_venta
Dim rs_datos8 As New ADODB.Recordset    'ao_compra_cabecera
Dim rs_datos11 As New ADODB.Recordset   'ac_tipo_compra_venta
Dim rs_datos12 As New ADODB.Recordset   'Gc_tipo_beneficiario
Dim rs_datos13 As New ADODB.Recordset   'Av_almacen_detalle
Dim rs_datos14 As New ADODB.Recordset   'ao_ventas_detalle  - Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset   'ac_bienes      'av_solicitud_cotiza_venta (antes)
Dim rs_datos16 As New ADODB.Recordset   'ao_ventas_cobranza_inst    - Ventas cobranzas Prog
Dim rs_datos17 As New ADODB.Recordset   'ac_bienes_grupo
Dim rs_datos18 As New ADODB.Recordset   'ao_solicitud_cotiza_venta
Dim rs_datos19 As New ADODB.Recordset   'ao_ventas_cobranza_inst    - Acumula Cobranzas Prog
Dim rs_datos20 As New ADODB.Recordset   'ao_solicitud_costos    - Acumula Costos

'AUXILIARES
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset      ' Verif. Prog
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset
Dim rs_aux13 As New ADODB.Recordset

Dim rstdestino As New ADODB.Recordset       'ao_compra_detalle
Dim rstcorrel_ing As New ADODB.Recordset    'fc_organismo_financiamiento - Correl

'OTROS
Dim rs_det2 As New ADODB.Recordset          'Adjudica Compra
Dim rs_det3 As New ADODB.Recordset          'Adjudica Compra Detalle
Dim rstdetsalalm As New ADODB.Recordset     'ao_detallesalidaalmacen
Dim RS_BENEF As New ADODB.Recordset         'gc_beneficiario - Deudor?
Dim rs_TipoCambio As New ADODB.Recordset    'gc_tipo_cambio
Dim rs_almacen2 As New ADODB.Recordset      'ao_almacen_totales
Dim rstacumdet As New ADODB.Recordset       'ao_ventas_detalle  -   Acumula
Dim rsAuxDetalle As New ADODB.Recordset     'ao_ventas_detalle  -   Para Almacen
Dim rsNada As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
Dim queryinicial2 As String

'Almacenes
Dim descri_bien As String
Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1 As Integer

'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2 As Integer
Dim nroventa, correlv, correldet2 As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CANT0, VAR_CANT9  As Integer
Dim VAR_CODANT, Var_Comp, VAR_SOL, VAR_ZPILOTO As Integer
Dim VAR_NUM, VAR_SOLTIPO As Integer
Dim VAR_COMPM As Long

Dim VAR_DCORR, VAR_HCORR As String

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double
Dim VAR_AUX4, VAR_AUX5 As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_COD2, VAR_COD3, VAR_COD4 As String
Dim VAR_TIPOV, VAR_UNIMED As String
Dim VAR_COBR0, VAR_OA, VAR_OA2, VAR_NEW As String
Dim VAR_PAIS, VAR_EQP, VAR_TIPOEQP As String
Dim VAR_DA, VAR_UORIGEN As String
Dim VAR_NOMD, VAR_NOMH As String
Dim VAR_JQ, VAR_VAL, VAR_FCONTROL As String
    
Private Sub CmdDetalle_Click()
    FrmCobranza.Visible = True
End Sub

Private Sub adosalalm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
        Select Case pRecordset.EditMode
        Case adEditNone
            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
            Set DataGrid2.DataSource = Nothing
            Set DataGrid2.DataSource = rstdetsalalm
            DataGrid2.ReBind
        End Select
End Sub

Private Sub Adodetallesolicitud_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not adoDetalleSolicitud.Recordset.BOF) And (Not adoDetalleSolicitud.Recordset.EOF) Then
        If Not IsNull(adoDetalleSolicitud.Recordset("correlativo_solicitud")) Then
            txtnosolicitud1.Text = adoDetalleSolicitud.Recordset("correlativo_solicitud")
            txtcorrdet.Text = adoDetalleSolicitud.Recordset("correlativo_detalle")
        Else
            txtnosolicitud1.Text = Ado_datos.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            dtccodpar.Text = " "
            dtcdescripar.Text = " "
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim descri_bien As String
Dim Cant_Alm As Integer
If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
   If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then
        nroventa = Ado_datos.Recordset!venta_codigo
        lbl_cerrado.Caption = ""
        If (Ado_datos.Recordset("estado_codigo") = "REG") Then
            BtnAprobar.Visible = True
'                BtnDesAprobar.Visible = False
            BtnModificar.Visible = True
            BtnEliminar.Visible = True
'            BtnVer.Visible = False
'            BtnModDetalle1.Visible = True
            If IsNull(Ado_datos.Recordset("venta_tipo")) Then
                FrmABMDet.Visible = False
                FrmABMDet1.Visible = False
                FrmABMDet2.Visible = False
                FrmCobranza.Visible = False
'                FrmAlcance.Visible = False
            Else
                FrmABMDet.Visible = True
                FrmABMDet1.Visible = True
                FrmABMDet2.Visible = True
                FrmCobranza.Visible = True
'                FrmAlcance.Visible = True
            End If
        Else
        'WWWWWWWWWWWWWWWWWWWWWWWWWW
            Select Case Ado_datos.Recordset!estado_cancelado
                Case "S"
                    lbl_cerrado.Caption = "TRAMITE CERRADO !!"
                    FrmABMDet2.Visible = False
                    BtnAñadir.Visible = False   'Cerrar Tramite
                    BtnVer3.Visible = False     'Provisional
                    FrmABMDet.Visible = False
'                    FraDet2.Visible = False
                    FrmABMDet1.Visible = False
                Case "P"
'                    lbl_cerrado.Caption = "TRAMITE PROVISIONAL !!"
'                    If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "CARIZACA" Then
'                        BtnModificar.Visible = True
'                        FrmABMDet.Visible = True
'                        BtnModDetalle.Visible = True
'                        BtnVer3.Visible = True     'Provisional
'                    Else
'                        BtnModificar.Visible = False
'                        FrmABMDet.Visible = False
'                        BtnModDetalle.Visible = False
'                        BtnVer3.Visible = False 'Provisional
'                    End If
'                    FrmABMDet2.Visible = True
'                    BtnAñadir.Visible = False   'Cerrar Tramite
                    
                Case Else
                    If glusuario = "MVALDIVIA" Or glusuario = "ADMIN" Or glusuario = "SPAREDES" Or glusuario = "DLAURA" Or glusuario = "MCOLLAO" Then
'                        BtnAñadir.Visible = True   'Cerrar Tramite
                        'BtnVer3.Visible = True     'Provisional
                    Else
                        'BtnVer3.Visible = False     'Provisional
                    End If
                    lbl_cerrado.Caption = ""
'                    FrmABMDet2.Visible = True
                    'FrmABMDet.Visible = True
                    'FraDet2.Visible = True
                    'FrmABMDet1.Visible = True
            End Select
'            BtnAprobar.Visible = False
'                BtnDesAprobar.Visible = True
'            BtnModificar.Visible = False
            BtnEliminar.Visible = False
'            BtnVer.Visible = True
'            BtnModDetalle1.Visible = False
            FrmABMDet.Visible = False
'            FrmABMDet1.Visible = False
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
'            FrmAlcance.Visible = True
            If (Ado_datos.Recordset!estado_codigo = "APR") Then
'                'CRONOGRAMA COMPRA SERVICIO
''                FraDet2.Visible = True
'                'Compra Cabecera Funcionario - Vendedor
'                Set rs_datos8 = New ADODB.Recordset
'                If rs_datos8.State = 1 Then rs_datos8.Close
'                rs_datos8.Open "select * from ao_compra_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenStatic
'                'Set Ado_datos4.Recordset = rs_datos8
'                'Compra Adjudica
'                Set rs_det2 = New ADODB.Recordset
'                If rs_det2.State = 1 Then rs_det2.Close
'                rs_det2.Open "select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'                Set Ado_detalle2.Recordset = rs_det2
'                If Ado_detalle2.Recordset.RecordCount > 0 Then
'                    Set rs_det3 = New ADODB.Recordset
'                    If rs_det3.State = 1 Then rs_det3.Close
'                    rs_det3.Open "select * from ao_compra_planilla_pagos where compra_codigo = " & rs_det2!compra_codigo & " and adjudica_codigo = " & rs_det2!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'                    Set Ado_detalle3.Recordset = rs_det3
''                    If Ado_detalle3.Recordset.RecordCount > 0 Then
''                        dg_det3.Visible = True
''                        Set dg_det3.DataSource = Ado_detalle3.Recordset
''                    Else
''                        dg_det3.Visible = False
''                        Set dg_det3.DataSource = rsNada
''                    End If
''                    dg_det2.Visible = True
''                    Set dg_det2.DataSource = Ado_detalle2.Recordset
                Else
'                    dg_det3.Visible = False
'                    Set dg_det3.DataSource = rsNada
'                    dg_det2.Visible = False
'                    Set dg_det2.DataSource = rsNada
                End If
            End If
        End If
'            If Ado_datos.Recordset("estado_codigo") = "APR" Then
'                BtnAprobar.Enabled = False
''                BtnDesAprobar.Enabled = False
'                FrmABMDet.Visible = False
'                BtnModDetalle.Visible = False
'                BtnAnlDetalle.Visible = False
'            Else
'                BtnAprobar.Enabled = True
'                FrmABMDet.Visible = True
'                BtnModDetalle.Visible = True
'                BtnAnlDetalle.Visible = True
'            End If
'            If (Ado_datos.Recordset("venta_tipo") = "C") And Ado_datos.Recordset("estado_codigo") = "APR" Then
'                FrmABMDet2.Visible = True
'                FrmCobranza.Visible = True
'            Else
'                FrmABMDet2.Visible = False
'                FrmCobranza.Visible = False
'            End If
        If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Or (Ado_datos.Recordset("venta_tipo") = "G") Or (Ado_datos.Recordset("venta_tipo") = "L") Then
            TxtPlazo.Visible = True
'            BtnAddDetalle2.Visible = True
        Else
            TxtPlazo.Visible = False
            If Ado_datos.Recordset("venta_tipo") = "E" Then
                BtnAddDetalle2.Visible = False
            End If
        End If
        
        If Dtc_deudor2.Text = "SI" Then
            Dtc_deudor2.backColor = &HFF&
        Else
            Dtc_deudor2.backColor = &H80000010
        End If
        'If Ado_datos.Recordset("beneficiario_codigo") <> "" And Ado_datos.Recordset("beneficiario_codigo") <> "VD" Then
        If Ado_datos.Recordset("beneficiario_codigo") <> "" Then
            Set RS_BENEF = New ADODB.Recordset
            If RS_BENEF.State = 1 Then RS_BENEF.Close
            RS_BENEF.Open "select * from gc_beneficiario where beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
            'RS_BENEF.Recordset.Requery
            If RS_BENEF.RecordCount > 0 Then
                If RS_BENEF!beneficiario_deudor = "SI" Then
                    Dtc_deudor2.backColor = &HFF&
                Else
                    Dtc_deudor2.backColor = &H80000010
                End If
            End If
            
        End If
        GlEdificio = Ado_datos.Recordset!edif_codigo
        Call ABRIR_TABLA_DET
'        FrmDetalle.Caption = "BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'        FrmCobranza.Caption = "CRONOGRAMA DE COBRANZAS DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        FrmDetalle.Caption = "BIENES DEL TRAMITE NRO. " + Str((Ado_datos.Recordset("solicitud_codigo")))
'        FrmCobranza.Caption = "CRONOGRAMA DE COBRANZAS DE TRAMITE NRO. " + Str((Ado_datos.Recordset("solicitud_codigo")))
'        Else
        End If
        'GlEdificio = Ado_datos.Recordset!edif_codigo
        FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
'        FrmAlcance.Visible = True
'    Else
'        FrmABMDet.Visible = False
'        FrmABMDet1.Visible = False
'        FrmABMDet2.Visible = False
''        FrmAlcance.Visible = False
'        FrmDetalle.Visible = False
'        FrmCobranza.Visible = False
'    End If
End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_datos14 = New ADODB.Recordset
    If rs_datos14.State = 1 Then rs_datos14.Close
    rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & nroventa & "'  ", db, adOpenKeyset, adLockOptimistic
    'rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & correlv & "'  ", db, adOpenKeyset, adLockOptimistic
    'rs_datos14.Open queryinicial2, db, adOpenKeyset, adLockOptimistic
    Set ado_datos14.Recordset = rs_datos14
    Set DtGLista.DataSource = ado_datos14.Recordset
    'ado_datos14.Recordset.Requery
    If ado_datos14.Recordset.RecordCount > 0 Then
        deta2 = 1
        ado_datos14.Recordset.Requery
        'TxtMontoBs.Text = Ado_datos.Recordset!monto_total_bS
        'TxtMontoUs.Text = Ado_datos.Recordset!deuda_cobrada
        'Text2.Text = Ado_datos.Recordset!saldo_p_cobrar
'        Call AbreAlmacen
'        If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Or (Ado_datos.Recordset("venta_tipo") = "G") Or (Ado_datos.Recordset("venta_tipo") = "L") Then
'            FrmABMDet2.Visible = True
'            FrmCobranza.Visible = True
'
'        Else
'            FrmABMDet2.Visible = False
'            FrmCobranza.Visible = False
'        End If
    Else
        deta2 = 0
        'TxtMontoBs.Text = 0
        'TxtMontoUs.Text = 0
        'Text2.Text = 0
        FrmABMDet2.Visible = False
        FrmCobranza.Visible = False
    End If
        
        Set rs_datos6 = New ADODB.Recordset
        If rs_datos6.State = 1 Then rs_datos6.Close
        rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & nroventa & "  ", db, adOpenKeyset, adLockBatchOptimistic
        Set Ado_datos6.Recordset = rs_datos6
'        Set DtgAlcance.DataSource = Ado_datos6.Recordset
        If Ado_datos6.Recordset.RecordCount > 0 Then
'            DtgAlcance.Visible = True
        Else
'            DtgAlcance.Visible = False
        End If
End Sub


Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
' If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
'    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
'        'BtnModDetalle2.Visible = False
'        If (Ado_datos16.Recordset("estado_codigo") = "REG") Then
'            If (Ado_datos.Recordset("estado_codigo") = "APR") Then
'                BtnAprobar2.Visible = False
'            Else
'                BtnAprobar2.Visible = True
'            End If
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = True
'            BtnAnlDetalle2.Visible = True
'            BtnModDetalle2.Visible = True
'        End If
'        If (Ado_datos16.Recordset("estado_codigo") = "APR") Then
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = False
'            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
'        End If
'        If (Ado_datos16.Recordset("estado_codigo") = "ANL") Then
'            BtnImprimir2.Visible = False
'            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
'            BtnAprobar2.Visible = False
'        End If
'    Else
'        BtnAprobar2.Visible = False
'        BtnImprimir2.Visible = False
'        BtnAnlDetalle2.Visible = False
'        BtnModDetalle2.Visible = False
'    End If
' Else
'    BtnAprobar2.Visible = False
'    BtnImprimir2.Visible = False
'    BtnAnlDetalle2.Visible = False
'    BtnModDetalle2.Visible = False
' End If
End Sub

Private Sub BtnAddDetalle_Click()
'  'marca1 = Ado_datos.Recordset.Bookmark
'  If ado_datos14.Recordset!estado_codigo = "REG" Then
'    Set rs_aux6 = New ADODB.Recordset
'    If rs_aux6.State = 1 Then rs_aux6.Close
'    rs_aux6.Open "select * from fc_partida_gasto where par_codigo = '43340' ", db, adOpenKeyset, adLockReadOnly
'    If rs_aux6.RecordCount > 0 Then
'        VAR_OA = "AO36" + LTrim(Str(rs_aux6!correlativo36 + 1))
'        Set rs_aux7 = New ADODB.Recordset
'        If rs_aux7.State = 1 Then rs_aux7.Close
'        rs_aux7.Open "select * from ac_bienes where bien_codigo = '" & VAR_OA & "' ", db, adOpenKeyset, adLockReadOnly
'        If rs_aux7.RecordCount > 0 Then
'            MsgBox "El equipo " + VAR_OA + " YA Existe, vuelva a intentar !! ", vbExclamation, "Atención!"
'            db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
'        Else
'            ado_datos14.Recordset!bien_codigo = Trim(VAR_OA)
'            db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
'            db.Execute "insert into ac_bienes(grupo_codigo, subgrupo_codigo, bien_codigo, par_codigo, bien_descripcion, bien_precio_compra, bien_precio_venta_base, bien_precio_venta_final, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo, bien_stock_minimo, bien_stock_inicial, bien_stock_ingreso, bien_stock_salida, bien_stock_actual, bien_total_compra_bs, bien_total_venta_bs, bien_utilidad_Bs, bien_codigo_anterior, bien_codigo_universal, bien_descripcion_anterior, pais_codigo, archivo_foto2, archivo_foto, estado_codigo, fecha_registro, usr_codigo) " & _
'            "VALUES ('40000', '43000', '" & VAR_OA & "', '43340', 'CAPACIDAD ' + '" & dtc_desc31.Text & "' + ' PERSONAS Y VELOCIDAD ' + '" & dtc_valor41.Text & "' + ' m/s', " & var_cod & ", '0', '0', 'EQP', 'EQP', '1', 'S/M', '1', '0', '0', '0', '0', '0', '0', '0', '-', '-', '-', 'NN', '" & VAR_COD3 & "' + '2.JPG', '" & VAR_COD3 & "' + '.JPG', 'REG', '" & Date & "', '" & glusuario & "') "
'        End If
'    End If
''    'If OptFilGral1.Value = True Then Call OptFilGral1_Click
''    'If OptFilGral2.Value = True Then Call OptFilGral2_Click
'''    Ado_datos.Recordset.Move marca1 - 1
''    swnuevo = 1
''    SSTab1.Tab = 1
''    SSTab1.TabEnabled(1) = True
''    SSTab1.TabEnabled(0) = False
''    SSTab1.TabEnabled(2) = False
''    FrmEdita.Visible = True
''    FrmEdita.Enabled = True
''    FraNavega.Enabled = False
''    FrmDetalle.Enabled = False
''    FrmCobranza.Visible = False
''    FrmABMDet.Visible = False
''    FrmABMDet2.Visible = False
''    'tipo Beneficiario
''    Set rs_datos12 = New ADODB.Recordset
''    If rs_datos12.State = 1 Then rs_datos12.Close
''    'rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
''    rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Dtc_aux2.Text & "' ", db, adOpenKeyset, adLockReadOnly
''    Set Ado_Datos12.Recordset = rs_datos12
''    Ado_Datos12.Refresh
''
''    ado_datos14.Recordset.AddNew
'  Else
'    MsgBox "El registro Aprobado o Anulado, NO pueden ser modificado !! ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnAprobar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'VALIDA EDIFICIO Y EQUIPOS
    If Ado_datos.Recordset!estado_acta <> "REG" Then
        MsgBox "No se puede APROBAR un registro ANULADO O APROBADO, revise vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    Set rs_aux10 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_aux10.State = 1 Then rs_aux10.Close
    rs_aux10.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & GlEdificio & "' and estado_codigo = 'APR' ", db, adOpenStatic
    If rs_aux10.RecordCount = 0 Then
        'Si Faltarian Aprobar
        MsgBox "No se puede APROBAR, verifique los datos del Edificio si estan correctos y si está Aprobado, luego vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    
    Set rs_aux11 = New ADODB.Recordset     'Equipos de Venta_Detalle
    If rs_aux11.State = 1 Then rs_aux11.Close
    rs_aux11.Open "Select * from mv_bienes_vs_venta_det WHERE venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenStatic
    If rs_aux11.RecordCount > 0 Then
        'Si Faltarian Aprobar
        MsgBox "No se puede APROBAR, verifique los datos de los EQUIPOS y si estos están Aprobados, luego vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    
    Set rs_aux12 = New ADODB.Recordset     'Partidas de Venta_Detalle
    If rs_aux12.State = 1 Then rs_aux12.Close
    rs_aux12.Open "Select * from ao_ventas_detalle WHERE venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "' and par_codigo=''  ", db, adOpenStatic
    If rs_aux12.RecordCount > 0 Then
        'Si Faltarian Partida
        MsgBox "No se puede APROBAR, verifique los datos de Detalle de Bienes , luego vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    
'   If IsNull(Ado_datos.Recordset("venta_tipo")) Or Ado_datos.Recordset("venta_tipo") = "" Or (Ado_datos.Recordset("venta_monto_total_bs") = 0) Or (Ado_datos.Recordset!estado_alcance = "N") Or (Ado_datos.Recordset!unidad_codigo_ant = "") Or IsNull(Ado_datos.Recordset!unidad_codigo_ant) Then
'        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'        Exit Sub
'   End If
    If Ado_datos.Recordset("estado_acta") = "REG" Then
        NumComp = Ado_datos.Recordset!venta_codigo
        GlEdificio = Ado_datos.Recordset!edif_codigo
        VAR_EDIF = Ado_datos.Recordset!edif_descripcion
        VAR_COD4 = Ado_datos.Recordset!unidad_codigo
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        VAR_ZPILOTO = Ado_datos.Recordset!zpiloto_codigo
        VAR_MED = "MES"
        VAR_EMPRESA = Ado_datos.Recordset!codigo_empresa
        VAR_TIPO = "6"                                      'Ado_datos.Recordset!tipo_solicitud
       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
       If sino = vbYes Then
           Call CRONO_MTTO
           db.Execute "Update ao_ventas_alcance set estado_acta ='APR', estado_codigo ='REG'  WHERE venta_codigo = " & NumComp & " AND solicitud_tipo = '6' "
           'ASIGNA A VARIABLES CAMPOS CLAVES
'           'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
'            Set rs_aux2 = New ADODB.Recordset
'            If rs_aux2.State = 1 Then rs_aux2.Close
'            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
'            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux2.RecordCount > 0 Then
'                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                Ado_datos.Recordset!doc_numero = rs_aux2!correl_doc
'                'Txt_campo1.Caption = rs_aux2!correl_doc
'                rs_aux2.Update
'            End If
'
           'FIN HABILITA ALMACEN PARA venta_tipo="V" (PREVENTA)
           
'           ' APRUEBA ao_ventas_cabecera
'           db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           'Actualiza Cite Trñamite (unidad_codigo_ant)
'           'FIN GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
'           'Call OptFilGral1_Click
           MsgBox "El Registro fue Aprobado Exitosamente... ", vbInformation, "Información!"
       End If
     End If
   'End If
 Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
 End If

End Sub

Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      'Call OptFilGral1_Click
      OptFilGral2.Value = True
      Call OptFilGral2_Click
      PosibleApliqueFiltro = False
      Dim rsNada As ADODB.Recordset
      Dim GrSqlAux As String
      Set ClBuscaGrid = New ClBuscaEnGridExterno
      Set ClBuscaGrid.Conexión = db
      ClBuscaGrid.EsTdbGrid = False
      Set ClBuscaGrid.GridTrabajo = dg_datos
      ClBuscaGrid.QueryUtilizado = queryinicial
      Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
      ClBuscaGrid.CamposVisibles = "110"
      ClBuscaGrid.Ejecutar
      PosibleApliqueFiltro = True
  Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnCancelar_Click()
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_datos.Recordset("estado_codigo") = "APR" Then
    Call OptFilGral2_Click
  Else
    Call OptFilGral1_Click
  End If
  FraNavega.Enabled = True
  FrmCabecera.Enabled = False
  Fra_datos.Enabled = True
  FrmDetalle.Visible = True
'  FrmCobranza.Visible = True
'  FrmAlcance.Visible = True
  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
'  FrmABMDet1.Visible = True
'  FrmABMDet2.Visible = True
'  TxtCobrado.Visible = False
'  Label7.Visible = False
'  Cmd_Cliente.Visible = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
'  SSTab1.TabEnabled(1) = True
'  SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnEliminar_Click()
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
'    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnGrabar_Click()
  NumComp = Ado_datos.Recordset!venta_codigo
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir un Vendedor !! Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo11 = "" Then
    MsgBox "Debe Elejir el Tipo de Venta!! (Credito, pago ne Efectivo, etc.), Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If dtc_codigo2 = "" Then
    MsgBox "Debe Elejir un Cliente para la Venta!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    Exit Sub
  End If
  If Txt_campo2.Text = "" And Txt_campo2.Text = " " Then
     MsgBox "Debe registrar el CITE de TRAMITE !!,  Vuelva a intentar ...", vbExclamation, "Atención"
  End If
    FrmCabecera.Enabled = False
    Call grabar
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    FrmCabecera.Enabled = False
    Fra_datos.Enabled = True
    dg_datos.Visible = True
    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
'    FrmAlcance.Visible = True
    Fra_Total.Visible = True
    FrmABMDet.Visible = True
'    FrmABMDet1.Visible = True
'    FrmABMDet2.Visible = True
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
'  End If

     'Ado_datos.Recordset.Update
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 Then
     'VAR_SW = ""
        rs_datos.Find "venta_codigo = " & NumComp & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos.MoveLast
     End If
    

End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        'fra_reportes.Visible = True
        
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

'    '    Dim rs As New ADODB.Recordset
'    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
'    '    i = 1
'    '    y = 1
'        Select Case Me.Ado_datos.Recordset!unidad_codigo
'          Case "DNINS"
'              var_titulo = "Módulo Instalaciones"
'          Case "DNAJS"
'              var_titulo = "Módulo Ajustes"
'          Case "DNMAN"
'              var_titulo = "Módulo Mantenimiento"
'          Case "DNREP"
'              var_titulo = "Módulo Reparaciones"
'          Case "DNEME"
'              var_titulo = "Módulo Emergencias"
'          Case "DNMOD"
'              var_titulo = "Módulo Modernización"
'          Case "DVTA", "DCOMS", "DCOMB", "DCOMC"
'              var_titulo = "Módulo Comercial"
'        End Select

        CryV01.ReportFileName = App.Path & "\reportes\comercial\ar_lista_actas_entrega_definitiva.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
'        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
'        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
'        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
    Else
        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub BtnModificar_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
        FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
'        FrmCobranza.Visible = False
'        FrmAlcance.Visible = False
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        Fra_datos.Enabled = True
        'Fra_Total.Visible = False
    '    If Ado_datos.Recordset!venta_tipo = "E" Then
    '        TxtCobrado.Visible = True
    '        Label7.Visible = True
    '    Else
    '        TxtCobrado.Visible = False
    '        Label7.Visible = False
    '    End If
    '    Cmd_Cliente.Visible = True
'        If IsNull(DTPfechasol) Then
'            DTPfechasol.Value = Date
'        End If
        FrmABMDet.Visible = False
'        FrmABMDet1.Visible = False
'        FrmABMDet2.Visible = False
    
        swgrabar = 0
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
    Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        Ado_datos.Recordset.Close
        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_datos14.State = 1 Then rs_datos14.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

Private Sub BtnAnlDetalle_Click()
' If Ado_datos.Recordset!estado_codigo = "REG" Then
'   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
'   If sino = vbYes Then
''     ado_datos14.Recordset.Delete
''     ado_datos14.Recordset.Update
''     rs_datos14.Requery
''     ado_datos14.Refresh
''     'cerea
''     ado_datos14.Refresh
'      db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
'   End If
'  Else
'    MsgBox "Los Bienes del registro Aprobado o Anulado, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnModDetalle_Click()
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    FraNavega.Enabled = False
'    FrmDetalle.Enabled = False
'    FrmCobranza.Visible = False
'    FrmAlcance.Visible = False
'    swgrabar = 0
'    swnuevo = 2
'    'marca1 = Ado_datos.Recordset.Bookmark
'    'txt_descripcion_venta.Enabled = True
'    correlv = Ado_datos.Recordset!venta_codigo
'    TxtNroVenta.Text = correlv  'Ado_datos.Recordset!venta_codigo  'txt_venta.Text
'    TxtNroVenta.Enabled = False
'    'lbltipoVenta.Caption = dtc_desc11.Text
''    lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(2) = False
'    FrmEdita.Visible = True
'    FrmEdita.Enabled = True
'    FrmABMDet.Visible = False
'    FrmABMDet1.Visible = False
'    FrmABMDet2.Visible = False
'    If ado_datos14.Recordset!modelo_elegido = "S" Then
'        OpMod1.Value = True
'        OpMod2.Value = False
'        OpMod3.Value = False
'    End If
'    If ado_datos14.Recordset!modelo_elegido_h = "S" Then
'        OpMod1.Value = False
'        OpMod2.Value = True
'        OpMod3.Value = False
'    End If
'    If ado_datos14.Recordset!modelo_elegido_x = "S" Then
'        OpMod1.Value = False
'        OpMod2.Value = False
'        OpMod3.Value = True
'    End If
'    'dtc_codigo13.Text
'    If ado_datos14.Recordset!par_codigo = "43340" Then
'        dtc_codigo13.Text = "0"
'        dtc_desc13.BoundText = dtc_codigo13.BoundText
'        dtc_desc13.backColor = &H80000013
'        dtc_desc13.ForeColor = &HFFFFFF
'    Else
'        dtc_desc13.backColor = &HFFFFFF
'        dtc_desc13.ForeColor = &H80000008
'    End If
'    Set rs_datos12 = New ADODB.Recordset
'    If rs_datos12.State = 1 Then rs_datos12.Close
'    rs_datos12.Open "select * from Gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    Set Ado_Datos12.Recordset = rs_datos12
'    'Ado_datos12.Refresh
'    Dtc_aux12.BoundText = dtc_codigo12.BoundText
'    dtc_desc12.BoundText = dtc_codigo12.BoundText
'
'    'Solo para Equipos (*)
'    Set rs_datos15 = New ADODB.Recordset
'    If rs_datos15.State = 1 Then rs_datos15.Close
'    rs_datos15.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
'    'rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
'    Set ado_datos15.Recordset = rs_datos15
'    ado_datos15.Refresh
'  Else
'    MsgBox "Los datos del registro Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
'  End If
End Sub

Private Sub BtnVer2_Click()
    'If Ado_datos.Recordset!estado_codigo = "REG" Then
  If Ado_datos.Recordset!estado_acta = "REG" Then
    NumComp = Ado_datos.Recordset!venta_codigo
    GlEdificio = Ado_datos.Recordset!edif_codigo
    'BUSCA ZONA PILOTO
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "SELECT * FROM tc_zona_piloto_edif WHERE EDIF_codigo = '" & GlEdificio & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux2.RecordCount > 0 Then
        VAR_ZPILOTO = rs_aux2!zpiloto_codigo
    Else
'        Set rs_aux18 = New ADODB.Recordset
'        If rs_aux18.State = 1 Then rs_aux18.Close
'        rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZONA & " ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux18.RecordCount > 0 Then
'            VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
'        Else
'            VAR_ORDEN = 1
'        End If
'
'       db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
'                  " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo) " & _
'                  " VALUES (" & VAR_ZPILOTO & ", '" & GlEdificio & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
'                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
                  
        MsgBox "Se necesita asignar la ZONA PILOTO, para generar el Cronograma de Mantenimiento Gratuito... Consulte al Personal del área Técnica ...", , "Atencion"
        Exit Sub
    End If
    'VERIFICA SI EXISTE ITEMS PARA CRONOGRAMA
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select max(cobranza_prog_codigo) as Codigo3 from ao_ventas_cobranza_inst where venta_codigo= " & NumComp & " ", db, adOpenStatic
    If IsNull(rs_aux3!codigo3) Then
        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
        corrprog = 0
        Call CRONO2
    Else
        sino = MsgBox("El Cronograma ya existe, desea volver a Generarlo ? (los items Aprobados no serán modificados)...", vbYesNo + vbQuestion, "Atención ...")
        If sino = vbYes Then
            'OJO BORRAR ao_ventas_cobranza_inst
            db.Execute "DELETE ao_ventas_cobranza_inst where venta_codigo= " & NumComp & " and estado_codigo = 'REG' "
            db.Execute "update ao_ventas_cobranza_inst set venta_codigo_new = cobranza_prog_codigo where venta_codigo= " & NumComp & " "
            db.Execute "update ao_ventas_cobranza_inst set cobranza_prog_codigo = venta_codigo_new + 100 where venta_codigo= " & NumComp & " "
            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & NumComp & " "
            'db.Execute "update ao_ventas_cabecera set tipo_moneda = 'BOB' where venta_codigo= " & NumComp & " "
            corrprog = 0
            Call CRONO2
        Else
        'If rs_aux3!codigo3 > corrprog Then
            'ACTUALIZAR CORRELATIVO CRONO
            db.Execute "update ao_ventas_cabecera set correl_cobro_prog = " & rs_aux3!codigo3 & " where venta_codigo= " & NumComp & " "
            'wwwwwwwwwwwwwwwwwww
            db.Execute "UPDATE ao_ventas_cobranza_inst SET ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' where venta_codigo = " & NumComp & " and ges_gestion <> '" & Ado_datos.Recordset!ges_gestion & "'  "
            'db.Execute "update ao_ventas_cabecera set estado_codigo_verif = 'APR' Where venta_codigo = " & NumComp & " "
            'wwwwwwwwwwwwwwwwwww
            corrprog = rs_aux3!codigo3
        'End If
        End If
    End If
  Else
    MsgBox "NO se puede procesar, el trámite ya fue APROBADO o ANULADO ...", , "Atencion"
  End If
End Sub

Private Sub CRONO2()
'    Set rs_aux5 = New ADODB.Recordset
'    If rs_aux5.State = 1 Then rs_aux5.Close
'    'rs_aux5.Open "select * from ao_ventas_cabecera where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
'    rs_aux5.Open "select * from av_ventas_alcance where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
'    'Set AdoAux.Recordset = rsAuxDetalle
'    If rs_aux5.RecordCount > 0 Then
    If IsNull(Ado_datos.Recordset!fecha_inicio_real) Then                   ' Is Null
        MsgBox "Debe registrar la Fecha Inicio Real, verifique y vuelva a intentar ...", , "Atencion"
        Exit Sub
    End If
      CONT2 = 1
      'FInicio = DTPfechasol.Value                        'Fecha Inicio Alcance
      FInicio = Ado_datos.Recordset!fecha_inicio_real                       '
'      CANTOT = rs_aux5!venta_cantidad_total
      gestion0 = Year(FInicio)                              'Ado_datos.Recordset!ges_gestion
      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
      VAR_MED2 = "MES"                                          'Ado_datos.Recordset!unimed_codigo_cobr
      'MES = DateDiff("M", fecha, Date)
      VAR_COBR2 = DateDiff("M", FInicio, Ado_datos.Recordset!fecha_fin_real)
      MControl = UCase(MonthName(Month(FInicio)))                        'Ado_datos.Recordset!mes_inicio_crono
      VAR_MES2 = Month(FInicio)
      FControl = RTrim("01/" + RTrim(Str(Month(FInicio))) + "/" + Str(Year(FInicio)))
      VAR_FCONTROL = Format(FControl, "dd/mm/yyyy")
      FControl = VAR_FCONTROL
      CONT3 = 0
      CONT4 = 0
      CONT_MED = 1              'MES = 1 (Mensual)
      Set rs_aux2 = New ADODB.Recordset
      If rs_aux2.State = 1 Then rs_aux2.Close
      rs_aux2.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "  ", db, adOpenKeyset, adLockOptimistic
      If rs_aux2.RecordCount = 0 Then
        db.Execute "UPDATE ao_ventas_cabecera SET correl_cobro_prog = '0' WHERE  venta_codigo = " & NumComp & " "
        correldet2 = "0"
      End If
      While (CONT2 <= VAR_COBR2)
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "  ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 And corrprog >= VAR_COBR2 Then
            MsgBox "El Cronograma ya fue generado... ", , "Atención"
            CONT2 = CONT2 + 1
        Else
           'wwwwwwwwwwwwwwwwwwwwww
'          If correldet2 > 0 Then
'            correldet2 = Ado_datos.Recordset!correl_cobro_prog + 1                          'rs_aux5!correl_cobro_prog + 1
'          End If
          correldet2 = correldet2 + 1
          corrprog = correldet2
          db.Execute "UPDATE ao_ventas_cabecera SET correl_cobro_prog = " & corrprog & " WHERE  venta_codigo = " & NumComp & " "

          Set rs_aux8 = New ADODB.Recordset
          If rs_aux8.State = 1 Then rs_aux8.Close
          rs_aux8.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(VAR_FCONTROL) & "'  AND MONTH(cobranza_fecha_prog) = '" & Month(VAR_FCONTROL) & "'  ", db, adOpenKeyset, adLockOptimistic
          If rs_aux8.RecordCount = 0 Then
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            'rs_aux2.AddNew
            'rs_aux2!cobranza_prog_codigo = correldet2
            'rs_aux2!beneficiario_codigo = VAR_BENEF                   'Codigo Beneficiario/Cliente
            ''OJO MODIFICAR COBRADOR - JQA 03-ENE-2015
            'rs_aux2!beneficiario_codigo_resp = IIf(dtc_codigo5.Text = "", "4333735", dtc_codigo5.Text)      ''Codigo Cobrador
            'Ado_datos.Recordset!correl_cobro_prog = corrprog
            
            fdia = Day(VAR_FCONTROL)
            fanio = Year(VAR_FCONTROL)
            'CONT3 = CONT2 * CONT_MED
            CONT3 = 1
            While (CONT3 <= CONT_MED)
                fmes = Month(VAR_FCONTROL)
                Select Case fmes
                    Case 2
                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
                            Dias_Mes = 29
                        Else
                            Dias_Mes = 28
                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        Dias_Mes = 31
                    Case 4, 6, 9, 11
                        Dias_Mes = 30
                End Select
                If Val(VAR_MES2) = Month(VAR_FCONTROL) Then
                    'rs_aux2!cobranza_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_cobro = FControl + 20
                    FControl = Format(VAR_FCONTROL, "dd/mm/yyyy")
                    VAR_MES2 = VAR_MES2 + CONT_MED
                    If Val(VAR_MES2) > 12 Then
                        VAR_MES2 = Val(VAR_MES2) - 12
                    End If
                End If
                'FControl = FControl + Dias_Mes
                VAR_FCONTROL = CDate(VAR_FCONTROL) + Dias_Mes
                CONT3 = CONT3 + 1
                CONT4 = CONT4 + Dias_Mes
            Wend
            'FControl = Str(fdia) + "/" + Str(fmes) + "/" + Str(fanio)
            'rs_aux2!cobranza_fecha_prog = FInicio + (30 * CONT2)
            'rs_aux2!cobranza_fecha_prog = FControl
            'If Ado_datos.Recordset!cobranza_fecha_prog = Null Then
                'rs_aux2!cobranza_fecha_prog = Date
                'FControl = Date
            '    VAR_FCONTROL = Date
            'End If
            'rs_aux2!gestion = Year(rs_aux2!cobranza_fecha_prog)
            'rs_aux2!cobranza_mes = Month(rs_aux2!cobranza_fecha_prog)
            
            ''VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), Date, rs_aux2!cobranza_fecha_prog)))
            'VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
            
            'CONT2 = CONT2 + 1
            'rs_aux2!cobranza_requisito_plazo = "S"
            ''rs_aux2!cobranza_concepto_plazo = "CONFORMIDAD DEL SERVICIO"
            'If rs_aux2!cobranza_programada_bs <> 0 Then
            '    rs_aux2!Literal = Literal(CStr(rs_aux2!cobranza_programada_bs)) + " BOLIVIANOS"
            'End If
            'rs_aux2!proceso_codigo = "TEC"
            'rs_aux2!subproceso_codigo = "TEC-02"
            'rs_aux2!etapa_codigo = "TEC-02-02"
            'rs_aux2!clasif_codigo = "TEC"
            'rs_aux2!doc_codigo = "R-105"    ' R-307 Certificado de Mantenimiento ' Colocar en la conformidad
            'rs_aux2!doc_numero = "0"        'NumComp
            'rs_aux2!poa_codigo = "3.2.3"
            'rs_aux2!estado_codigo = "REG"
            'rs_aux2!usr_codigo = glusuario
            'rs_aux2!fecha_registro = Format(Date, "dd/mm/yyyy")
            'rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
            'rs_aux2!correl_ac = 0
            'rs_aux2!estado_ac = "REG"
            'rs_aux2.Update
            ' JQA 2022-10-22
            VAR_FEC2 = UCase(MonthName(Month(FControl)))
            CONT2 = CONT2 + 1
            VAR_GLOSA = "MTTO. GRATUITO POR LA GESTION: " + Str(Year(CDate(FControl))) + " - MES: " + VAR_FEC2
            VAR_SOLTIPO = Ado_datos.Recordset!solicitud_tipo
            VAR_ZPILOTO = Ado_datos.Recordset!zpiloto_codigo
            'cobranza_fecha_conformidad, correl_prog , usr_aprueba, fecha_aprueba
            db.Execute "INSERT INTO ao_ventas_cobranza_inst (venta_codigo, cobranza_prog_codigo, venta_codigo_new, ges_gestion, cobranza_fecha_prog, Gestion, cobranza_mes, fmes_plan, doc_numero_crono, edif_codigo, zpiloto_codigo, " & _
                       "  trans_codigo, estado_codigo, usr_codigo, fecha_registro, Observaciones)  " & _
                       " VALUES (" & NumComp & ", " & correldet2 & ", '0', '" & gestion0 & "', '" & FControl & "', '" & Year(FControl) & "', '" & Month(FControl) & "', '0', '0', '" & GlEdificio & "', " & VAR_ZPILOTO & ",  " & _
                       " '42', 'REG', '" & glusuario & "', '" & Date & "', '" & VAR_GLOSA & "' )"
                       
            'Asigna IdCrono (fmes_plan)
            'VAR_ZPILOTO = Ado_datos.Recordset!zpiloto_codigo
            Set rs_aux18 = New ADODB.Recordset
            If rs_aux18.State = 1 Then rs_aux18.Close
            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZPILOTO & " AND ges_gestion = '" & Year(FControl) & "' AND fmes_correl = " & Month(FControl) & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux18.RecordCount > 0 Then
                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            Else
                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            End If
            '
          Else
            db.Execute "UPDATE ao_ventas_cobranza_inst SET gestion = '" & Year(rs_aux2!cobranza_fecha_prog) & "', cobranza_mes = '" & Month(rs_aux2!cobranza_fecha_prog) & "' where  venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & ""
            db.Execute "UPDATE ao_ventas_cobranza_inst SET cobranza_prog_codigo = " & correldet2 & " where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(FControl) & "'  AND MONTH(cobranza_fecha_prog) ='" & Month(FControl) & "'  "
            
            'Asigna IdCrono (fmes_plan) WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux18 = New ADODB.Recordset
            If rs_aux18.State = 1 Then rs_aux18.Close
            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZPILOTO & " AND ges_gestion = '" & rs_aux2!gestion & "' AND fmes_correl = " & rs_aux2!cobranza_mes & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux18.RecordCount > 0 Then
                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            Else
                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            fdia = Day(FControl)
            fanio = Year(FControl)
            'CONT3 = CONT2 * CONT_MED
            CONT3 = 1
            While (CONT3 <= CONT_MED)
                fmes = Month(FControl)
                Select Case fmes
                    Case 2
                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
                            Dias_Mes = 29
                        Else
                            Dias_Mes = 28
                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
                        End If
                    Case 1, 3, 5, 7, 8, 10, 12
                        Dias_Mes = 31
                    Case 4, 6, 9, 11
                        Dias_Mes = 30
                End Select
                If Val(VAR_MES2) = Month(FControl) Then
                    rs_aux2!cobranza_fecha_prog = FControl
                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
                    rs_aux2!cobranza_fecha_cobro = FControl + 20
                    VAR_MES2 = VAR_MES2 + CONT_MED
                    If Val(VAR_MES2) > 12 Then
                        VAR_MES2 = Val(VAR_MES2) - 12
                    End If
                End If
                FControl = FControl + Dias_Mes
                CONT3 = CONT3 + 1
                CONT4 = CONT4 + Dias_Mes
            Wend
            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
            CONT2 = CONT2 + 1
          End If
        End If
      Wend
      MsgBox "El Cronograma fue Generado Exitosamente... ", , "Atención"
      'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo_verif = 'APR' Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
      If corrprog > 0 Then
        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & corrprog & "' "
        db.Execute "update ao_ventas_cabecera set venta_plazo_dias_calendario = " & CONT4 & " "
      End If

'    Else
'       MsgBox "Error Verifique la Venta de Productos..."
'    End If
End Sub


Private Sub CRONO_MTTO()
    Set rs_aux0 = New ADODB.Recordset
    If rs_aux0.State = 1 Then rs_aux0.Close
    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & GlEdificio & "'   ", db, adOpenStatic
    If rs_aux0.RecordCount > 0 Then
        VAR_EDIF = Ado_datos.Recordset!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
    End If
    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
    
    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
    ' jalar ORDEN de tc_zona_piloto_edif
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & GlEdificio & "'    ", db, adOpenStatic
    If rs_datos6.RecordCount > 0 Then
        DIA_ORDEN = rs_datos6!zona_edif_orden
    Else
        Set rs_aux18 = New ADODB.Recordset
        If rs_aux18.State = 1 Then rs_aux18.Close
        rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZPILOTO & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux18.RecordCount > 0 Then
            VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
        Else
            VAR_ORDEN = 1
        End If
        'gestion0 = "2022"
        'VAR_MED = "MES"
        'VAR_EMPRESA = 1
        'VAR_TIPO = 6
        
       db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
                  " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo) " & _
                  " VALUES (" & VAR_ZPILOTO & ", '" & GlEdificio & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
                  
        DIA_ORDEN = "1"
        
    End If
    
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "Select * from to_cronograma_diario_FINAL WHERE fmes_plan > '3545' AND edif_codigo = '" & GlEdificio & "'    ", db, adOpenStatic
    If rs_aux4.RecordCount = 0 Then
        db.Execute "UPDATE to_cronograma_diario SET bien_codigo ='', bien_orden ='0', unidad_codigo_tec ='', tec_plan_codigo ='0', edif_codigo ='', edif_descripcion ='' WHERE fmes_plan > '3545' AND edif_codigo= '" & GlEdificio & "'  "        'bien_codigo =''
    End If
    'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
    FInicio = Ado_datos.Recordset!fecha_inicio_real                       '
    MControl = Month(FInicio)               'Ado_datos.Recordset!mes_inicio_crono_tec                     'mes_inicio_crono
    'NumComp = Ado_datos.Recordset!venta_codigo
    Set rs_aux1 = New ADODB.Recordset
    'rs_aux1.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
    rs_aux1.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "    ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux1.RecordCount > 0 Then
        var_cod5 = rs_aux1.RecordCount
        rs_aux1.MoveFirst
        While Not rs_aux1.EOF
            VAR_AUX2 = rs_aux1!fmes_plan
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
            rs_aux2.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2.MoveFirst
                While Not rs_aux2.EOF
                    'JQA 23-10-2022
                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
                    Set rs_aux21 = New ADODB.Recordset
                    If rs_aux21.State = 1 Then rs_aux21.Close
                    rs_aux21.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  ", db, adOpenKeyset, adLockBatchOptimistic
                    If rs_aux21.RecordCount > 0 Then
                        db.Execute "update to_cronograma_diario set unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & GlEdificio & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "   "
                        db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                        db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
                    Else
                        Set rs_aux3 = New ADODB.Recordset
                        If rs_aux3.State = 1 Then rs_aux3.Close
                        rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux3.RecordCount > 0 Then
                            rs_aux3.MoveFirst
                            'If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
                                'db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & GlEdificio & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                'VAR_COD0 = VAR_COD0 + 1
                                'CONT3 = 1
                                db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
                                'VAR_EMES = "NADA"
                            'End If
                        Else
                            'POR SI NO TIENE fmes_plan
                        End If
                    End If
                    rs_aux2.MoveNext
                Wend
            rs_aux1.MoveNext
            End If
        Wend
    End If
End Sub

Private Sub dtc_aux2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_aux2.BoundText
    dtc_desc2.BoundText = Dtc_aux2.BoundText
    Dtc_deudor2.BoundText = Dtc_aux2.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_aux4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_aux4.BoundText
    dtc_desc4.BoundText = dtc_aux4.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    Dtc_aux2.BoundText = dtc_codigo2.BoundText
    Dtc_deudor2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    Dtc_aux2.BoundText = dtc_desc2.BoundText
    Dtc_deudor2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    dtc_aux4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub Dtc_deudor2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_deudor2.BoundText
    Dtc_aux2.BoundText = Dtc_deudor2.BoundText
    dtc_desc2.BoundText = Dtc_deudor2.BoundText
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    DataCombo2.Text = DataCombo1.BoundText
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DataCombo1.Text = DataCombo2.BoundText
End Sub

Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc11_LostFocus()
'    If dtc_codigo11.Text = "L" Or dtc_codigo11.Text = "G" Then         'Hoja de Costos - CLIENTE - Importación Directa
'        'cotiza_precio_total_dol_cli
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "Select sum(cotiza_precio_total_bs_cli) as totbs, sum(cotiza_precio_total_dol_cli) as totdl , sum(cotiza_cantidad) as cantot from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " AND estado_codigo_verif = 'APR' ", db, adOpenKeyset, adLockBatchOptimistic
'        'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
'        If rs_aux5.RecordCount > 0 Then
'            TxtMontoBs.Text = IIf(IsNull(rs_aux5!totbs), 0, rs_aux5!totbs * rs_aux5!CANTOT)
'            TxtMontoUsd.Text = IIf(IsNull(rs_aux5!totdl), 0, rs_aux5!totdl * rs_aux5!CANTOT)
'            TxtCobrado.Text = 0
'            TxtCobradoUsd.Text = 0
'            TxtBstotal.Text = CDbl(TxtMontoBs.Text)
'            TxtBstotalUsd.Text = CDbl(TxtMontoUsd.Text)
'        End If
'        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + Txt_campo2.Text
'    End If
'    If dtc_codigo11.Text = "V" Then     'Facturación Local
'        'cotiza_precio_total_dol_cge
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "Select sum(cotiza_precio_total_bs_cge) as totbs, sum(cotiza_precio_total_dol_cge) as totdl , sum(cotiza_cantidad) as cantot from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " AND estado_codigo_verif = 'APR' ", db, adOpenKeyset, adLockBatchOptimistic
'        'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
'        If rs_aux5.RecordCount > 0 Then
'            TxtMontoBs.Text = IIf(IsNull(rs_aux5!totbs), 0, rs_aux5!totbs * rs_aux5!CANTOT)
'            TxtMontoUsd.Text = IIf(IsNull(rs_aux5!totdl), 0, rs_aux5!totdl * rs_aux5!CANTOT)
'            TxtCobrado.Text = 0
'            TxtCobradoUsd.Text = 0
'            TxtBstotal.Text = CDbl(TxtMontoBs.Text)
'            TxtBstotalUsd.Text = CDbl(TxtMontoUsd.Text)
'        End If
'        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + Txt_campo2.Text
'        'TxtPlazo.Visible = True
'    End If
'    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "E" Then
'            TxtConcepto.Text = "VENTA AL CONTADO - " + Txt_campo2.Text
'            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
''        Else
''        'dtc_codigo2.Text = "VD"
''        'dtc_desc2.Text = "VENTA DIRECTA"
''        'TxtCobrado.Visible = True
''        'Label7.Visible = True
''            TxtConcepto.Text = "VENTA DIRECTA AL CLIENTE"
''            TxtPlazo.Text = 0
''            TxtPlazo.Visible = False
'    End If
End Sub

Private Sub dtccodmanejo_Click(Area As Integer)
'    DtCCodigo.BoundText = dtccodmanejo.BoundText
'    DtCDescripcion.BoundText = dtccodmanejo.BoundText
'    dtcunidadmedida.BoundText = dtccodmanejo.BoundText
'    dtccodpeso.BoundText = dtccodmanejo.BoundText
End Sub

Private Sub dtccodpeso_Click(Area As Integer)
'    DtCCodigo.BoundText = dtccodpeso.BoundText
'    DtCDescripcion.BoundText = dtccodpeso.BoundText
'    dtcunidadmedida.BoundText = dtccodpeso.BoundText
'    dtccodmanejo.BoundText = dtccodpeso.BoundText
End Sub

Private Sub dtccodpar_Click(Area As Integer)
    dtcdescripar.Text = dtccodpar.BoundText
End Sub

Private Sub dtccodpoa_Click(Area As Integer)
'    dtcdespoa.Text = dtccodpoa.BoundText
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
'    DtCCodigo.BoundText = DtCDescripcion.BoundText
'    dtcunidadmedida.BoundText = DtCDescripcion.BoundText
'    dtccodmanejo.BoundText = DtCDescripcion.BoundText
'    dtccodpeso.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub dtc_precioventabase15_Click(Area As Integer)
'    dtc_desc15.BoundText = dtc_precioventabase15.BoundText
'    dtc_unimed15.BoundText = dtc_precioventabase15.BoundText
'    dtc_stocktotal15.BoundText = dtc_precioventabase15.BoundText
'    dtc_grupo15.BoundText = dtc_precioventabase15.BoundText
'    dtc_subgrupo15.BoundText = dtc_precioventabase15.BoundText
'    Dtc_partida15.BoundText = dtc_precioventabase15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_precioventabase15.BoundText
'    dtc_codigo15.BoundText = dtc_precioventabase15.BoundText
'    dtc_preciocompra15.BoundText = dtc_precioventabase15.BoundText
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

Private Sub dtc_desc2_LostFocus()
    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
    If Dtc_deudor2.Text = "SI" Then
        Dtc_deudor2.backColor = &HFF&
    Else
        Dtc_deudor2.backColor = &H80000010
    End If
    
End Sub

Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
''    Ado_datos4.Refresh
'End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'parametro = "estado_codigo" + " = " + "'REG'"
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    VAR_UORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            Aux = "DCOMB"
            VAR_DPTO = "3"
        Case "1.7"    'Santa Cruz
            Aux = "DCOMS"
            VAR_DPTO = "7"
        Case "1.2", "1.3"    'La Paz - Comercial
            Aux = "DVTA"
            VAR_DPTO = "2"
        Case "1.8"    ' Chuquisaca
            Aux = "DCOMC"
            VAR_DPTO = "1"
        Case "1.3"    ' Modernizacion
            Aux = "DNMOD"
            VAR_DPTO = "2"
        Case "0"    ' TODO
            Aux = "DVTA"
            VAR_DPTO = "2"
     End Select
    parametro = Aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    If Ado_datos.Recordset.RecordCount > 0 Then
        nroventa = Ado_datos.Recordset!venta_codigo
    Else
        nroventa = 0
    End If
'    Call ABRIR_TABLA_DET
'    If glusuario = "ADMIN" Then
'        Command1.Visible = True
'    Else
'        Command1.Visible = False
'    End If
    'txt_codigo.Enabled = True
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    GlNombFor = "F04"
    'LblUsuario.Caption = GlUsuario
    marca1 = 1
    deta2 = 0
'    BtnImprimir2.Visible = False
'    BtnImprimir3.Visible = False
'    FrmEdita.Enabled = False
'    FrmCobros.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    'SSTab1.TabEnabled(1) = False
    'SSTab1.TabEnabled(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
    VAR_NEW = "X"
'    Chk_plazo.Value = 0
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset     'UNIDAD EJECUTORA
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_unidad_ejecutora WHERE estado_codigo= 'APR' order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    rs_datos2.Open "Select * from gc_beneficiario WHERE estado_codigo= 'APR' order by beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones WHERE estado_codigo= 'APR' order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    'Beneficiario Funcionario - Vendedor
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & Aux & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Set rs_datos4A = New ADODB.Recordset     'Beneficiario Funcionario - Cobrador
    If rs_datos4A.State = 1 Then rs_datos4A.Close
    Select Case parametro
        Case "DVTA"    'La Paz - Comercial
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DCOMB"    'Cochabamba
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DADMB' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DCOMS"    'Santa Cruz
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOMS' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DCOMC"    'Chuquisaca
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOMC' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case "DNMOD"    'Modernizacion
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DNMOD' ORDER BY beneficiario_denominacion ", db, adOpenStatic
        Case Else    ' TODO
            rs_datos4A.Open "select * from rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
     End Select
    '    rs_datos4A.Open "gp_listar_gc_beneficiario_funcionario", db, adOpenStatic
    Set ado_datos4A.Recordset = rs_datos4A
'    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
    
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If parametro = "DNMOD" Then
    '    rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo = 'C'  ", db, adOpenStatic
    'Else
        rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo = 'L' or venta_tipo = 'V' or venta_tipo = 'G' ", db, adOpenStatic
    'End If
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh
    
    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    rs_datos15.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    'rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
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
'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub valida_campos()
  If dtc_codigo1 = "" Then
    MsgBox "Debe Elejir ... " + lbl_campo1, vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  'Al Aprobar   Or dtc_codigo2 = "0"
  If dtc_codigo2 = "" Then
    MsgBox "Debe Elejir ... " + lbl_campo2, vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3 = "" Then
    MsgBox "Debe Elejir ... " + lbl_campo3, vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11 = "" Then
    MsgBox "Debe Elejir el Tipo de Venta!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir ... " + lbl_campo4, vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo2 = "" Then
    MsgBox "Debe Registrar el Cite de Trámite, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If TxtConcepto = "" Then
'    MsgBox "Debe Registrar ... " + lbl_concepto, vbExclamation, "Atención"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub grabar()
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
  'db.BeginTrans
'       db.Execute " update ao_ventas_cabecera set venta_tipo = '" & dtc_codigo11.Text & "', venta_fecha= '" & DTPfechasol.Value & "' , unidad_codigo_ant = '" & Txt_campo2.Text & "' , beneficiario_codigo_resp= '" & dtc_codigo4.Text & "', venta_descripcion='" & TxtConcepto.Text & "' , venta_monto_total_dol= " & CDbl(TxtMontoUsd.Text) & " , venta_monto_total_bs= " & CDbl(TxtMontoBs.Text) & ", estado_codigo = 'REG', usr_codigo = '" & glusuario & "', fecha_registro = '" & Format(Date, "dd/mm/yyyy") & "'  where unidad_codigo = '" & dtc_codigo1.Text & "'  and solicitud_codigo = " & txt_codigo.Caption & " "
'       If VAR_UORIGEN = "DNMOD" Then
'            db.Execute " update ao_ventas_cabecera set proceso_codigo = 'TEC', subproceso_codigo= 'TEC-05' , etapa_codigo = 'TEC-05-01' , clasif_codigo= 'TEC', doc_codigo= 'R-313' , poa_codigo= '3.2.7'  where unidad_codigo = '" & dtc_codigo1.Text & "'  and solicitud_codigo = " & txt_codigo.Caption & " "
'       Else
'            db.Execute " update ao_ventas_cabecera set proceso_codigo = 'COM', subproceso_codigo= 'COM-02' , etapa_codigo = 'COM-02-01' , clasif_codigo= 'COM', doc_codigo= 'R-223' , poa_codigo= '3.1.2'  where unidad_codigo = '" & dtc_codigo1.Text & "'  and solicitud_codigo = " & txt_codigo.Caption & " "
'       End If
       
    'db.CommitTrans
    If Ado_datos.Recordset.RecordCount > 0 Then
        NumComp = Ado_datos.Recordset!venta_codigo
       marca1 = Ado_datos.Recordset.Bookmark
       db.Execute "Update ao_ventas_alcance set fecha_inicio_real = '" & DTPfechasol.Value & "', fecha_fin_real = '" & DTPfechaFin.Value & "', doc_codigo='R-321', correl_doc=" & Val(Txt_campo1.Text) & "  WHERE venta_codigo = " & NumComp & " AND solicitud_tipo = '6' "
'       If Ado_datos.Recordset("venta_tipo") = "E" Then
'           db.Execute "INSERT INTO ao_ventas_cobranza_inst (venta_codigo, ges_gestion, beneficiario_codigo, beneficiario_codigo_resp, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, literal, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, factura_impresa, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
'           "VALUES ('" & Ado_datos.Recordset!venta_codigo & "', '" & Ado_datos.Recordset!ges_gestion & "', '" & Ado_datos.Recordset!beneficiario_codigo & "', '" & Ado_datos.Recordset!beneficiario_codigo_resp & "', " & Ado_datos.Recordset!venta_monto_total_bs & ", '" & Ado_datos.Recordset!venta_monto_total_dol & "', '0', '0', " & Ado_datos.Recordset!venta_monto_total_bs & ", " & Ado_datos.Recordset!venta_monto_total_dol & ", '" & Date & "', '" & Date & "', 'CANCELADO', 'CERO', 'COM', 'COM-02', 'COM-02-02', 'ADM', 'R-103', '0', 'R-101', '0', '0', 'N', '3.1.2', 'REG', '" & glusuario & "', '" & Date & "', '09:00')"
'           '  cobranza_codigo       'Especif. de Identidad
'       End If
''       Call OptFilGral1_Click
'       'Ado_datos.Refresh
'       'Ado_datos.Recordset.Move marca1 - 1
''        If swgrabar = 1 Then
''            Ado_datos.Refresh
''            Ado_datos.Recordset.MoveLast
''        End If
    End If
    
   Else
        MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
   End If
     
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
    Fra_Monto.Enabled = True
    Txt_modelo.Text = Txt_modelo1.Text
    Set rs_datos18 = New ADODB.Recordset
    If rs_datos18.State = 1 Then rs_datos18.Close
    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and cotiza_codigo = " & ado_datos14.Recordset!cotiza_codigo & " ", db, adOpenKeyset, adLockReadOnly
    If rs_datos18.RecordCount > 0 Then
        TxtDescuento.Text = "0"
        TxtPrecioU.Text = IIf(IsNull(rs_datos18!cotiza_precio_fob_dol), 0, rs_datos18!cotiza_precio_fob_dol)
        'TxtPrecioU.Text = IIf(IsNull(rs_datos18!cotiza_fob_seg_dol), 0, rs_datos18!cotiza_fob_seg_dol)
    End If
    'Set ado_datos17.Recordset = rs_datos18
    'ado_datos17.Refresh
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_aux13 = New ADODB.Recordset
    If rs_aux13.State = 1 Then rs_aux13.Close
    rs_aux13.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux13.RecordCount > 0 Then
        usuario2 = rs_aux13!beneficiario_codigo
        VAR_DA = rs_aux13!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario      'VAR_DA
        Case "ADMIN", "CSALINAS"
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR' AND estado_acta='REG') "
        Case "AURBINA", "CPLATA", "GSOLIZ", "DTERCEROS"
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR' AND estado_acta='REG') "
        Case Else
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR' AND estado_acta='REG') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
      Set rs_aux13 = New ADODB.Recordset
    If rs_aux13.State = 1 Then rs_aux13.Close
    rs_aux13.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux13.RecordCount > 0 Then
        usuario2 = rs_aux13!beneficiario_codigo
        VAR_DA = rs_aux13!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case glusuario      'VAR_DA
        Case "ADMIN", "CSALINAS"
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR') "
        Case "AURBINA", "CPLATA", "GSOLIZ", "DTERCEROS"
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR') "
        Case Else
            queryinicial = "select * From av_ventas_alcance WHERE (solicitud_tipo_alcance = '6' AND estado_codigo='APR') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
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
  txtTDC.Text = GlTipoCambioOficial
  
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

Private Sub sstab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        'SSTab1.TabEnabled(0) = True
        'SSTab1.TabEnabled(1) = False
    Else
'           FrmEditaDet.Visible = False
'           DtGLista.Visible = False
'           adoao_solicitud_lista.Visible = False
    End If

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

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMontoUsd_LostFocus()
    If TxtMontoUsd.Text = "" Or TxtMontoUsd.Text = "0" Or TxtMontoUsd.Text = "0.00" Then
        TxtMontoBs.Text = "0"
        TxtMontoUsd.Text = "0"
        TxtBstotalUsd = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd)
    Else
        TxtMontoBs.Text = Round(CDbl(TxtMontoUsd.Text) * GlTipoCambioMercado, 2)
    End If
    TxtBstotalUsd.Text = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd)
    TxtBstotal.Text = CDbl(TxtMontoBs) - CDbl(TxtCobrado)
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtPrecioU_LostFocus()
    If TxtPrecioU.Text = "" Or TxtPrecioU.Text = "0" Or TxtPrecioU.Text = "0.00" Then
        TxtDescuento.Text = "0"
        TxtPrecioU.Text = "0"
        TxtTotal.Text = Round(CDbl(TxtPrecioU) - CDbl(TxtDescuento), 2)
    Else
        TxtTotal.Text = Round(CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento), 2)
    End If
End Sub

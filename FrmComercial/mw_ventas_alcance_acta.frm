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
   ScaleHeight     =   5.48397e6
   ScaleMode       =   0  'User
   ScaleWidth      =   3.9812e8
   WindowState     =   2  'Maximized
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
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18000
         Picture         =   "mw_ventas_alcance_acta.frx":058A
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
         Picture         =   "mw_ventas_alcance_acta.frx":0D4C
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
         Picture         =   "mw_ventas_alcance_acta.frx":1501
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
         Picture         =   "mw_ventas_alcance_acta.frx":1D34
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   66
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -15
         Picture         =   "mw_ventas_alcance_acta.frx":2480
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
         Picture         =   "mw_ventas_alcance_acta.frx":2D95
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
         Picture         =   "mw_ventas_alcance_acta.frx":3662
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
         Picture         =   "mw_ventas_alcance_acta.frx":3F4E
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
      Height          =   2220
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   9
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   165
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Codificar"
         Height          =   640
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Codifica Equipos"
         Top             =   195
         Visible         =   0   'False
         Width           =   765
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
      TabPicture(0)   =   "mw_ventas_alcance_acta.frx":4724
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
            Bindings        =   "mw_ventas_alcance_acta.frx":4740
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
               Format          =   119406593
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
               Format          =   119406593
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
               Bindings        =   "mw_ventas_alcance_acta.frx":4759
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
               Bindings        =   "mw_ventas_alcance_acta.frx":4773
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
            Bindings        =   "mw_ventas_alcance_acta.frx":478D
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
            Bindings        =   "mw_ventas_alcance_acta.frx":47A6
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
            Bindings        =   "mw_ventas_alcance_acta.frx":47BF
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
            Bindings        =   "mw_ventas_alcance_acta.frx":47D8
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
            Bindings        =   "mw_ventas_alcance_acta.frx":47F1
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
            Bindings        =   "mw_ventas_alcance_acta.frx":480A
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
      Height          =   2295
      Left            =   2160
      TabIndex        =   7
      Top             =   5580
      Width           =   14895
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "mw_ventas_alcance_acta.frx":4823
         Height          =   1905
         Left            =   240
         TabIndex        =   8
         Top             =   225
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   3360
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
Dim rs_datos16 As New ADODB.Recordset   'ao_ventas_cobranza_prog    - Ventas cobranzas Prog
Dim rs_datos17 As New ADODB.Recordset   'ac_bienes_grupo
Dim rs_datos18 As New ADODB.Recordset   'ao_solicitud_cotiza_venta
Dim rs_datos19 As New ADODB.Recordset   'ao_ventas_cobranza_prog    - Acumula Cobranzas Prog
Dim rs_datos20 As New ADODB.Recordset   'ao_solicitud_costos    - Acumula Costos

'AUXILIARES
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
Dim VAR_CODANT, Var_Comp, VAR_SOL, VAR_TIPOS As Integer
Dim VAR_NUM As Integer
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
Dim VAR_JQ, VAR_VAL As String
    
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
            BtnModDetalle1.Visible = True
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
'            FraDet2.Visible = False
'            If (Ado_datos.Recordset!unidad_codigo = "DVTA") Then
'                BtnAddDetalle.Visible = True
'            Else
'                BtnAddDetalle.Visible = False
'            End If
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
'            ' por si es nuevo
'            dtccodpoa.Text = " "
'            dtcdespoa.Text = dtccodpoa.BoundText
'            dtc_codigo4.Text = " "
'            Dtcpaternosol.Text = dtc_codigo4.BoundText
'            dtcmaternosol.Text = " "
'            dtcnombresol.Text = " "
'            dtccodpuesto.Text = " "
'            dtcdenopuesto.Text = dtccodpuesto.BoundText
'            dtccoduni.Text = " "
'            dtcdescripuni.Text = dtccoduni.BoundText
'            dtc_codigo15.Text = " "
'            dtc_desc15.Text = " "
'            TxtMonto_bolivianos.Text = 0
'            Txtobservaciones.Text = ""
'            Txtcaracteristicas.Text = ""
'            txtsolpeso.Text = 0
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
 If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
        'BtnModDetalle2.Visible = False
        If (Ado_datos16.Recordset("estado_codigo") = "REG") Then
'            If (Ado_datos.Recordset("estado_codigo") = "APR") Then
'                BtnAprobar2.Visible = False
'            Else
'                BtnAprobar2.Visible = True
'            End If
            BtnImprimir2.Visible = True
            BtnAprobar2.Visible = True
            BtnAnlDetalle2.Visible = True
            BtnModDetalle2.Visible = True
        End If
        If (Ado_datos16.Recordset("estado_codigo") = "APR") Then
            BtnImprimir2.Visible = True
            BtnAprobar2.Visible = False
            BtnAnlDetalle2.Visible = False
            BtnModDetalle2.Visible = False
        End If
        If (Ado_datos16.Recordset("estado_codigo") = "ANL") Then
            BtnImprimir2.Visible = False
            BtnAnlDetalle2.Visible = False
            BtnModDetalle2.Visible = False
            BtnAprobar2.Visible = False
        End If
    Else
        BtnAprobar2.Visible = False
        BtnImprimir2.Visible = False
        BtnAnlDetalle2.Visible = False
        BtnModDetalle2.Visible = False
    End If
 Else
    BtnAprobar2.Visible = False
    BtnImprimir2.Visible = False
    BtnAnlDetalle2.Visible = False
    BtnModDetalle2.Visible = False
 End If
End Sub

Private Sub BtnAddDetalle_Click()
  'marca1 = Ado_datos.Recordset.Bookmark
  If ado_datos14.Recordset!estado_codigo = "REG" Then
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "select * from fc_partida_gasto where par_codigo = '43340' ", db, adOpenKeyset, adLockReadOnly
    If rs_aux6.RecordCount > 0 Then
        VAR_OA = "AO36" + LTrim(Str(rs_aux6!correlativo36 + 1))
        Set rs_aux7 = New ADODB.Recordset
        If rs_aux7.State = 1 Then rs_aux7.Close
        rs_aux7.Open "select * from ac_bienes where bien_codigo = '" & VAR_OA & "' ", db, adOpenKeyset, adLockReadOnly
        If rs_aux7.RecordCount > 0 Then
            MsgBox "El equipo " + VAR_OA + " YA Existe, vuelva a intentar !! ", vbExclamation, "Atención!"
            db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
        Else
            ado_datos14.Recordset!bien_codigo = Trim(VAR_OA)
            db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
            db.Execute "insert into ac_bienes(grupo_codigo, subgrupo_codigo, bien_codigo, par_codigo, bien_descripcion, bien_precio_compra, bien_precio_venta_base, bien_precio_venta_final, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo, bien_stock_minimo, bien_stock_inicial, bien_stock_ingreso, bien_stock_salida, bien_stock_actual, bien_total_compra_bs, bien_total_venta_bs, bien_utilidad_Bs, bien_codigo_anterior, bien_codigo_universal, bien_descripcion_anterior, pais_codigo, archivo_foto2, archivo_foto, estado_codigo, fecha_registro, usr_codigo) " & _
            "VALUES ('40000', '43000', '" & VAR_OA & "', '43340', 'CAPACIDAD ' + '" & dtc_desc31.Text & "' + ' PERSONAS Y VELOCIDAD ' + '" & dtc_valor41.Text & "' + ' m/s', " & var_cod & ", '0', '0', 'EQP', 'EQP', '1', 'S/M', '1', '0', '0', '0', '0', '0', '0', '0', '-', '-', '-', 'NN', '" & VAR_COD3 & "' + '2.JPG', '" & VAR_COD3 & "' + '.JPG', 'REG', '" & Date & "', '" & glusuario & "') "
        End If
    End If
'    'If OptFilGral1.Value = True Then Call OptFilGral1_Click
'    'If OptFilGral2.Value = True Then Call OptFilGral2_Click
''    Ado_datos.Recordset.Move marca1 - 1
'    swnuevo = 1
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    SSTab1.TabEnabled(2) = False
'    FrmEdita.Visible = True
'    FrmEdita.Enabled = True
'    FraNavega.Enabled = False
'    FrmDetalle.Enabled = False
'    FrmCobranza.Visible = False
'    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
'    'tipo Beneficiario
'    Set rs_datos12 = New ADODB.Recordset
'    If rs_datos12.State = 1 Then rs_datos12.Close
'    'rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
'    rs_datos12.Open "select * from gc_tipo_beneficiario where tipoben_codigo = '" & Dtc_aux2.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set Ado_Datos12.Recordset = rs_datos12
'    Ado_Datos12.Refresh
'
'    ado_datos14.Recordset.AddNew
  Else
    MsgBox "El registro Aprobado o Anulado, NO pueden ser modificado !! ", vbExclamation, "Atención!"
  End If
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
    rs_aux10.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & dtc_aux3.Text & "' and estado_codigo = 'APR' ", db, adOpenStatic
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
        VAR_SOLA = Ado_datos.Recordset!venta_codigo
       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
       If sino = vbYes Then
           db.Execute "Update ao_ventas_alcance set estado_acta ='APR'  WHERE venta_codigo = " & VAR_SOLA & " AND solicitud_tipo = '6' "
           'ASIGNA A VARIABLES CAMPOS CLAVES
'           'INI Deptos de Bolivia
'            Select Case VAR_DPTO
'                 Case "1"
'                     VAR_DPTOD = "CHUQUISACA"
'                 Case "2"
'                     VAR_DPTOD = "LA PAZ"
'                 Case "3"
'                     VAR_DPTOD = "COCHABAMBA"
'                 Case "4"
'                     VAR_DPTOD = "ORURO"
'                 Case "5"
'                     VAR_DPTOD = "POTOSI"
'                 Case "6"
'                     VAR_DPTOD = "TARIJA"
'                 Case "7"
'                     VAR_DPTOD = "SANTA CRUZ"
'                 Case "8"
'                     VAR_DPTOD = "BENI"
'                 Case "9"
'                     VAR_DPTOD = "PANDO"
'            End Select
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
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_codigo") = "REG" Then
      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
          'Dim rstdestino As New ADODB.Recordset
          'Set rstdestino = New ADODB.Recordset
          'If rstdestino.State = 1 Then rstdestino.Close
          'rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  ", db, adOpenDynamic, adLockOptimistic
          'If Not rstdestino.BOF Then rstdestino.MoveFirst
          'If Not rstdestino.BOF And Not rstdestino.EOF Then
          '    rstdestino("estado_codigo") = "E"
          '    rstdestino.Update
          'End If
          'If rstdestino.State = 1 Then rstdestino.Close
          marca1 = Ado_datos.Recordset.Bookmark
          'Ado_datos.Recordset.Requery
          'Ado_datos.Refresh
          Call OptFilGral1_Click
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede ANULAR el registro que ya fue Aprobado o previamente Anulado.", , "Atencion"
    End If
  Else
    MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnGrabar_Click()
VAR_SOLA = Ado_datos.Recordset!venta_codigo
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
  If txt_campo2.Text = "" And txt_campo2.Text = " " Then
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
        rs_datos.Find "venta_codigo = " & VAR_SOLA & "   ", , , 1
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

Private Sub Chk_plazo_Click()
    If Chk_plazo.Value = 1 Then
        lbl_plazo.Visible = True
        txt_plazo.Visible = True
        
    Else
        lbl_plazo.Visible = False
        txt_plazo.Visible = False
    End If
End Sub

'Private Sub Contabiliza_venta()
'    Call graba_proyecto
'    Call graba_ingreso
'  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
'  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
'  'If sino = vbYes Then
'    ' INI CORRECCION 18-JUN-2014
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    Select Case VAR_CODTIPO
'        Case "DEI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
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
'        Case "DEY"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEY' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'        Case "REC"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "DES"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "ANI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'        Case "DVI"
'            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'            If rstdestino.RecordCount > 0 Then
'                j = rstdestino.RecordCount
'            Else
'              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'              Exit Sub
'            End If
'
'            '' 02/07/2014 VERIFICAR
'            'If rstdestino.State = 1 Then rstdestino.Close
'            'rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'            'If rstdestino2.State = 1 Then rstdestino2.Close
'            'rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "REC" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DYR" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DES" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "ANI" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DVI" Then
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
''        rs_aux2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
''        If rstdestino2.State = 1 Then rstdestino2.Close
''        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
'    fte_codigo1 = VAR_FTE
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
'        yacontabilizo = 1
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
'
'        'R-112, R-110, R-111
'          Set rs_aux14 = New ADODB.Recordset
'          SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = 'R-112' "          '  '" & txt_codigo1 & "' "
'          rs_aux14.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'          If rs_aux14.RecordCount > 0 Then
'                rs_aux14!correl_doc = rs_aux14!correl_doc + 1
'                VAR_COMPM = rs_aux14!correl_doc
'                rs_aux14.Update
'          End If
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
'      rs_aux2("venta_compra") = correlv
'      If yacontabilizo = 0 Then
'        rs_aux2("Fecha_transacion") = Date
'      End If
'      rs_aux2("mes_trasaccion") = UCase(MonthName(Month(Date)))
'      rs_aux2("ges_gestion") = Year(Date)     'glGestion
'      rs_aux2("beneficiario_codigo") = VAR_BENEF
'      rs_aux2("glosa") = "INGRESO POR: " + VAR_GLOSA
'      rs_aux2("unidad_codigo") = VAR_COD4       'Ado_datos.Recordset("unidad_codigo")
'      rs_aux2("solicitud_codigo") = VAR_SOL     'Ado_datos.Recordset("solicitud_codigo")
'      rs_aux2("tipo_moneda") = VAR_MONEDA
'      rs_aux2("unidad_codigo_ant") = VAR_CITE
'
'      rs_aux2("proceso_codigo") = "FIN"
'      rs_aux2("subproceso_codigo") = "FIN-02"
'      Select Case VAR_CODTIPO
'        Case "DEI"
'            rs_aux2("etapa_codigo") = "FIN-02-01"
'        Case "DEY"
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
'      rs_aux2("doc_codigo") = "R-112"
'      rs_aux2("doc_numero") = VAR_COMPM         'Var_Comp
'      rs_aux2("pro_codigo_det") = VAR_PROY2
'
'      rs_aux2("estado_codigo") = "APR"
'
'      If yacontabilizo = 0 Then
'        rs_aux2("usr_codigo") = glusuario
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
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
''      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
''      End If
'
''      If v_Tipo_Comp(1, i) = "DVI" Then
''        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
''        If rstdestino.State = 1 Then rstdestino.Close
''        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
''        If rstdestino2.State = 1 Then rstdestino2.Close
''        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
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
'      '2115
'      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
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
'        VAR_DCORR = rs_aux1("correl")
'      End If
'      If rs_aux1.State = 1 Then rs_aux1.Close
'      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
'      If rs_aux1.RecordCount > 0 Then
'        h_cta_nombre_1 = rs_aux1("NombreCta")
'        h_aux1_1 = rs_aux1("aux1")
'        h_aux2_1 = rs_aux1("aux2")
'        h_aux3_1 = rs_aux1("aux3")
'        VAR_HCORR = rs_aux1("correl")
'      End If
'      If rs_aux1.State = 1 Then rs_aux1.Close
'      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
'      If rs_aux1.RecordCount > 0 Then
'        VAR_NOMD = rs_aux1("NombreCta")
'      End If
'      If rs_aux1.State = 1 Then rs_aux1.Close
'      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and nivel = '4' ", db, adOpenKeyset, adLockReadOnly
'      If rs_aux1.RecordCount > 0 Then
'        VAR_NOMH = rs_aux1("NombreCta")
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
'      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "DEY") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
'        rstdestino2("D_Cuenta") = cta_deb1
'        rstdestino2("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_deb11
'        rstdestino2("D_SubCta2") = Subcta_deb21
'        rstdestino2("D_Aux1") = d_aux1_1
'        rstdestino2("D_Aux2") = d_aux2_1
'        rstdestino2("D_Aux3") = d_aux3_1
'        rstdestino2("NOMCTADEBE") = VAR_NOMD
'        rstdestino2("D_Correl") = VAR_DCORR
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
'                rstdestino2("D_Des_Aux1") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux1") = VAR_CTA
'                rstdestino2("D_Des_Aux1") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux1") = VAR_PROY2
'                rstdestino2("D_Des_Aux1") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux1") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux1") = VAR_DPTO
'                rstdestino2("D_Des_Aux1") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux1") = VAR_ORG
'                rstdestino2("D_Des_Aux1") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'        End Select
'
'        Select Case d_aux2_1
'            Case "01"
'                rstdestino2("D_Cta_Aux2") = VAR_BENEF
'                rstdestino2("D_Des_Aux2") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux2") = VAR_CTA
'                rstdestino2("D_Des_Aux2") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux2") = VAR_PROY2
'                rstdestino2("D_Des_Aux2") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux2") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux2") = VAR_DPTO
'                rstdestino2("D_Des_Aux2") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux2") = VAR_ORG
'                rstdestino2("D_Des_Aux2") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'        End Select
'
'        Select Case d_aux3_1
'            Case "01"
'                rstdestino2("D_Cta_Aux3") = VAR_BENEF
'                rstdestino2("D_Des_Aux3") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux3") = VAR_CTA
'                rstdestino2("D_Des_Aux3") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux3") = VAR_PROY2
'                rstdestino2("D_Des_Aux3") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux3") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux3") = VAR_DPTO
'                rstdestino2("D_Des_Aux3") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux3") = VAR_ORG
'                rstdestino2("D_Des_Aux3") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'        End Select
''        If d_aux1_1 = "01" Then
''          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''        End If
''        If d_aux1_1 = "02" Then
''          rstdestino2("D_Cta_Aux1") = VAR_CTA
''        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        ' CORREGIR MONTOS JQA 2014-JUL-08
'        If j > 1 Then
'            If i = 1 Then
'                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
'                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
'            Else
'                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
'                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
'            End If
'        Else
'            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
'            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
'        End If
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
'        rstdestino2("NOMCTAHABER") = VAR_NOMH
'        rstdestino2("h_Correl") = VAR_HCORR
'        'rstdestino2("H_Cta_Aux1") = ""
'        Select Case h_aux1_1
'            Case "01"
'                rstdestino2("H_Cta_Aux1") = VAR_BENEF
'                rstdestino2("H_Des_Aux1") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux1") = VAR_CTA
'                rstdestino2("H_Des_Aux1") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux1") = VAR_PROY2
'                rstdestino2("H_Des_Aux1") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux1") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux1") = VAR_DPTO
'                rstdestino2("H_Des_Aux1") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux1") = VAR_ORG
'                rstdestino2("H_Des_Aux1") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'        End Select
'
'        Select Case h_aux2_1
'            Case "01"
'                rstdestino2("H_Cta_Aux2") = VAR_BENEF
'                rstdestino2("H_Des_Aux2") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux2") = VAR_CTA
'                rstdestino2("H_Des_Aux2") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux2") = VAR_PROY2
'                rstdestino2("H_Des_Aux2") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux2") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux2") = VAR_DPTO
'                rstdestino2("H_Des_Aux2") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux2") = VAR_ORG
'                rstdestino2("H_Des_Aux2") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'        End Select
'
'        Select Case h_aux3_1
'            Case "01"
'                rstdestino2("H_Cta_Aux3") = VAR_BENEF
'                rstdestino2("H_Des_Aux3") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux3") = VAR_CTA
'                rstdestino2("H_Des_Aux3") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux3") = VAR_PROY2
'                rstdestino2("H_Des_Aux3") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux3") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux3") = VAR_DPTO
'                rstdestino2("H_Des_Aux3") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux3") = VAR_ORG
'                rstdestino2("H_Des_Aux3") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
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
'        If j > 1 Then
'            If i = 1 Then
'                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
'                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
'            Else
'                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
'                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
'            End If
'        Else
'            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
'            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
'        End If
'        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
'      End If
'
'      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
'      If (VAR_CODTIPO = "DES") Or (VAR_CODTIPO = "ANI") Or (VAR_CODTIPO = "DVI") Then
'        'desafecta un devengado
'        rstdestino2("D_Cuenta") = cta_credito1
'        rstdestino2("D_Nombre") = RTrim(h_cta_nombre_1) ' CAMPO PARA ELIMINAR
'        rstdestino2("D_Subcta1") = Subcta_cred11
'        rstdestino2("D_SubCta2") = Subcta_cred21
'        rstdestino2("D_Aux1") = h_aux1_1
'        rstdestino2("D_Aux2") = h_aux2_1
'        rstdestino2("D_Aux3") = h_aux3_1
''        rstdestino2("D_Cta_Aux1") = "VESCT"
'        Select Case h_aux1_1
'            Case "01"
'                rstdestino2("D_Cta_Aux1") = VAR_BENEF
'                rstdestino2("D_Des_Aux1") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux1") = VAR_CTA
'                rstdestino2("D_Des_Aux1") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux1") = VAR_PROY2
'                rstdestino2("D_Des_Aux1") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux1") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux1") = VAR_DPTO
'                rstdestino2("D_Des_Aux1") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux1") = VAR_ORG
'                rstdestino2("D_Des_Aux1") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux1") = ""
'                rstdestino2("D_Des_Aux1") = ""
'        End Select
'
'        Select Case h_aux2_1
'            Case "01"
'                rstdestino2("D_Cta_Aux2") = VAR_BENEF
'                rstdestino2("D_Des_Aux2") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux2") = VAR_CTA
'                rstdestino2("D_Des_Aux2") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux2") = VAR_PROY2
'                rstdestino2("D_Des_Aux2") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux2") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux2") = VAR_DPTO
'                rstdestino2("D_Des_Aux2") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux2") = VAR_ORG
'                rstdestino2("D_Des_Aux2") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux2") = ""
'                rstdestino2("D_Des_Aux2") = ""
'        End Select
'
'        Select Case h_aux3_1
'            Case "01"
'                rstdestino2("D_Cta_Aux3") = VAR_BENEF
'                rstdestino2("D_Des_Aux3") = VAR_BEND
'            Case "02"
'                rstdestino2("D_Cta_Aux3") = VAR_CTA
'                rstdestino2("D_Des_Aux3") = VAR_CTAD
'            Case "03"
'                rstdestino2("D_Cta_Aux3") = VAR_PROY2
'                rstdestino2("D_Des_Aux3") = VAR_EDIFD
'            Case "04"
'                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("D_Des_Aux3") = VAR_UNID
'            Case "05"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "06"
'                rstdestino2("D_Cta_Aux3") = VAR_DPTO
'                rstdestino2("D_Des_Aux3") = VAR_DPTOD
'            Case "07"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "08"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "09"
'                rstdestino2("D_Cta_Aux3") = VAR_ORG
'                rstdestino2("D_Des_Aux3") = VAR_ORGD
'            Case "10"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "11"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "12"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'            Case "00"
'                rstdestino2("D_Cta_Aux3") = ""
'                rstdestino2("D_Des_Aux3") = ""
'        End Select
''        If h_aux1_1 = "01" Then
''          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''        End If
''        If h_aux1_1 = "02" Then
''          rstdestino2("D_Cta_Aux1") = VAR_CTA
''        End If
''        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        If i = 1 Then
'            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
'            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
'        Else
'            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
'            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
'        End If
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
'                rstdestino2("H_Des_Aux1") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux1") = VAR_CTA
'                rstdestino2("H_Des_Aux1") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux1") = VAR_PROY2
'                rstdestino2("H_Des_Aux1") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux1") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux1") = VAR_DPTO
'                rstdestino2("H_Des_Aux1") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux1") = VAR_ORG
'                rstdestino2("H_Des_Aux1") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux1") = ""
'                rstdestino2("H_Des_Aux1") = ""
'        End Select
'
'        Select Case d_aux2_1
'            Case "01"
'                rstdestino2("H_Cta_Aux2") = VAR_BENEF
'                rstdestino2("H_Des_Aux2") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux2") = VAR_CTA
'                rstdestino2("H_Des_Aux2") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux2") = VAR_PROY2
'                rstdestino2("H_Des_Aux2") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux2") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux2") = VAR_DPTO
'                rstdestino2("H_Des_Aux2") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux2") = VAR_ORG
'                rstdestino2("H_Des_Aux2") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux2") = ""
'                rstdestino2("H_Des_Aux2") = ""
'        End Select
'
'        Select Case d_aux3_1
'            Case "01"
'                rstdestino2("H_Cta_Aux3") = VAR_BENEF
'                rstdestino2("H_Des_Aux3") = VAR_BEND
'            Case "02"
'                rstdestino2("H_Cta_Aux3") = VAR_CTA
'                rstdestino2("H_Des_Aux3") = VAR_CTAD
'            Case "03"
'                rstdestino2("H_Cta_Aux3") = VAR_PROY2
'                rstdestino2("H_Des_Aux3") = VAR_EDIFD
'            Case "04"
'                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
'                rstdestino2("H_Des_Aux3") = VAR_UNID
'            Case "05"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "06"
'                rstdestino2("H_Cta_Aux3") = VAR_DPTO
'                rstdestino2("H_Des_Aux3") = VAR_DPTOD
'            Case "07"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "08"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "09"
'                rstdestino2("H_Cta_Aux3") = VAR_ORG
'                rstdestino2("H_Des_Aux3") = VAR_ORGD
'            Case "10"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "11"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "12"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'            Case "00"
'                rstdestino2("H_Cta_Aux3") = ""
'                rstdestino2("H_Des_Aux3") = ""
'        End Select
''        If d_aux1_1 = "01" Then
''          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
''          'DtCCta_descripcion_larga
''        End If
''        If d_aux1_1 = "02" Then
''          rstdestino2("H_Cta_Aux1") = VAR_CTA
''        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        If i = 1 Then
'            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
'            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
'        Else
'            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
'            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
'        End If
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
'        rstdestino2("Usr_codigo") = glusuario
'        rstdestino2("Fecha_registro") = Date
'        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
'      End If
'
'      rstdestino2.Update
'      If rstdestino2.State = 1 Then rstdestino2.Close
'      '======= fin registra co_diario ==========
'      rstdestino.MoveNext
'    Next i
'      '-Actualiza SubTitulo Debe
'      db.Execute "UPDATE co_diario SET co_diario.NOMCTADEBE = ltrim(cv_diario_subtitulo_debe.NombreCta) FROM co_diario INNER JOIN cv_diario_subtitulo_debe on co_diario.D_Cuenta = cv_diario_subtitulo_debe.Cuenta where co_diario.Cod_Comp = " & Var_Comp & " "
'      '--Actualiza SubTitulo Haber
'      db.Execute "UPDATE co_diario SET co_diario.NOMCTAHABER = ltrim(cv_diario_subtitulo_haber.NombreCta) FROM co_diario INNER JOIN cv_diario_subtitulo_haber on co_diario.H_Cuenta = cv_diario_subtitulo_haber.Cuenta where co_diario.Cod_Comp = " & Var_Comp & " "
'      '--Actualiza D_Nombre Debe
'      db.Execute "UPDATE co_diario SET co_diario.D_Nombre  = ltrim(cc_plan_cuentas.NombreCta) FROM co_diario INNER JOIN cc_plan_cuentas on co_diario.D_Cuenta = cc_plan_cuentas.Cuenta and co_diario.D_Subcta1 = cc_plan_cuentas.SubCta1 and co_diario.D_SubCta2 = cc_plan_cuentas.SubCta2 where co_diario.Cod_Comp = " & Var_Comp & " "
'      '--Actualiza H_Nombre Haber
'      db.Execute "UPDATE co_diario SET co_diario.H_Nombre  = ltrim(cc_plan_cuentas.NombreCta) FROM co_diario INNER JOIN cc_plan_cuentas on co_diario.H_Cuenta  = cc_plan_cuentas.Cuenta and co_diario.H_Subcta1  = cc_plan_cuentas.SubCta1 and co_diario.H_SubCta2  = cc_plan_cuentas.SubCta2 where co_diario.Cod_Comp = " & Var_Comp & " "
'
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
'        If rstdestino("codigo_tipo") = "DEI" Or (VAR_CODTIPO = "DEY") Then
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

'Private Sub graba_proyecto()
'    Select Case VAR_COD4
'        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
'            VAR_PROY = 12
'        Case "GCOM"
'            VAR_PROY = 17
'        Case "DVTA", "DCOMB", "DCOMS", "DCOMC"
'            VAR_PROY = 18
'
'    End Select
'
'    Set rs_aux1 = New ADODB.Recordset
'    If rs_aux1.State = 1 Then rs_aux1.Close
'    'SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
'    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & VAR_PROY2 & "' "
'    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & RTrim(dtc_desc3.Text) & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & VAR_PROY2 & "' "
'    Else
'        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
'           "VALUES (" & VAR_PROY & ", '" & VAR_PROY2 & "', '" & RTrim(dtc_desc3.Text) & "', '" & VAR_COD4 & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
'    End If
'End Sub

'Private Sub graba_ingreso()
'    '======= Ini grabado de datos
'   'swgraba = 0
'   'Call valida
'   'VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'   Select Case VAR_COD4
'        Case "DVTA", "DCOMB", "DCOMS", "DCOMC"             'INI COMERCIAL
'            VAR_ORG = "111"
'            VAR_TIPOS = 3
'            VAR_PARTIDA = "11310"
'        Case "COMEX"            'INI COMEX
'            VAR_ORG = "111"
'            VAR_PARTIDA = "11310"
'        Case "DNINS"            'INI INSTALACIONES
'            VAR_ORG = "111"
'            VAR_TIPOS = 4
'            VAR_PARTIDA = "11350"
'        Case "DNAJS"            'INI AJUSTE
'            VAR_ORG = "113"
'            VAR_TIPOS = 5
'            VAR_PARTIDA = "11350"
'        Case "DNMAN"            'INI MANTENIMIENTO
'            VAR_ORG = "112"
'            VAR_TIPOS = 6
'            VAR_PARTIDA = "11320"
'        Case "DNREP"            'INI REPARACIONES
'            VAR_ORG = "113"
'            VAR_TIPOS = 7
'            VAR_PARTIDA = "11330"
'        Case "DNMOD"            'INI MODERNIZACION
'            VAR_ORG = "114"
'            VAR_TIPOS = 9
'            VAR_PARTIDA = "11340"
'        Case "DNEME"            'INI EMERGENCIAS
'            VAR_ORG = "113"
'            VAR_TIPOS = 8
'            VAR_PARTIDA = "11330"
'        Case Else               'INI COMPRAS
'            VAR_ORG = "311"
'            VAR_TIPOS = 10
'            VAR_PARTIDA = "11330"
'   End Select
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
'         rstdestino.Open "select * from fo_ingresos_cabecera order by org_codigo, ingreso_codigo   ", db, adOpenDynamic, adLockOptimistic
'         rstdestino.AddNew
'         rstdestino("Ges_Gestion") = glGestion      'Year(Date)     'Ado_datos.Recordset("ges_gestion")
'         rstdestino("ingreso_codigo") = correlativo1
'         VAR_CODANT = correlativo1
'         'CAMBIAR org_codigo
'         rstdestino("org_codigo") = VAR_ORG
'         'CAMBIAR org_codigo
'         'CAMBIAR COD ingreso_codigo_anterior
'         rstdestino("ingreso_codigo_anterior") = correlativo1
'         'CAMBIAR COD ingreso_codigo_anterior
'         rstdestino("proceso_codigo") = "FIN"
'         rstdestino("subproceso_codigo") = "FIN-01"
'         rstdestino("etapa_codigo") = "FIN-01-01"
'         rstdestino("clasif_codigo") = "ADM"
'         rstdestino("doc_codigo") = "R-110"
'         rstdestino("doc_numero") = correlativo1
'         rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos.Recordset("unidad_codigo")
'         rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos.Recordset("solicitud_codigo")
'         rstdestino("solicitud_tipo") = VAR_TIPOS   '"3"
'
'         rstdestino("beneficiario_codigo") = VAR_BENEF  ' Ado_datos.Recordset("beneficiario_codigo")
''         VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
'         rstdestino("fecha_ingreso") = Date
'         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
'         rstdestino("tipo_moneda") = "BOB"
'         VAR_MONEDA = "BOB"
'         'VAR_GLOSA = Ado_datos.Recordset("venta_descripcion")
'         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA
'         If Ado_datos.Recordset("venta_tipo") = "E" Then
'            VAR_CODTIPO = "DYR"
'         Else
'            If VAR_TIPOV = "V" Then
'                VAR_CODTIPO = "DEI"
'            Else
'                VAR_CODTIPO = "DEI"
'            End If
'            'AQUUIIII       DEY         WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'         End If
'         'CAMBIAR DEI O REC
'         rstdestino("Codigo_tipo") = VAR_CODTIPO
'         rstdestino("tipo_comp") = VAR_CODTIPO
'         'CAMBIAR DEI O REC
'         'INI FTE
'         Select Case VAR_ORG
'             Case "111"              'INI SERVICIOS DE PROVISION E INSTALACION
'                 VAR_FTE = "10"
'             Case "112"            'INI SERVICIO DE MANTENIMIENTO - MANTENIMIENTO PREVENTIVO
'                 VAR_FTE = "10"
'             Case "113"            'INI SERVICIO DE REPARACIONES - MANTENIMIENTO CORRECTIVO
'                 VAR_FTE = "10"
'             Case "114"            'INI SERVICIO DE MODERNIZACION
'                 VAR_FTE = "10"
'             Case "211"            'INI APORTES DE CAPITAL
'                 VAR_FTE = "20"
'             Case "311"            'INI BANCO MERCANTIL SANTA CRUZ
'                 VAR_FTE = "30"
'             Case "312"            'INI BANCO DE CREDITO
'                 VAR_FTE = "30"
'             Case "411"            'INI AMT - REPOSICION DE PIEZAS Y PARTES
'                 VAR_FTE = "40"
'             Case Else               'INI OTROS
'                 VAR_FTE = "10"
'         End Select
'         rstdestino("fte_codigo") = VAR_FTE
'         'FIN FTE
'         'CAMBIAR RUBROS
'         rstdestino("rubro_codigo") = VAR_PARTIDA       '"11200"
'         'VAR_PARTIDA = "11200"
'         'CAMBIAR RUBROS
'         rstdestino("cheque_o_trf") = ""
'         rstdestino("Bco_codigo") = "NN"
'         'CAMBIAR CTA
'         rstdestino("cta_codigo") = "NN"
'         VAR_CTA = "NN"
'         'CAMBIAR CTA
'         rstdestino("numero_documento") = "0"
'         rstdestino("unidad_codigo_ant") = VAR_CITE     ' Ado_datos.Recordset("unidad_codigo_ant")
'         'VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
'         rstdestino("monto_dolares") = VAR_DOL2 'Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         'VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'         rstdestino("monto_bolivianos") = VAR_BS2   'Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         'VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'         rstdestino("monto_recaudado_dolares") = 0
'         rstdestino("monto_recaudado_bolivianos") = 0
'         rstdestino("convenio_codigo") = "NN"
'         rstdestino("pro_codigo_det") = VAR_PROY2   'Ado_datos.Recordset("edif_codigo")
'         'VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
'         rstdestino("estado_CODIGO") = "APR"
'         'rstdestino("estado_codigo_dr") = "DEI"
'
'         rstdestino("usr_CODIGO") = glusuario
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

'Private Sub add_correl()
'  'FALTAAAAA!! org_codigo JQA 2014-07-10
'  Set rstcorrel_ing = New ADODB.Recordset
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
'  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
'  If rstcorrel_ing.RecordCount = 0 Then
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = VAR_ORG
'     rstcorrel_ing("ges_gestion") = glGestion       'Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
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


Private Sub BtnAnlDetalle_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
'     ado_datos14.Recordset.Delete
'     ado_datos14.Recordset.Update
'     rs_datos14.Requery
'     ado_datos14.Refresh
'     'cerea
'     ado_datos14.Refresh
      db.Execute "update ao_ventas_detalle set ao_ventas_detalle.estado_codigo = 'ANL' Where ao_ventas_detalle.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_detalle.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_detalle.venta_codigo_det = " & ado_datos14.Recordset("venta_codigo_det") & " "
   End If
  Else
    MsgBox "Los Bienes del registro Aprobado o Anulado, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    FraNavega.Enabled = False
    FrmDetalle.Enabled = False
    FrmCobranza.Visible = False
    FrmAlcance.Visible = False
    swgrabar = 0
    swnuevo = 2
    'marca1 = Ado_datos.Recordset.Bookmark
    'txt_descripcion_venta.Enabled = True
    correlv = Ado_datos.Recordset!venta_codigo
    TxtNroVenta.Text = correlv  'Ado_datos.Recordset!venta_codigo  'txt_venta.Text
    TxtNroVenta.Enabled = False
    'lbltipoVenta.Caption = dtc_desc11.Text
'    lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = False
    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    FrmABMDet.Visible = False
    FrmABMDet1.Visible = False
    FrmABMDet2.Visible = False
    If ado_datos14.Recordset!modelo_elegido = "S" Then
        OpMod1.Value = True
        OpMod2.Value = False
        OpMod3.Value = False
    End If
    If ado_datos14.Recordset!modelo_elegido_h = "S" Then
        OpMod1.Value = False
        OpMod2.Value = True
        OpMod3.Value = False
    End If
    If ado_datos14.Recordset!modelo_elegido_x = "S" Then
        OpMod1.Value = False
        OpMod2.Value = False
        OpMod3.Value = True
    End If
    'dtc_codigo13.Text
    If ado_datos14.Recordset!par_codigo = "43340" Then
        dtc_codigo13.Text = "0"
        dtc_desc13.BoundText = dtc_codigo13.BoundText
        dtc_desc13.backColor = &H80000013
        dtc_desc13.ForeColor = &HFFFFFF
    Else
        dtc_desc13.backColor = &HFFFFFF
        dtc_desc13.ForeColor = &H80000008
    End If
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "select * from Gc_tipo_beneficiario where tipoben_codigo = '" & Ado_datos.Recordset!tipoben_codigo & "' ", db, adOpenKeyset, adLockReadOnly     'where venta_codigo = '" & TxtNroVenta.Text & "'
    Set Ado_datos12.Recordset = rs_datos12
    'Ado_datos12.Refresh
    Dtc_aux12.BoundText = dtc_codigo12.BoundText
    dtc_desc12.BoundText = dtc_codigo12.BoundText
    
    'Solo para Equipos (*)
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    rs_datos15.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    'rs_datos15.Open "select * from av_solicitud_cotiza_venta ", db, adOpenKeyset, adLockReadOnly
    Set ado_datos15.Recordset = rs_datos15
    ado_datos15.Refresh
  Else
    MsgBox "Los datos del registro Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
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
    If dtc_codigo11.Text = "L" Or dtc_codigo11.Text = "G" Then         'Hoja de Costos - CLIENTE - Importación Directa
        'cotiza_precio_total_dol_cli
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        rs_aux5.Open "Select sum(cotiza_precio_total_bs_cli) as totbs, sum(cotiza_precio_total_dol_cli) as totdl , sum(cotiza_cantidad) as cantot from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " AND estado_codigo_verif = 'APR' ", db, adOpenKeyset, adLockBatchOptimistic
        'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
        If rs_aux5.RecordCount > 0 Then
            TxtMontoBs.Text = IIf(IsNull(rs_aux5!totbs), 0, rs_aux5!totbs * rs_aux5!CANTOT)
            TxtMontoUsd.Text = IIf(IsNull(rs_aux5!totdl), 0, rs_aux5!totdl * rs_aux5!CANTOT)
            TxtCobrado.Text = 0
            TxtCobradoUsd.Text = 0
            TxtBstotal.Text = CDbl(TxtMontoBs.Text)
            TxtBstotalUsd.Text = CDbl(TxtMontoUsd.Text)
        End If
        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + txt_campo2.Text
    End If
    If dtc_codigo11.Text = "V" Then     'Facturación Local
        'cotiza_precio_total_dol_cge
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        rs_aux5.Open "Select sum(cotiza_precio_total_bs_cge) as totbs, sum(cotiza_precio_total_dol_cge) as totdl , sum(cotiza_cantidad) as cantot from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " AND estado_codigo_verif = 'APR' ", db, adOpenKeyset, adLockBatchOptimistic
        'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
        If rs_aux5.RecordCount > 0 Then
            TxtMontoBs.Text = IIf(IsNull(rs_aux5!totbs), 0, rs_aux5!totbs * rs_aux5!CANTOT)
            TxtMontoUsd.Text = IIf(IsNull(rs_aux5!totdl), 0, rs_aux5!totdl * rs_aux5!CANTOT)
            TxtCobrado.Text = 0
            TxtCobradoUsd.Text = 0
            TxtBstotal.Text = CDbl(TxtMontoBs.Text)
            TxtBstotalUsd.Text = CDbl(TxtMontoUsd.Text)
        End If
        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + txt_campo2.Text
        'TxtPlazo.Visible = True
    End If
    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "E" Then
            TxtConcepto.Text = "VENTA AL CONTADO - " + txt_campo2.Text
            TxtPlazo.Text = 0
            TxtPlazo.Visible = False
'        Else
'        'dtc_codigo2.Text = "VD"
'        'dtc_desc2.Text = "VENTA DIRECTA"
'        'TxtCobrado.Visible = True
'        'Label7.Visible = True
'            TxtConcepto.Text = "VENTA DIRECTA AL CLIENTE"
'            TxtPlazo.Text = 0
'            TxtPlazo.Visible = False
    End If
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

Private Sub dtc_precioventabase15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_precioventabase15.BoundText
    dtc_unimed15.BoundText = dtc_precioventabase15.BoundText
    dtc_stocktotal15.BoundText = dtc_precioventabase15.BoundText
    dtc_grupo15.BoundText = dtc_precioventabase15.BoundText
    dtc_subgrupo15.BoundText = dtc_precioventabase15.BoundText
    Dtc_partida15.BoundText = dtc_precioventabase15.BoundText
    dtc_precioventafinal15.BoundText = dtc_precioventabase15.BoundText
    dtc_codigo15.BoundText = dtc_precioventabase15.BoundText
    dtc_preciocompra15.BoundText = dtc_precioventabase15.BoundText
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

Private Sub dtc_desc15_LostFocus()
    txt_descripcion_venta.Text = dtc_desc15.Text
    TxtDescuento.Text = "0"
    TxtPrecioU.Text = dtc_precioventabase15.Text
'    Call AbreAlmacen
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
  If txt_campo2 = "" Then
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
        VAR_SOLA = Ado_datos.Recordset!venta_codigo
       marca1 = Ado_datos.Recordset.Bookmark
       db.Execute "Update ao_ventas_alcance set fecha_inicio_real = '" & DTPfechasol.Value & "', fecha_fin_real = '" & DTPfechaFin.Value & "', doc_codigo='R-321', correl_doc=" & Val(txt_campo1.Text) & "  WHERE venta_codigo = " & VAR_SOLA & " AND solicitud_tipo = '6' "
'       If Ado_datos.Recordset("venta_tipo") = "E" Then
'           db.Execute "INSERT INTO ao_ventas_cobranza_prog (venta_codigo, ges_gestion, beneficiario_codigo, beneficiario_codigo_resp, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, literal, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, factura_impresa, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
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

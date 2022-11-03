VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_almacen_salida 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacenes - Salida de Almacen"
   ClientHeight    =   10410
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   15270
   Icon            =   "aw_almacen_salida.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   1.90842e9
   ScaleMode       =   0  'User
   ScaleWidth      =   7.18315e9
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   2175
      Left            =   6120
      TabIndex        =   131
      Top             =   5880
      Visible         =   0   'False
      Width           =   6135
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   6120
         TabIndex        =   132
         Top             =   0
         Width           =   6120
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "aw_almacen_salida.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   134
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "aw_almacen_salida.frx":12CF
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   133
            Top             =   0
            Width           =   1400
         End
         Begin VB.Label Label22 
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
            Left            =   14175
            TabIndex        =   135
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   960
         TabIndex        =   138
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118751233
         CurrentDate     =   42880
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   3600
         TabIndex        =   139
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118751233
         CurrentDate     =   42880
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE FIN"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3600
         TabIndex        =   141
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE INICIO"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   140
         Top             =   1080
         Width           =   1620
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   10.688
      ScaleMode       =   4  'Character
      ScaleWidth      =   13.625
      TabIndex        =   89
      Top             =   5920
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H80000018&
         Height          =   525
         Left            =   120
         Picture         =   "aw_almacen_salida.frx":1BBB
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Imprime Kardex del Bien"
         Top             =   1830
         Width           =   1365
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   120
         Picture         =   "aw_almacen_salida.frx":2488
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Anula la Cobranza Identificada"
         Top             =   1275
         Width           =   1365
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Height          =   525
         Left            =   120
         Picture         =   "aw_almacen_salida.frx":2BD4
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   720
         Width           =   1365
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000015&
         Height          =   525
         Left            =   120
         Picture         =   "aw_almacen_salida.frx":34E9
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Adiciona Detalle"
         Top             =   150
         Width           =   1365
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   79
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton Command1 
         Caption         =   "Desaprobar"
         Height          =   375
         Left            =   7920
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6360
         Picture         =   "aw_almacen_salida.frx":3CA8
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   81
         ToolTipText     =   "Comprobante de Salida de Almacenes"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5040
         Picture         =   "aw_almacen_salida.frx":4575
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   82
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2520
         Picture         =   "aw_almacen_salida.frx":4D2A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   84
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3720
         Picture         =   "aw_almacen_salida.frx":5476
         ScaleHeight     =   735
         ScaleWidth      =   1320
         TabIndex        =   83
         ToolTipText     =   "Aprueba Entrega de Insumos"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "aw_almacen_salida.frx":5CA9
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   94
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "Digitaliza"
         Height          =   710
         Left            =   9600
         Picture         =   "aw_almacen_salida.frx":6468
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Height          =   855
         Left            =   3800
         Picture         =   "aw_almacen_salida.frx":68AA
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   -60
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1185
         Picture         =   "aw_almacen_salida.frx":72A1
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   85
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "aw_almacen_salida.frx":7BB6
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   80
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   12840
         TabIndex        =   88
         Top             =   200
         Width           =   1815
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
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20280
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "aw_almacen_salida.frx":8378
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   77
         Top             =   0
         Width           =   1280
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "aw_almacen_salida.frx":8B4E
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   76
         Top             =   0
         Width           =   1455
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13095
         TabIndex        =   78
         Top             =   180
         Width           =   1005
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   6600
      TabIndex        =   13
      Top             =   765
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "SOLICITUDES A ALMACEN"
      TabPicture(0)   =   "aw_almacen_salida.frx":943A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DETALLE BIENES (Insumos)"
      TabPicture(1)   =   "aw_almacen_salida.frx":9456
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrmEdita 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   4590
         Left            =   -75000
         TabIndex        =   18
         Top             =   360
         Width           =   11895
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   6225
            TabIndex        =   143
            Top             =   1830
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   8460
            TabIndex        =   100
            Top             =   2415
            Width           =   255
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   5220
            TabIndex        =   99
            Top             =   2415
            Width           =   255
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   2415
            Width           =   255
         End
         Begin VB.PictureBox FraGrabarDet 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   900
            Left            =   0
            ScaleHeight     =   840
            ScaleWidth      =   11880
            TabIndex        =   63
            Top             =   0
            Width           =   11940
            Begin VB.CommandButton CmdGrabaDet 
               BackColor       =   &H80000015&
               Height          =   650
               Left            =   4515
               Picture         =   "aw_almacen_salida.frx":9472
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   120
               Width           =   1365
            End
            Begin VB.CommandButton CmdCancelaDet 
               BackColor       =   &H80000016&
               Height          =   650
               Left            =   5850
               MaskColor       =   &H00000000&
               Picture         =   "aw_almacen_salida.frx":9C48
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Cancelar"
               Top             =   120
               Width           =   1365
            End
            Begin VB.CommandButton cmdElige 
               BackColor       =   &H80000018&
               Caption         =   "New Prod"
               Height          =   640
               Left            =   7440
               MaskColor       =   &H00000000&
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   120
               Visible         =   0   'False
               Width           =   825
            End
         End
         Begin VB.TextBox Txt_modelo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   6840
            TabIndex        =   61
            Text            =   "0"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   10980
            TabIndex        =   44
            Top             =   2900
            Width           =   255
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   8520
            TabIndex        =   43
            Top             =   1830
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   280
            Left            =   10980
            TabIndex        =   42
            Top             =   3850
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   10980
            TabIndex        =   41
            Top             =   1830
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_preciocompra15 
            Bindings        =   "aw_almacen_salida.frx":A534
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3840
            TabIndex        =   38
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
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
            Bindings        =   "aw_almacen_salida.frx":A54E
            CausesValidation=   0   'False
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   4320
            TabIndex        =   33
            Top             =   2400
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   741
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "subgrupo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_grupo15 
            Bindings        =   "aw_almacen_salida.frx":A568
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   960
            TabIndex        =   32
            Top             =   2400
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   741
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "grupo_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txt_descripcion_venta 
            CausesValidation=   0   'False
            DataField       =   "concepto_venta"
            DataSource      =   "ado_datos14"
            Enabled         =   0   'False
            Height          =   340
            Left            =   240
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3840
            Width           =   8505
         End
         Begin VB.TextBox TxtNroVenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "venta_codigo"
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
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox TxtCantidad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "venta_det_cantidad"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0"
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TxtDescuento 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "bien_cantidad_por_empaque"
            DataSource      =   "ado_datos14"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6600
            TabIndex        =   5
            Text            =   "0"
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0"
            Top             =   3255
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtPrecioU 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   3120
            TabIndex        =   9
            Text            =   "0"
            Top             =   3255
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dtc_precioventafinal15 
            Bindings        =   "aw_almacen_salida.frx":A582
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   4485
            TabIndex        =   19
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
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
            Bindings        =   "aw_almacen_salida.frx":A59C
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6600
            TabIndex        =   21
            Top             =   1815
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc15 
            Bindings        =   "aw_almacen_salida.frx":A5B6
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   240
            TabIndex        =   0
            Top             =   1815
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   5520
            TabIndex        =   6
            Top             =   1080
            Visible         =   0   'False
            Width           =   5730
            _ExtentX        =   10107
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_unimed15 
            Bindings        =   "aw_almacen_salida.frx":A5D0
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   9840
            TabIndex        =   34
            Top             =   1815
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   741
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "unimed_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_stocktotal15 
            Bindings        =   "aw_almacen_salida.frx":A5EA
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   9840
            TabIndex        =   36
            Top             =   3840
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   741
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "bien_stock_actual"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_codigo13 
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7560
            TabIndex        =   39
            Top             =   840
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
            Bindings        =   "aw_almacen_salida.frx":A604
            Height          =   315
            Left            =   9840
            TabIndex        =   40
            Top             =   2880
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   741
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo Dtc_partida15 
            Bindings        =   "aw_almacen_salida.frx":A61E
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   7320
            TabIndex        =   45
            Top             =   2400
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "par_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_precioventabase15 
            Bindings        =   "aw_almacen_salida.frx":A638
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3120
            TabIndex        =   60
            Top             =   2760
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
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
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Grupo"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            TabIndex        =   95
            Top             =   2400
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Partida"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   6480
            TabIndex        =   59
            Top             =   2415
            Width           =   645
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Almacen de Origen:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3600
            TabIndex        =   26
            Top             =   1100
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Actual (Total Nacional)"
            ForeColor       =   &H00000000&
            Height          =   480
            Left            =   9675
            TabIndex        =   37
            Top             =   3360
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Medida"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9795
            TabIndex        =   35
            Top             =   1560
            Width           =   1515
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Venta:"
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
            Height          =   240
            Left            =   240
            TabIndex        =   31
            Top             =   1095
            Width           =   1500
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción y/o Características Complementarias"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   30
            Top             =   3555
            Width           =   4425
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Entregada"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   4560
            TabIndex        =   29
            Top             =   3000
            Width           =   1800
         End
         Begin VB.Label lbl_des_bien 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Bien"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   28
            Top             =   1560
            Width           =   1860
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Bien"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6600
            TabIndex        =   27
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Solicitada"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   225
            TabIndex        =   25
            Top             =   3000
            Width           =   1770
         End
         Begin VB.Label Label24 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   2415
            Width           =   690
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Actual (Almacen Origen)"
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   9660
            TabIndex        =   23
            Top             =   2400
            Width           =   1785
         End
      End
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
         Height          =   4590
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   11895
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "----------------------------- DESTINO "
            ForeColor       =   &H00C00000&
            Height          =   1815
            Left            =   5960
            TabIndex        =   119
            Top             =   2715
            Width           =   5895
            Begin MSDataListLib.DataCombo dtc_desc5 
               Bindings        =   "aw_almacen_salida.frx":A652
               DataField       =   "beneficiario_codigo_tec"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1400
               TabIndex        =   4
               Top             =   300
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_codigo5 
               Bindings        =   "aw_almacen_salida.frx":A66B
               DataField       =   "beneficiario_codigo_tec"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4275
               TabIndex        =   123
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_desc20 
               Bindings        =   "aw_almacen_salida.frx":A684
               DataField       =   "almacen_codigo_d"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1395
               TabIndex        =   124
               Top             =   840
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "almacen_descripcion"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_desc22 
               Bindings        =   "aw_almacen_salida.frx":A69E
               DataField       =   "depto_codigo_d"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1400
               TabIndex        =   125
               Top             =   1320
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo20 
               Bindings        =   "aw_almacen_salida.frx":A6B8
               DataField       =   "almacen_codigo_d"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4755
               TabIndex        =   126
               Top             =   840
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "almacen_codigo"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo22 
               Bindings        =   "aw_almacen_salida.frx":A6D2
               DataField       =   "depto_codigo_d"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4755
               TabIndex        =   127
               Top             =   1320
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_Aux20 
               Bindings        =   "aw_almacen_salida.frx":A6EC
               DataField       =   "almacen_codigo_d"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   3600
               TabIndex        =   129
               Top             =   840
               Visible         =   0   'False
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin VB.Label lbl_Rdestino 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Regional "
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   480
               TabIndex        =   122
               Top             =   1365
               Width           =   870
            End
            Begin VB.Label lbl_Adestino 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Almacen "
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   480
               TabIndex        =   121
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Entregado a:"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   120
               Top             =   360
               Width           =   1155
            End
         End
         Begin VB.Frame Fra_datos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "-------------------------------- ORIGEN "
            ForeColor       =   &H00C00000&
            Height          =   1845
            Left            =   40
            TabIndex        =   107
            Top             =   2715
            Width           =   5895
            Begin VB.ComboBox cmb_mes_ini 
               DataField       =   "mes_inicio_crono"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   0
               TabIndex        =   109
               Text            =   "SEPTIEMBRE"
               Top             =   1080
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.ComboBox cmd_unimed2 
               DataField       =   "unimed_codigo_cobr"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   6210
               TabIndex        =   108
               Text            =   "ANUAL"
               Top             =   1080
               Visible         =   0   'False
               Width           =   555
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               Bindings        =   "aw_almacen_salida.frx":A706
               DataField       =   "beneficiario_codigo_alm"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   110
               Top             =   300
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               Bindings        =   "aw_almacen_salida.frx":A71F
               DataField       =   "almacen_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   111
               Top             =   840
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "almacen_descripcion"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               Bindings        =   "aw_almacen_salida.frx":A739
               DataField       =   "almacen_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   112
               Top             =   855
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "almacen_codigo"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               Bindings        =   "aw_almacen_salida.frx":A753
               DataField       =   "beneficiario_codigo_alm"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   113
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin MSDataListLib.DataCombo dtc_desc21 
               Bindings        =   "aw_almacen_salida.frx":A76C
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   1380
               TabIndex        =   114
               Top             =   1320
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "depto_descripcion"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo21 
               Bindings        =   "aw_almacen_salida.frx":A786
               DataField       =   "depto_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   4320
               TabIndex        =   115
               Top             =   1320
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "depto_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_Aux11 
               Bindings        =   "aw_almacen_salida.frx":A7A0
               DataField       =   "almacen_codigo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2880
               TabIndex        =   128
               Top             =   840
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "depto_codigo"
               BoundColumn     =   "almacen_codigo"
               Text            =   ""
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Responsable"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   120
               TabIndex        =   118
               Top             =   360
               Width           =   1230
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Almacen"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   525
               TabIndex        =   117
               Top             =   855
               Width           =   795
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Regional"
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   480
               TabIndex        =   116
               Top             =   1365
               Width           =   825
            End
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   6190
            TabIndex        =   104
            Top             =   390
            Width           =   270
         End
         Begin VB.TextBox TxtConcepto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "venta_descripcion"
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   2280
            Width           =   10515
         End
         Begin VB.TextBox Txt_campo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "unidad_codigo_ant"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            DataSource      =   "Ado_datos"
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   5280
            TabIndex        =   73
            Text            =   "0"
            Top             =   1300
            Width           =   1815
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8025
            TabIndex        =   58
            Top             =   390
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "aw_almacen_salida.frx":A7BA
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6465
            TabIndex        =   57
            Top             =   380
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "aw_almacen_salida.frx":A7D3
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   165
            TabIndex        =   1
            Top             =   380
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   280
            Left            =   7980
            TabIndex        =   62
            Top             =   870
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "aw_almacen_salida.frx":A7EC
            DataField       =   "unidad_destino"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6585
            TabIndex        =   2
            Top             =   1800
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "aw_almacen_salida.frx":A805
            DataField       =   "unidad_destino"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   50
            Top             =   1800
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   4210752
            ListField       =   "unidad_descripcion"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "aw_almacen_salida.frx":A81E
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6960
            TabIndex        =   53
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "unidad_codigo"
            BoundColumn     =   "unidad_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "aw_almacen_salida.frx":A837
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1860
            TabIndex        =   130
            Top             =   860
            Width           =   6405
            _ExtentX        =   11298
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
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "aw_almacen_salida.frx":A850
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6960
            TabIndex        =   56
            Top             =   240
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "estado_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "fecha_verif"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   10160
            TabIndex        =   106
            Top             =   1800
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            Format          =   118751233
            CurrentDate     =   42682
            MaxDate         =   55153
            MinDate         =   2
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Destino"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   142
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Entrega al Solicitante"
            ForeColor       =   &H00000000&
            Height          =   480
            Left            =   8400
            TabIndex        =   105
            Top             =   1695
            Width           =   1710
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFF80&
            X1              =   11880
            X2              =   8520
            Y1              =   1160
            Y2              =   1160
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF80&
            X1              =   8520
            X2              =   8520
            Y1              =   0
            Y2              =   1160
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   103
            Top             =   2280
            Width           =   1020
            WordWrap        =   -1  'True
         End
         Begin VB.Label DTPFechaFin 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_fecha_fin"
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
            Left            =   5535
            TabIndex        =   102
            Top             =   2040
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label DTPFechaIni 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_fecha_inicio"
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
            Left            =   2115
            TabIndex        =   101
            Top             =   2040
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio Contrato"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   60
            TabIndex        =   97
            Top             =   2055
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Fin Contrato"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   3795
            TabIndex        =   96
            Top             =   2055
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label lbl_cerrado 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "TRAMITE CERRADO !!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   3480
            TabIndex        =   74
            Top             =   -30
            Width           =   4875
         End
         Begin VB.Label lbl_cite 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cite de Trámite"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3780
            TabIndex        =   72
            Top             =   1320
            Width           =   1485
         End
         Begin VB.Label txt_venta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
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
            Left            =   10500
            TabIndex        =   71
            Top             =   1300
            Width           =   1245
         End
         Begin VB.Label txt_codigo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_codigo_alm"
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
            Left            =   10395
            TabIndex        =   70
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Código de Registro"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   8610
            TabIndex        =   69
            Top             =   735
            Width           =   1770
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Documento"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   13
            Left            =   8880
            TabIndex        =   68
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "doc_numero_alm"
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
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   10395
            TabIndex        =   67
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Edificio / Destino"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   55
            Top             =   100
            Width           =   1500
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
            Left            =   1620
            TabIndex        =   54
            Top             =   1300
            Width           =   1245
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Código Trámite"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   52
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lbl_campo1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Solicitante"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   180
            TabIndex        =   51
            Top             =   880
            Width           =   1635
         End
         Begin VB.Label Label15 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Venta"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9060
            TabIndex        =   17
            Top             =   1320
            Width           =   1245
         End
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   5040
      Left            =   120
      TabIndex        =   46
      Top             =   720
      Width           =   6465
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salidas"
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
         Left            =   4680
         TabIndex        =   137
         Top             =   4635
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Traspasos"
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
         TabIndex        =   136
         Top             =   4635
         Width           =   1395
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
         Left            =   2280
         TabIndex        =   49
         Top             =   4635
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
         Left            =   960
         TabIndex        =   48
         Top             =   4635
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "aw_almacen_salida.frx":A869
         Height          =   4290
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   7567
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
         ColumnCount     =   12
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
            DataField       =   "almacen_codigo"
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
         BeginProperty Column02 
            DataField       =   "doc_numero_alm"
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
         BeginProperty Column03 
            DataField       =   "fecha_verif"
            Caption         =   "Fecha Salida"
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
         BeginProperty Column04 
            DataField       =   "edif_descripcion"
            Caption         =   "Destino/Nombre_Edificio"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Trámite/OS"
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
            Caption         =   "Unidad"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Tramite"
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
            DataField       =   "estado_almacen"
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
         BeginProperty Column09 
            DataField       =   "edif_codigo"
            Caption         =   "Edif/Destino"
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
         BeginProperty Column11 
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4560
         Width           =   6225
         _ExtentX        =   10980
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
      Caption         =   "DETALLE DE BIENES"
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
      Height          =   2745
      Left            =   1920
      TabIndex        =   14
      Top             =   5820
      Width           =   16695
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "aw_almacen_salida.frx":A881
         Height          =   2340
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   4128
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
            Caption         =   "Cantidad Solicitada"
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
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Cantidad Entregada"
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
            DataField       =   "estado_almacen"
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
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   6480
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
            EndProperty
         EndProperty
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   0
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   6720
      Top             =   8760
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
      Left            =   2160
      Top             =   8760
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
      Left            =   11280
      Top             =   9120
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
      Left            =   9000
      Top             =   9120
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
      Top             =   9120
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
      Left            =   13560
      Top             =   9120
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
      Left            =   6720
      Top             =   9120
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   11280
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_Datos12 
      Height          =   330
      Left            =   2160
      Top             =   9120
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
      Left            =   4440
      Top             =   9120
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
      Left            =   13560
      Top             =   8760
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
      Left            =   4440
      Top             =   8760
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
      Top             =   8760
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
      Left            =   9000
      Top             =   8760
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
      Left            =   480
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos20 
      Height          =   330
      Left            =   0
      Top             =   9480
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   2280
      Top             =   9480
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
      Caption         =   "Ado_datos21"
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
   Begin MSAdodcLib.Adodc Ado_datos22 
      Height          =   330
      Left            =   4560
      Top             =   9480
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
      Caption         =   "Ado_datos22"
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
   Begin MSAdodcLib.Adodc AdoAux9 
      Height          =   330
      Left            =   0
      Top             =   9840
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
      Caption         =   "AdoAux9"
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
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "aw_almacen_salida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Ventas
Dim rs_datos As New ADODB.Recordset     'VENTAS
Dim rs_datos1 As New ADODB.Recordset    'UNIDAD EJECUTORA
Dim rs_datos2 As New ADODB.Recordset    'Beneficiario Personas Nat. y Juridicas (menos de CGI)
Dim rs_datos3 As New ADODB.Recordset    'Proyecto de Edificacion
Dim rs_datos4 As New ADODB.Recordset    'Beneficiario Funcionario de CGI (Vendedor, Cobrador, Admin, etc.)
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset
Dim rs_datos13 As New ADODB.Recordset
Dim rs_datos14 As New ADODB.Recordset   'Ventas_detalle
Dim rs_datos15 As New ADODB.Recordset
Dim rs_datos16 As New ADODB.Recordset   'Ventas cobranzas
Dim rs_datos17 As New ADODB.Recordset
Dim rs_datos18 As New ADODB.Recordset

Dim rs_datos19 As New ADODB.Recordset   'Acumula Cobranzas
Dim rs_datos20 As New ADODB.Recordset
Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos22 As New ADODB.Recordset

'AUXILIARES
Dim rs_Ventas_lista As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_precio As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstcorrel_ing As New ADODB.Recordset
'Dim rs_precio As New ADODB.Recordset
'OTROS
'Dim rstdetsalalm As New ADODB.Recordset
Dim RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rs_almacen2 As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rsAuxDetalle As New ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir, accion As String
'Dim queryinicial As String
Dim queryinicial2 As String

'Almacenes
Dim descri_bien As String
Dim VAR_ALMX As String
Dim VAR_ALMT As String
Dim tipo_alm As String
Dim VAR_DOC As String
Dim VAR_DA As String
Dim VAR_ALMD As String
Dim VAR_ORIGEN As String
Dim VAR_DOCI, VAR_DOCR, VAR_DOCH, VAR_DOCA As String
Dim VAR_BENI, VAR_BENR, VAR_BENH, VAR_BENA As String
Dim VAR_BENDI, VAR_BENDR, VAR_BENDH, VAR_BENDA As String
Dim VAR_NUMI, VAR_NUMR, VAR_NUMH, VAR_NUMA As String

Dim Cant_Alm, VAR_CANT As Integer
Dim correlativo1 As Integer
Dim VAR_ALMI, VAR_ALMR, VAR_ALMH, VAR_ALMA As Integer
Dim VAR_ALMDI, VAR_ALMDR, VAR_ALMDH, VAR_ALMDA As Integer

'VARIABLES
Dim marca1 As Variant

Dim swgrabar, swnuevo, deta2, CONT_MED As Integer
Dim nroventa, correlv, correldet2, corrprog As Integer
Dim VAR_PARTIDA, VAR_PROY, correldetalle As Integer
Dim VAR_CODANT, Var_Comp, VAR_SOL, CANTOT, var_cod5 As Integer
Dim CONT2, CONT3, CONT4, VAR_TIPO As Integer
Dim fdia, fmes, fanio, Dias_Mes, TimeD  As Integer
Dim VAR_COBR1, VAR_COBR2, VAR_CONTR As Integer
Dim VAR_NUM, var_cod, VAR_COD2 As Integer
Dim VAR_DET As String

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double

Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_BIEN2, VAR_COD3, VAR_COD4 As String
Dim VAR_MED, VAR_MED2 As String
Dim VAR_TIPOV, VAR_VAL As String
Dim VAR_FEC2, MControl, VAR_MES2 As String
Dim VAR_BEN2, VAR_BEN3, VAR_ALM As String
Dim VAR_BIEN, VAR_R As String
Dim VAR_N1, VAR_N2, VAR_N3, VAR_POA As String

Dim FInicio, FFin, FControl, FVenta, VAR_FECHA As Date
Dim precio_tot, precio_uni As Double


Private Sub CmdDetalle_Click()
'    FrmCobranza.Visible = True
End Sub

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
DTPfechasol.Value = IIf(IsNull(Ado_datos.Recordset!fecha_verif), Date, Ado_datos.Recordset!fecha_verif)
If parametro <> Ado_datos.Recordset!unidad_codigo Then
BtnAnlDetalle.Visible = False
Else
BtnAnlDetalle.Visible = True
End If
   If Not IsNull(Ado_datos.Recordset("venta_codigo")) Then
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
        If (Ado_datos.Recordset("estado_almacen") = "REG") Then
            BtnAprobar.Visible = True
            BtnDesAprobar.Visible = False
            BtnModificar.Visible = True
            BtnEliminar.Visible = True
'            BtnVer.Visible = False
'            BtnAñadir.Visible = False   'Cerrar Tramite
'            BtnVer3.Visible = False     'Provisional
            lbl_cerrado.Caption = ""
'            BtnImprimir2.Visible = True
'            If IsNull(Ado_datos.Recordset("venta_tipo")) Then
'                FrmABMDet.Visible = False
''                FrmABMDet2.Visible = False
''                FrmCobranza.Visible = False
'            Else
            FrmABMDet.Visible = True
''                FrmABMDet2.Visible = True
''                FrmCobranza.Visible = True
'            End If
        Else
            BtnAprobar.Visible = False
            BtnDesAprobar.Visible = True
            BtnModificar.Visible = False
            BtnEliminar.Visible = False
'            BtnVer.Visible = True
'            Select Case Ado_datos.Recordset!estado_cancelado
'                Case "S"
'                    lbl_cerrado.Caption = "TRAMITE CERRADO !!"
''                    FrmABMDet2.Visible = False
''                    BtnAñadir.Visible = False   'Cerrar Tramite
''                    BtnVer3.Visible = False     'Provisional
'                Case "P"
'                    lbl_cerrado.Caption = "TRAMITE PROVISIONAL !!"
''                    FrmABMDet2.Visible = True
''                    BtnAñadir.Visible = False   'Cerrar Tramite5
''                    BtnVer3.Visible = False     'Provisional
'                Case Else
''                    BtnAñadir.Visible = True   'Cerrar Tramite
''                    BtnVer3.Visible = True     'Provisional
'                    lbl_cerrado.Caption = ""
''                    FrmABMDet2.Visible = True
'            End Select
            FrmABMDet.Visible = False
''            FrmCobranza.Visible = True
'            BtnImprimir2.Visible = True
        End If
        
'        If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Then
''            TxtPlazo.Visible = True
''            BtnAddDetalle2.Visible = True
'        Else
''            TxtPlazo.Visible = False
'            If Ado_datos.Recordset("venta_tipo") = "E" Then
''                BtnAddDetalle2.Visible = False
'            End If
'        End If

        If Ado_datos.Recordset!edif_codigo = "20101-2" Or Ado_datos.Recordset!edif_codigo = "70101-2" Or Ado_datos.Recordset!edif_codigo = "30101-2" Or Ado_datos.Recordset!edif_codigo = "10101-2" Then
            dtc_desc20.Visible = True
            lbl_Adestino.Visible = True
            dtc_desc22.Visible = True
            lbl_Rdestino.Visible = True
        Else
            dtc_desc20.Visible = False
            lbl_Adestino.Visible = False
            dtc_desc22.Visible = False
            lbl_Rdestino.Visible = False
        End If
        Call AbrirDetalle
        
        FrmDetalle.Caption = "BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        If Ado_datos.Recordset!unidad_codigo = "DNREP" Or Ado_datos.Recordset!unidad_codigo = "DNEME" Then
            lbl_cite = "Cite / O.S."
        Else
            lbl_cite = "Cite Trámite"
        End If

        End If
        FrmDetalle.Visible = True
'        FrmCobranza.Visible = True
    Else
        FrmABMDet.Visible = False
        FrmDetalle.Visible = False
        If buscados = 0 Then
           OptFilGral1.Visible = True
           OptFilGral2.Visible = True
        Else
           OptFilGral1.Visible = False
           OptFilGral2.Visible = False
        End If
    End If
        BtnEliminar.Visible = True
End Sub

Private Sub AbrirDetalle()
    Set rs_datos14 = New ADODB.Recordset
        If rs_datos14.State = 1 Then rs_datos14.Close
        '
        rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  and almacen_tipo = '" & VAR_ALMT & "'  order by  par_codigo, bien_codigo ", db, adOpenKeyset, adLockOptimistic
        'rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  and par_codigo <> '43340'  order by  par_codigo, bien_codigo ", db, adOpenKeyset, adLockOptimistic
        rs_datos14.Sort = "hora_registro"
        Set ado_datos14.Recordset = rs_datos14
        ado_datos14.Recordset.Requery
        If ado_datos14.Recordset.RecordCount > 0 Then
            deta2 = 1
                        
            Call AbreAlmacen
        Else
            deta2 = 0
        End If
End Sub

Private Sub AbreAlmacen()
    Set rs_datos13 = New ADODB.Recordset
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet where coddetalle= '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from Av_almacen_detalle where bien_codigo = '" & dtc_codigo15.Text & "' ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh

End Sub

Private Sub Ado_datos16_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 If (Not Ado_datos16.Recordset.BOF) And (Not Ado_datos16.Recordset.EOF) Then
    If Not IsNull(Ado_datos16.Recordset("venta_codigo")) Then
        'BtnModDetalle2.Visible = False
        If (Ado_datos16.Recordset("estado_almacen") = "REG") Then
'            If (Ado_datos.Recordset("estado_codigo") = "APR") Then
'                BtnAprobar2.Visible = False
'            Else
'                BtnAprobar2.Visible = True
'            End If
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = True
'            BtnAnlDetalle2.Visible = True
'            BtnModDetalle2.Visible = True
        End If
        If (Ado_datos16.Recordset("estado_almacen") = "APR") Then
'            BtnImprimir2.Visible = True
'            BtnAprobar2.Visible = False
'            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
        End If
        If (Ado_datos16.Recordset("estado_almacen") = "ANL") Then
''            'BtnImprimir2.Visible = False
'            BtnAnlDetalle2.Visible = False
'            BtnModDetalle2.Visible = False
'            BtnAprobar2.Visible = False
        End If
    Else
        BtnAprobar2.Visible = False
''        BtnImprimir2.Visible = False
'        BtnAnlDetalle2.Visible = False
        BtnModDetalle2.Visible = False
    End If
 Else
    BtnAprobar2.Visible = False
    BtnImprimir2.Visible = False
'    BtnAnlDetalle2.Visible = False
    BtnModDetalle2.Visible = False
 End If
End Sub

Private Sub BtnAddDetalle_Click()
On Error GoTo UpdateErr
If Ado_datos.Recordset("estado_almacen") = "REG" Then
  If Ado_datos.Recordset!almacen_codigo <> "" Then
    If Ado_datos.Recordset!doc_numero_alm > 0 Then
        Text9.Visible = False
        FraNavega.Enabled = False
        FrmDetalle.Enabled = False
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(0) = False
        FrmEdita.Visible = True
        FrmEdita.Enabled = True
        FrmABMDet.Visible = False
        swnuevo = 1
        'Bienes por almacen
        Set rs_datos15 = New ADODB.Recordset
        If rs_datos15.State = 1 Then rs_datos15.Close
        'Select Case parametro
        Select Case VAR_ORIGEN
            Case "UALMI"    'INSUMOS
                'rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'I' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
                rs_datos15.Open "select * from av_ac_bienes_vs_ao_almacenes_totales where almacen_tipo = 'I' AND almacen_codigo = " & Ado_datos.Recordset!almacen_codigo & " and stock_actual > 0 ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
                Set ado_datos15.Recordset = rs_datos15
                ado_datos15.Refresh
    '            VAR_DET = "30000"
            Case "UALMR"    'REPUESTOS
                rs_datos15.Open "select * from av_ac_bienes_vs_ao_almacenes_totales where almacen_tipo = 'R' AND almacen_codigo = " & Ado_datos.Recordset!almacen_codigoR & " and stock_actual > 0 ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
    '            VAR_DET = "39800"
                Set ado_datos15.Recordset = rs_datos15
                ado_datos15.Refresh
            Case "UALMH"    'HERRAMIENTAS
                rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'H' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
    '            VAR_DET = "34800"
                Set ado_datos15.Recordset = rs_datos15
                ado_datos15.Refresh
        End Select
        'WWWWWWWWWWWWWWWWWWWW DIF-01
        Dtc_Stock13.Text = ""
        ado_datos14.Recordset.AddNew
        dtc_desc15.SetFocus
        TxtNroVenta.Text = Ado_datos.Recordset!venta_codigo  'txt_venta.Text
        TxtNroVenta.Locked = True
    Else
        MsgBox "Debe generar el Nro. Documento, verifique en Solicitudes a Almacen y vuelva a intentar ...", , "Atención"
    End If
  Else
        MsgBox "Debe registrar el Almacen ORIGEN, verifique en Solicitudes a Almacen y vuelva a intentar ...", , "Atención"
  End If
End If
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnAñadir_Click()
accion = "NEW"
On Error GoTo UpdateErr
    Ado_datos.Recordset.AddNew
    txt_codigo1.Caption = VAR_R
    If parametro = "" Then
        dtc_codigo1.Text = "DCONT"
    Else
        dtc_codigo1.Text = parametro
    End If
    Fra_datos.Enabled = True
    dtc_desc3.backColor = &H80000005
    dtc_desc3.ForeColor = &H80000008
    txt_campo1.Caption = "0"
    dtc_desc3.Locked = False
    dtc_desc3.Width = 5955
    lbl_campo4.Visible = False
    DTPFechaIni.Visible = False
    lbl_campo5.Visible = False
    DTPFechaFin.Visible = False
    DTPfechasol.Value = Date
    swgrabar = 1
    FrmCabecera.Enabled = True
    FrmDetalle.Visible = False
    FraNavega.Enabled = False
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    Fra_datos.Enabled = True
    '        Fra_Total.Visible = False
    FrmABMDet.Visible = False
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    Call CARGAPARAM
    If VAR_ALMT <> "" Then
    'If VAR_ALMH = "" Then  'VAR_ALMH <> ""
        dtc_codigo4.Text = VAR_BENEF
        dtc_desc4.BoundText = dtc_codigo4.BoundText
        dtc_codigo11.Text = VAR_ALMX
        dtc_desc11.BoundText = dtc_codigo11.BoundText
        dtc_Aux11.BoundText = dtc_codigo11.BoundText
        dtc_codigo21.Text = VAR_DPTO
        dtc_desc21.BoundText = dtc_codigo21.BoundText
        
        'ac_almacenes ' Origen
        Set rs_datos11 = New ADODB.Recordset
        If rs_datos11.State = 1 Then rs_datos11.Close
        'rs_datos11.Open "select * from ac_almacenes where depto_codigo = '" & VAR_DPTO & "' AND almacen_tipo = '" & VAR_ALMT & "' ", db, adOpenStatic
        If VAR_BENEF = "0" Then
            rs_datos11.Open "select * from ac_almacenes where almacen_codigo <> '1' and almacen_codigo <> '2'  ", db, adOpenStatic
        Else
            rs_datos11.Open "select * from ac_almacenes where beneficiario_codigo = '" & VAR_BENEF & "'  ", db, adOpenStatic
        End If
        Set Ado_datos11.Recordset = rs_datos11
        dtc_desc11.BoundText = dtc_codigo11.BoundText
        If Ado_datos11.Recordset.RecordCount > 0 Then
           Ado_datos11.Recordset.MoveFirst
           VAR_ALMT = rs_datos11!almacen_tipo
           VAR_DPTO = rs_datos11!depto_codigo
           VAR_ALMX = rs_datos11!almacen_codigo
'           dtc_desc11.BoundText = VAR_ALMX
'           dtc_desc21.BoundText = VAR_DPTO
'           dtc_desc4.BoundText VAR_BENEF
        Else
           VAR_ALMT = ""
           VAR_DPTO = ""
           VAR_ALMX = ""
        End If
        'ac_almacenes - Destino
        Set rs_datos20 = New ADODB.Recordset
        If rs_datos20.State = 1 Then rs_datos20.Close
        rs_datos20.Open "select * from ac_almacenes where beneficiario_codigo <> '" & VAR_BENEF & "'  ", db, adOpenStatic
        'rs_datos20.Open "select * from ac_almacenes where depto_codigo <> '" & VAR_DPTO & "' AND almacen_tipo <> '" & VAR_ALMT & "' ", db, adOpenStatic
        Set Ado_datos20.Recordset = rs_datos20
        dtc_desc20.BoundText = dtc_codigo20.BoundText
        
        'gc_departamento - Origen
        Set rs_datos21 = New ADODB.Recordset
        If rs_datos21.State = 1 Then rs_datos21.Close
        'rs_datos21.Open "select * from gc_departamento where depto_codigo = '" & Left(dtc_codigo3.Text, 1) & "'  ", db, adOpenStatic      ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
        rs_datos21.Open "select * from gc_departamento where depto_codigo = '" & VAR_DPTO & "'  ", db, adOpenStatic      ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
        Set Ado_datos21.Recordset = rs_datos21
        dtc_desc21.BoundText = dtc_codigo21.BoundText
        
        'gc_departamento - Destino
        Set rs_datos22 = New ADODB.Recordset
        If rs_datos22.State = 1 Then rs_datos22.Close
        rs_datos22.Open "select * from gc_departamento where depto_codigo <>  '" & VAR_DPTO & "'  ", db, adOpenStatic       ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
        Set Ado_datos22.Recordset = rs_datos22
        dtc_desc22.BoundText = dtc_codigo22.BoundText
    End If
    
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

'Private Function ExisteReg(Unidad As String, Codigo As Integer) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & Codigo & " and estado_codigo = 'APR'   "
''    <> 'ANL'
'    rs.Open GlSqlAux, db, adOpenStatic
'    ExisteReg = rs!Cuantos > 0
'End Function

Private Sub BtnAprobar_Click()

'On Error GoTo UpdateErr
'  VAR_VAL = "OK"
'  Call valida_campos
'  If VAR_VAL = "OK" Then
'     If Ado_datos.Recordset.RecordCount > 0 Then
'       If ado_datos14.Recordset.RecordCount = 0 Then       'Or (Ado_datos.Recordset("venta_monto_total_bs") = 0)       ' JQA ENE-2016
'       'If IsNull(Ado_datos.Recordset("venta_tipo")) Or (Ado_datos.Recordset("venta_monto_total_bs") = 0) Or (Ado_datos.Recordset!estado_alcance = "N") Then
'            MsgBox "No se puede APROBAR el registro, verifique el Detalle y vuelva a intentar ...", , "Atención"
'            Exit Sub
'       Else
'         If Ado_datos.Recordset("estado_almacen") = "REG" Then
'           sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'           If sino = vbYes Then
'               correlv = Ado_datos.Recordset!venta_codigo
'               VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'               VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'               VAR_MED = Ado_datos.Recordset!unimed_codigo
'               VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
'               FInicio = Ado_datos.Recordset!venta_fecha_inicio
'               FFin = Ado_datos.Recordset!venta_fecha_fin
'               TimeD = Ado_datos.Recordset!venta_plazo_dias_calendario
'               CANTOT = Ado_datos.Recordset!venta_cantidad_total
'               VAR_COD1 = Ado_datos.Recordset!unidad_codigo
'               VAR_GLOSA2 = Ado_datos.Recordset!venta_descripcion
'               VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'               VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant
'               VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'               VAR_ALMX = Ado_datos.Recordset!almacen_codigo
'               VAR_BEND = dtc_desc2.Text
'               VAR_EDIFD = dtc_desc3.Text
'               VAR_UNID = dtc_desc1.Text
'               VAR_DPTO = Left(VAR_PROY2, 1)
'               VARG_ORGD = ""
'               VAR_CTAD = ""
'               'Dim VARG_ORGD, VAR_CTAD, ,  As String
'               'Actualiza venta_precio_total_bs y venta_precio_total_dol
'               'db.Execute "update ao_ventas_detalle set venta_precio_total_bs = round(venta_det_cantidad * venta_precio_unitario_bs,2)  "
'               'db.Execute "update ao_ventas_detalle set venta_precio_total_dol = venta_det_cantidad * venta_precio_unitario_dol  "
'
'               'INI Deptos de Bolivia
'               Select Case VAR_DPTO
'                    Case "1"
'                        VAR_DPTOD = "CHUQUISACA"
'                    Case "2"
'                        VAR_DPTOD = "LA PAZ"
'                    Case "3"
'                        VAR_DPTOD = "COCHABAMBA"
'                    Case "4"
'                        VAR_DPTOD = "ORURO"
'                    Case "5"
'                        VAR_DPTOD = "POTOSI"
'                    Case "6"
'                        VAR_DPTOD = "TARIJA"
'                    Case "7"
'                        VAR_DPTOD = "SANTA CRUZ"
'                    Case "8"
'                        VAR_DPTOD = "BENI"
'                    Case "9"
'                        VAR_DPTOD = "PANDO"
'               End Select
'
''               If Ado_datos.Recordset("estado_alcance") = "N" Then
''                    db.Execute "INSERT INTO ao_ventas_alcance (ges_gestion, venta_codigo, solicitud_tipo, solicitud_tipo_descripcion, unidad_codigo_tec, venta_tiempo_dias, fecha_inicio_alcance, fecha_fin_alcance , estado_codigo, usr_codigo, fecha_registro) VALUES ('" & glGestion & "', '" & correlv & "', " & VAR_TIPO & ", 'MANTENIMIENTO PREVENTIVO DE EQUIPOS', '" & VAR_COD1 & "', '" & TimeD & "', '" & FInicio & "' , '" & FFin & "', 'APR', '" & glusuario & "', '" & Date & "' )"
''                    db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_alcance = 'S' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
''               End If
'
''               Call acumulaMont(Ado_datos.Recordset("ges_gestion"), correlv)
''                'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
''                Set rs_aux5 = New ADODB.Recordset
''                If rs_aux5.State = 1 Then rs_aux5.Close
''                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_R & "'  "
''                rs_aux5.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''                If rs_aux5.RecordCount > 0 Then
''                    rs_aux5!correl_doc = rs_aux5!correl_doc + 1
''                    VAR_NUM = rs_aux5!correl_doc
''                    'Txt_campo1.Caption = rs_aux2!correl_doc
''                    rs_aux5.Update
''                End If
''                db.Execute "UPDATE ao_ventas_cabecera SET doc_codigo_alm = '" & VAR_R & "' WHERE venta_codigo = " & correlv & " "
''                db.Execute "UPDATE ao_ventas_cabecera SET doc_numero_alm = " & VAR_NUM & " WHERE venta_codigo = " & correlv & " "
''                db.Execute "UPDATE ao_ventas_detalle SET doc_codigo_alm = '" & VAR_R & "' WHERE (par_codigo <> '43340') AND (bien_cantidad_por_empaque > 0) AND venta_codigo = " & correlv & " "
''                db.Execute "UPDATE ao_ventas_detalle SET doc_numero_alm = " & VAR_NUM & " WHERE (par_codigo <> '43340') AND (bien_cantidad_por_empaque > 0) AND venta_codigo = " & correlv & " "
'                'INI GRABA ao_ventas_alcance
'               Select Case VAR_COD1
'                    Case "DNINS"
'                        VAR_TIPO = 4
'                    Case "DNAJS"
'                        VAR_TIPO = 5
'                    Case "DNMAN"
'                        VAR_TIPO = 10
'                    Case "DNREP"
'                        VAR_TIPO = 7
'                    Case "DNEME"
'                        VAR_TIPO = 8
'                    Case "DNMOD"
'                        VAR_TIPO = 9
'                    Case "UALMI", "ALMIS", "ALMIB", "ALMIC"
'                        db.Execute "UPDATE ao_ventas_cabecera SET doc_numero = " & VAR_NUM & " WHERE venta_codigo = " & correlv & " "
'                    Case "UALMR", "ALMRS", "ALMRB", "ALMRC"
'                        db.Execute "UPDATE ao_ventas_cabecera SET doc_numero = " & VAR_NUM & " WHERE venta_codigo = " & correlv & " "
'                    Case "UALMH", "ALMHS", "ALMHB", "ALMHC"
'                        db.Execute "UPDATE ao_ventas_cabecera SET doc_numero = " & VAR_NUM & " WHERE venta_codigo = " & correlv & " "
'                    Case "DCONT", "GADM"
'                        db.Execute "UPDATE ao_ventas_cabecera SET doc_numero = " & VAR_NUM & " WHERE venta_codigo = " & correlv & " "
'                    Case Else
'                        MsgBox "No se ha definido la Unidad " & vbCrLf & " vuelva a intentar ... ", vbOKOnly + vbCritical, "Error de aprobación... "
'                        Exit Sub
'               End Select
''               'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
''                If ado_datos14.Recordset!doc_numero_m = 0 Or IsNull(ado_datos14.Recordset!doc_numero_m = 0) Then
''                    Set rs_aux2 = New ADODB.Recordset
''                    If rs_aux2.State = 1 Then rs_aux2.Close
''                    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "' "
''                    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''                    If rs_aux2.RecordCount > 0 Then
''                       rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''                       VAR_NUM = rs_aux2!correl_doc
''                       rs_aux2.Update
''
''                       db.Execute "Update to_cronograma_diario_final SET doc_numero_m = " & VAR_NUM & "  Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
''                       db.Execute "Update to_cronograma_diario_final SET fecha_almi = '" & Date & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
''                       db.Execute "Update to_cronograma_diario_final SET ok_almacen = 'True' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
''                       db.Execute "Update to_cronograma_diario_final SET observaciones2 = '" & txt_obs.Text & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
''
''                       ' Actualiza Saldos ac_bienes
''                       'RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
''                       db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = av_acumula_ventas_detalle.venta_det_cantidad from ac_bienes, av_acumula_ventas_detalle Where ac_bienes.grupo_codigo = av_acumula_ventas_detalle.grupo_codigo And ac_bienes.subgrupo_codigo = av_acumula_ventas_detalle.subgrupo_codigo And ac_bienes.bien_codigo = av_acumula_ventas_detalle.bien_codigo"
''                       db.Execute "update ac_bienes set bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida"
''                    End If
''                End If
'
''                ' GRABA Nombre de Archivo en ao_ventas_cabecera. VERIFICAR JQA 2014-07-08
''                'rs_datos!doc_numero = Txt_campo1.Caption
''                'VAR_ARCH = RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
''                VAR_ARCH = "TEC_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
''                db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
''                db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
''                db.Execute "update ao_solicitud set unidad_codigo_ant = '" & Ado_datos.Recordset!unidad_codigo_ant & "' Where unidad_codigo= '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " "
''               ' REVISAR JQ-2014-JUL-05
'                'INI HABILITA ALMACEN PARA venta_tipo="V" (PREVENTA)
'                'correlv = 2
'               'If VAR_TIPOV<> "V" Then
'               'If VAR_TIPOV <> "V" And VAR_TIPOV <> "C" Then
'
''                 Set rsAuxDetalle = New ADODB.Recordset
''                 If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
''                 rsAuxDetalle.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & " and almacen_codigo = " & VAR_ALMX & " and bien_cantidad_por_empaque > 0 ", db, adOpenKeyset, adLockBatchOptimistic
''                 If rsAuxDetalle.RecordCount > 0 Then
''                   rsAuxDetalle.MoveFirst
''                   While Not rsAuxDetalle.EOF   ' AdoAux.Recordset.EOF
''                     Set rs_almacen2 = New ADODB.Recordset
''                     If rs_almacen2.State = 1 Then rs_almacen2.Close
''                     rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = '" & rsAuxDetalle!almacen_codigo & "' and bien_codigo = '" & rsAuxDetalle!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
''                     If rs_almacen2.RecordCount > 0 Then
''                         db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida = " & rsAuxDetalle!bien_cantidad_por_empaque & "  from ao_almacen_totales, ao_ventas_detalle Where ao_almacen_totales.almacen_codigo = '" & rsAuxDetalle!almacen_codigo & "'   And ao_almacen_totales.bien_codigo = '" & rsAuxDetalle!bien_codigo & "'   "
''                         'AdoAux.Recordset.MoveNext
''                     Else
'''                        db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_salida) SELECT " & Ado_datos.Recordset!almacen_codigo & ", '" & rsAuxDetalle!bien_codigo & "', '" & rsAuxDetalle!grupo_codigo & "', '" & rsAuxDetalle!subgrupo_codigo & "', '" & rsAuxDetalle!par_codigo & "' , '" & rsAuxDetalle!bien_cantidad_por_empaque & "' FROM av_acumula_ventas_detalle WHERE almacen_codigo = '" & Ado_datos.Recordset!almacen_codigo & "'   And bien_codigo = '" & rsAuxDetalle!bien_codigo & "'    "
''                        db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_salida) SELECT av_acumula_ventas_detalle.almacen_codigo , av_acumula_ventas_detalle.bien_codigo , av_acumula_ventas_detalle.grupo_codigo, av_acumula_ventas_detalle.subgrupo_codigo, av_acumula_ventas_detalle.par_codigo  ,av_acumula_ventas_detalle.bien_cantidad_por_empaque  FROM av_acumula_ventas_detalle WHERE almacen_codigo = '" & Ado_datos.Recordset!almacen_codigo & "'   And bien_codigo = '" & rsAuxDetalle!bien_codigo & "'    "
'''                         'GRABA ALMACEN DETALLE
'''                        Set rs_aux4 = New ADODB.Recordset
'''                        If rs_aux4.State = 1 Then rs_aux4.Close
'''                        rs_aux4.Open "Select * from av_acumula_compras_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = '" & Ado_datos.Recordset!solicitud_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
'''                        'rs_aux4.Open "Select * from ao_almacen_totales where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
'''                        If rs_aux4.RecordCount > 0 Then
'''                            db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
'''                        Else
'''                            'If Ado_datos.Recordset!venta_tipo = "V" Then
'''                            '    'db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
'''                                db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) VALUES (" & rsAuxDetalle!almacen_codigo & ", '" & rsAuxDetalle!bien_codigo & "', '" & rsAuxDetalle!grupo_codigo & "', '" & rsAuxDetalle!subgrupo_codigo & "', '" & rsAuxDetalle!par_codigo & "' , " & rsAuxDetalle!venta_det_cantidad & ")"
'''                            'Else
'''                            '    MsgBox "Error Verifique la Adjudicación de Bienes (Equipos, Repuestos u otros) ..."
'''                            'End If
'''                        End If
''                     End If
''                     rsAuxDetalle.MoveNext
''                   Wend
''                   db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida "
''                 Else
''                    MsgBox "Error Verifique la Registro de Salida ..."
''                 End If
''               ' Actualiza Saldos ac_bienes
''               db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = av_acumula_ventas_detalle.bien_cantidad_por_empaque from ac_bienes, av_acumula_ventas_detalle Where ac_bienes.grupo_codigo = av_acumula_ventas_detalle.grupo_codigo And ac_bienes.subgrupo_codigo = av_acumula_ventas_detalle.subgrupo_codigo And ac_bienes.bien_codigo = av_acumula_ventas_detalle.bien_codigo"
''               db.Execute "update ac_bienes set bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida-bien_stock_salida_mant"
'
'               'End If
'               'FIN HABILITA ALMACEN PARA venta_tipo="V" (PREVENTA)
'               ' APRUEBA ao_ventas_cabecera
'               'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
'              db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_almacen = 'APR' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'               'marca1 = Ado_datos.Recordset.Bookmark
'               'Ado_datos.Recordset.Requery
'        '       Ado_datos.Refresh
'               'Ado_datos.Recordset.Move marca1 - 1
'               'Call Contabiliza_venta              ' JQA ENE-2016
'
'               Call OptFilGral1_Click
'           End If
'         End If
'       End If
'     Else
'        MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'     End If
'  End If
'  Exit Sub
'UpdateErr:
'    MsgBox Err.Description

On Error GoTo UpdateErr

If Ado_datos.Recordset!estado_almacen = "REG" Then

If rs_datos14.State = 1 Then rs_datos14.Close
rs_datos14.Open "select * from ao_ventas_detalle where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  and almacen_tipo = '" & VAR_ALMT & "'  order by  par_codigo, bien_codigo ", db, adOpenKeyset, adLockOptimistic
rs_datos14.MoveFirst
While Not rs_datos14.EOF
        
         Set rs_precio = New ADODB.Recordset
            If rs_precio.State = 1 Then rs_precio.Close
            rs_precio.Open "SELECT TOP 1 * from ao_almacen_ingresos where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & Ado_datos.Recordset!almacen_codigo & " and bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' ORDER BY fecha_ingreso DESC", db, adOpenKeyset, adLockOptimistic
            
            If rs_precio.RecordCount > 0 Then
                precio_uni = IIf(IsNull(rs_precio!precio_unitario_bs) Or rs_precio!precio_unitario_bs = "0", 1, rs_precio!precio_unitario_bs)
            Else
                precio_uni = 1
            End If
            precio_tot = 0
                db.Execute "ap_ventas_grla 2 ,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""

                If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
                
                db.Execute "ap_compras_grla 2,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo_d & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "', " & Ado_datos.Recordset!venta_codigo & ", '" & Ado_datos.Recordset!beneficiario_codigo_alm & "', '" & Ado_datos.Recordset!fecha_verif & "', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "', " & precio_tot & ""
                
             Set rs_almacen2 = New ADODB.Recordset
             
             If rs_almacen2.State = 1 Then rs_almacen2.Close
             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & Ado_datos.Recordset!almacen_codigo_d & " and bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
             If rs_almacen2.RecordCount > 0 Then
                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida = av_acumula_ventas_detalle.bien_cantidad_salida from ao_almacen_totales inner join av_acumula_ventas_detalle on ao_almacen_totales.almacen_codigo = av_acumula_ventas_detalle.almacen_codigo and ao_almacen_totales.bien_codigo = av_acumula_ventas_detalle.bien_codigo WHERE ao_almacen_totales.almacen_codigo = " & VAR_ALMX & " And ao_almacen_totales.bien_codigo = '" & VAR_BIEN2 & "'     "
                 db.Execute "ap_almacen_totales 2," & Ado_datos.Recordset!almacen_codigo_d & ", '" & ado_datos14.Recordset!bien_codigo & "', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", 0" & ", " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & precio_tot & ", 0, 0, " & precio_tot / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
             Else
                 db.Execute "ap_almacen_totales 1," & Ado_datos.Recordset!almacen_codigo_d & ", '" & ado_datos14.Recordset!bien_codigo & "', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", 0" & ", " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & precio_tot & ", 0, 0, " & precio_tot / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
            End If
            
                End If

            
            'ACtualiza ac_bienes    'Todos Los Almacenes
            db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = total_salidas_js.cantidad_salida from total_salidas_js Where ac_bienes.bien_codigo = total_salidas_js.bien_codigo"
            'db.Execute "update ac_bienes set bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida"
            db.Execute "update ac_bienes set bien_stock_actual = ISNULL(bien_stock_ingreso,0) - bien_stock_salida"
            
            'ACTUALIZA ao_almacen_totales   'Actualiza en el Almacen ORIGEN
            Set rs_almacen2 = New ADODB.Recordset
             If rs_almacen2.State = 1 Then rs_almacen2.Close
             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & Ado_datos.Recordset!almacen_codigo & " and bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
             If rs_almacen2.RecordCount > 0 Then
                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida = av_acumula_ventas_detalle.bien_cantidad_salida from ao_almacen_totales inner join av_acumula_ventas_detalle on ao_almacen_totales.almacen_codigo = av_acumula_ventas_detalle.almacen_codigo and ao_almacen_totales.bien_codigo = av_acumula_ventas_detalle.bien_codigo WHERE ao_almacen_totales.almacen_codigo = " & VAR_ALMX & " And ao_almacen_totales.bien_codigo = '" & VAR_BIEN2 & "'     "
                 db.Execute "ap_almacen_totales 2," & Ado_datos.Recordset!almacen_codigo & ", '" & ado_datos14.Recordset!bien_codigo & "', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", 0" & ", " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & precio_tot & ", 0, 0, " & precio_tot / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
             Else
                 db.Execute "ap_almacen_totales 1," & Ado_datos.Recordset!almacen_codigo & ", '" & ado_datos14.Recordset!bien_codigo & "', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", 0" & ", " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & precio_tot & ", 0, 0, " & precio_tot / GlTipoCambioOficial & ", 0, 0, 'REG','" & glusuario & "'"
            End If

rs_datos14.MoveNext
Wend
Ado_datos.Recordset!estado_almacen = "APR"
Ado_datos.Recordset.Update
Set rs_datos15 = New ADODB.Recordset
            If rs_datos15.State = 1 Then rs_datos15.Close
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'I' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
            Set ado_datos15.Recordset = rs_datos15
            ado_datos15.Refresh
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = False
            FraNavega.Enabled = True
            FrmDetalle.Enabled = True
            FrmABMDet.Visible = True
            FrmEdita.Enabled = False
Call AbrirDetalle
Else
MsgBox "No se puede aprobar el registro actual"
End If
Exit Sub
UpdateErr:
MsgBox Err.Description
End Sub


Private Sub GENERA_COMPRA()
'    If rs_datos!estado_cotiza = "REG" Then
'      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
'      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'        ' MANTENIMIENTO PREVENTIVO - INSUMOS y/o COMPRAS BB y SS
'                'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & parametro & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
'                       correldetalle = rs_aux2!correl_negocia
'                       rs_aux2.Update
'                    End If
'                    'WWWWWWWWWWWWWWW
'                    'correlv = Ado_datos.Recordset!venta_codigo
'                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = parametro
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_COD4
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = Ado_datos.Recordset!solicitud_tipo       '"10"
'                        rs_aux3!venta_tipo = "E"
'                        rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_descripcion = "COMPRA POR: " + lbl_titulo.Caption
'                        rs_aux3!compra_observaciones = "Edificio: " + Trim(dtc_desc3.Text)
'                        rs_aux3!compra_cantidad_total = 1   'Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = 0     'VAR_BS2
'                        rs_aux3!tipo_moneda = "BOB"
'                        rs_aux3!compra_monto_dol = 0        'VAR_DOL2
'                        rs_aux3!proceso_codigo = "TEC"
'                        rs_aux3!subproceso_codigo = "TEC-06"
'                        rs_aux3!etapa_codigo = "TEC-06-01"
'                        rs_aux3!clasif_codigo = "ADM"
'                        rs_aux3!doc_codigo = "R-114"
'                        rs_aux3!poa_codigo = "3.2.8"
'                        rs_aux3!estado_codigo_eqp = "REG"
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'
'                        'DETALLE Carga ao_ventas_detalle
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rstdestino.RecordCount > 0 Then
'                        End If
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        'rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & rs_aux3!compra_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo= " & VAR_SOL & "  and grupo_codigo = '30000' ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            VAR_REG = 1
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                              If rs_aux4!grupo_codigo = "30000" Then
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'
'                                db.Execute "Update ao_compra_detalle SET ao_compra_detalle.compra_concepto  = ac_bienes.bien_descripcion From ao_compra_detalle INNER JOIN ac_bienes ON ao_compra_detalle.bien_codigo = ac_bienes.bien_codigo where ao_compra_detalle.compra_codigo = " & rs_aux3!compra_codigo & " and ao_compra_detalle.bien_codigo = '" & rs_aux4!bien_codigo & "' "
'                                VAR_REG = VAR_REG + 1
'                              End If
'                               rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                    End If
'                    'WWWWWWWWWW
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_cotiza = "APR"
'        rs_datos!fecha_aprueba = Date
'        rs_datos!usr_codigo_aprueba = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
End Sub

'Private Sub BtnAprobar2_Click()
' If IsNull(Ado_datos16.Recordset("cobranza_observaciones")) Or (Ado_datos16.Recordset("cobranza_programada_bs") = 0) Or Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
'    'If Ado_datos16.Recordset!beneficiario_codigo_resp = "" Or IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp) Then
'    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'    Exit Sub
' Else
'    If Ado_datos.Recordset("estado_codigo") = "REG" Then
'        MsgBox "No se puede APROBAR el registro (Cronograma), previamente debe APROBAR la Venta (Cabecera) y vuelva a intentar ...", , "Atención"
'        Exit Sub
'    End If
'    If Ado_datos16.Recordset("estado_codigo") = "REG" Then
'       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'            db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & Ado_datos.Recordset!venta_codigo & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
'            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo, venta_codigo, beneficiario_codigo, beneficiario_codigo_fac, beneficiario_codigo_resp, cobranza_programada_bs, cobranza_programada_dol, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, Literal, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo, estado_codigo, usr_codigo, fecha_registro) " & _
'            "VALUES ('" & glGestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & Ado_datos16.Recordset!venta_codigo & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo_resp & "', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '0', '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_observaciones & "', 'FIN', 'FIN-01', 'FIN-01-02', 'ADM', 'R-105', '0', 'R-101', '0', '0', '3.1.2', 'REG', '" & glusuario & "', '" & Date & "')"
'
''            Set rs_aux1 = New ADODB.Recordset
''            If rs_aux1.State = 1 Then rs_aux1.Close
''            rs_aux1.Open "select * from ao_ventas_cobranza where venta_codigo= " & Ado_datos16.Recordset!venta_codigo & "  and cobranza_prog_codigo= " & Ado_datos16.Recordset!cobranza_prog_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
''            If rs_aux1.RecordCount <= 0 Then
''                rs_aux1.AddNew
''            End If
''                rs_aux1!ges_gestion = Ado_datos16.Recordset!ges_gestion
''                rs_aux1!cobranza_prog_codigo = Ado_datos16.Recordset!cobranza_prog_codigo
''                rs_aux1!venta_codigo = Ado_datos16.Recordset!venta_codigo
''                rs_aux1!beneficiario_codigo = Ado_datos16.Recordset!beneficiario_codigo                 'Codigo Beneficiario/Cliente
''                rs_aux1!beneficiario_codigo_resp = Ado_datos16.Recordset!beneficiario_codigo_resp       'Codigo Cobrador
''
''                rs_aux1!cobranza_programada_bs = Ado_datos16.Recordset!cobranza_programada_bs           'Monto Programado Bs
''                rs_aux1!cobranza_programada_dol = Ado_datos16.Recordset!cobranza_programada_dol         'Monto Programado en Dolares
''                rs_aux1!cobranza_deuda_bs = Ado_datos16.Recordset!cobranza_programada_bs                'Monto Cobrado
''                rs_aux1!cobranza_deuda_dol = Ado_datos16.Recordset!cobranza_programada_dol              'Monto en Dolares
''                rs_aux1!cobranza_descuento_bs = 0     'Ado_datos16.Recordset!cobranza_descuento_bs      'Descuento Bs
''                rs_aux1!cobranza_descuento_dol = 0    'Ado_datos16.Recordset!cobranza_descuento_dol     'Descuento Dol
''                rs_aux1!cobranza_total_bs = Ado_datos16.Recordset!cobranza_programada_bs                'Monto Total Bs
''                rs_aux1!cobranza_total_dol = Ado_datos16.Recordset!cobranza_programada_dol              'Monto Total Dol
''                rs_aux1!Literal = Ado_datos16.Recordset!Literal
''                rs_aux1!cobranza_fecha_prog = Ado_datos16.Recordset!cobranza_fecha_prog                 'Fecha de Programada
''                rs_aux1!cobranza_fecha_cobro = Ado_datos16.Recordset!cobranza_fecha_prog                'Fecha de Cobranza
''
''                rs_aux1!cobranza_observaciones = Ado_datos16.Recordset!cobranza_observaciones
''                rs_aux1!proceso_codigo = "COM"
''                rs_aux1!subproceso_codigo = "COM-02"
''                rs_aux1!etapa_codigo = "COM-02-04"
''                rs_aux1!clasif_codigo = "ADM"
''                rs_aux1!doc_codigo = "R-103"
''                rs_aux1!doc_numero = rs_aux1.RecordCount
''                rs_aux1!doc_codigo_fac = ""
''                rs_aux1!cobranza_nro_factura = "0"
''                rs_aux1!cobranza_nro_autorizacion = "0"
''                rs_aux1!poa_codigo = "3.1.2"
''                rs_aux1!estado_codigo = "REG"
''                rs_aux1!usr_codigo = GlUsuario
''                rs_aux1!fecha_registro = Format(Date, "dd/mm/yyyy")
''                rs_aux1!hora_registro = Format(Time, "hh:mm:ss")
''                rs_aux1.Update
'            ' APRUEBA ao_ventas_cobranza_prog
'            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
'            'db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
'            Ado_datos16.Refresh
'       End If
'    End If
' End If
'End Sub

Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
    'JQA
    '  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
    '  Dim ClBuscaSec As ClBuscaSecuencialEnRS
      buscados = 1
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
    OptFilGral1.Visible = True
    OptFilGral2.Visible = True
  End If
End Sub

Private Sub BtnCancelar_Click()
On Error GoTo UpdateErr
  If swgrabar = 2 Then
    var_cod5 = Ado_datos.Recordset!venta_codigo
  End If
  'Ado_datos.Refresh
  fraOpciones.Visible = True
  FraGrabarCancelar.Visible = False
  'marca1 = Ado_datos.Recordset.Bookmark
  FraNavega.Enabled = True
  FrmCabecera.Enabled = False
  Fra_datos.Enabled = True
  FrmDetalle.Visible = True

'  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  dtc_desc3.backColor = &H80000008
  dtc_desc3.ForeColor = &H80000005
  
  'Refrescar Grid
  If OptFilGral1.Value = True Then
       Call OptFilGral1_Click        'Pendientes
  Else
       Call OptFilGral2_Click        'TODOS
  End If
  If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
  End If
  If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Then
       rs_datos.Find "venta_codigo = " & var_cod5 & "   ", , , 1
       dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
  Else
       rs_datos.MoveLast
  End If
  swgrabar = 0
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub

Private Sub BtnCancelar3_Click()
        Fra_reporte.Visible = False
End Sub

Private Sub BtnEliminar_Click()
'  If Ado_datos.Recordset.RecordCount > 0 Then
'    If Ado_datos.Recordset("estado_almacen") = "REG" Then
'      sino = MsgBox("Esta seguro de ANULAR la venta registrada ?", vbYesNo, "Confirmando")
'      If sino = vbYes Then
'          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'ANL' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  "
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
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_almacen = "REG" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
'     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
          rs_datos!estado_almacen = "ANL"
'          rs_datos!fecha_registro = Date
'          rs_datos!usr_codigo = glusuario
'           Ado_datos.Recordset.Requery
'           Ado_datos.Refresh
           db.Execute "ap_ventas_grla 1 ,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
           Call AbrirDetalle
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()

  If dtc_codigo2 = "" Then
    MsgBox "Debe Elejir La Unidad Destino, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  If dtc_codigo4 = "" Then
    MsgBox "Debe Elejir Responsable Almacen ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo11 = "" Then
    MsgBox "Debe Elejir el Almacen!! , Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo5 = "" Then
    MsgBox "Debe Elejir ... Entregado a:, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3 = "" Then
    MsgBox "Debe Registrar el Edificio / Destino, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo21 = "" Then
    MsgBox "Debe Elejir Regional ORIGEN, Vuelva a Intentar ...", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo UpdateErr
On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If swgrabar = 2 Then
        var_cod5 = Ado_datos.Recordset!venta_codigo
        FInicio = IIf(IsNull(Ado_datos.Recordset!venta_fecha_inicio), Date, Ado_datos.Recordset!venta_fecha_inicio)
        CANTOT = IIf(IsNull(Ado_datos.Recordset!venta_cantidad_total), 1, Ado_datos.Recordset!venta_cantidad_total)
        gestion0 = IIf(IsNull(Ado_datos.Recordset!ges_gestion), glGestion, Ado_datos.Recordset!ges_gestion)
        VAR_BENEF = IIf(IsNull(Ado_datos.Recordset!beneficiario_codigo), "0", Ado_datos.Recordset!beneficiario_codigo)
        corrprog = Ado_datos.Recordset!correl_cobro_prog
        VAR_MED = Ado_datos.Recordset!unimed_codigo
        VAR_UNI = Ado_datos.Recordset!unidad_codigo
        FControl = IIf(IsNull(Ado_datos.Recordset!fecha_verif), Date, Ado_datos.Recordset!fecha_verif)
        'Ado_datos.Recordset("fecha_verif") = DTPfechasol.Value
        '        rs_datos!fecha_verif = Date
    End If
    FrmCabecera.Enabled = False
    Call grabar
    '
    db.Execute "update ao_almacen_salidas set concepto = '" & TxtConcepto.Text & "' WHERE venta_codigo = " & var_cod5
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraNavega.Enabled = True
    FrmCabecera.Enabled = False
    Fra_datos.Enabled = True
    dg_datos.Visible = True
    FrmDetalle.Visible = True
    dtc_desc3.backColor = &H80000008
    dtc_desc3.ForeColor = &H80000005
'    Fra_Total.Visible = True
    FrmABMDet.Visible = True
    'Refrescar Grid
    If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And swgrabar = 2 Or swgrabar = 1 Then
        rs_datos.Find "venta_codigo = " & var_cod5 & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     swgrabar = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
  End If
    accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description

End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        CryV01.Reset
        CryV01.WindowState = crptMaximized
        CryV01.WindowShowSearchBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.WindowShowPrintSetupBtn = True
        
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        Call CARGAPARAM
        If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
            CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_trfi.rpt"
            var_titulo = "NOTA DE TRASPASO"
        Else
            Select Case VAR_BIEN
                Case "INSUMOS"
                    CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes.rpt"
                Case "REPUESTOS"
                    CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_repuestos.rpt"
                Case "HERRAMIENTAS"
                    CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_herramientas.rpt"
                Case "ADMINISTRACION"
                    'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes.rpt"
            End Select
            var_titulo = "SALIDA DE ALMACENES"
        End If
        
'        If VAR_BIEN = "INSUMOS" Then
'            CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes.rpt"
'        End If
'        If Left(Ado_datos.Recordset!doc_codigo_alm, 5) = "R-115" Then
'            CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes.rpt"
'        Else
'            CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_trf.rpt"
'        End If
        
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
        'var_titulo = "MODULO ALMACENES"
        CryV01.Formulas(0) = "titulo = '" & var_titulo & "' "
        
        CryV01.Formulas(1) = "subtitulo = '" & "ALMACEN DE " & "' + '" & VAR_BIEN & "' "
        'If Ado_datos.Recordset!unidad_codigo = "UALMI" Or Ado_datos.Recordset!unidad_codigo = "UALMR" Or Ado_datos.Recordset!unidad_codigo = "UALMH" Then
        '    CryV01.Formulas(1) = "subtitulo = '" & "SALIDA ALMACEN DE " & "' + '" & VAR_BIEN & "' "
        'Else
        '    CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
        'End If
        'cr01.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    
'    If Ado_datos.Recordset.RecordCount > 0 Then
'        Dim iResult As Variant, i%, Y%
'        Dim co As New ADODB.Command
'
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
'          Case "DVTA"
'              var_titulo = "Módulo Comercial"
'        End Select
'
'        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas.rpt"
'        CryV01.WindowShowPrintSetupBtn = True
'        CryV01.WindowShowRefreshBtn = True
'        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
'
'        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
'        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
'        iResult = CryV01.PrintReport
'        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'    Else
'        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'    End If
End Sub


Private Sub BtnImprimir1_Click()
    Fra_reporte.Visible = True
'   If Ado_datos.Recordset.RecordCount > 0 Then
'      If ado_datos14.Recordset.RecordCount > 0 Then
'        Dim iResult As Variant, i%, Y%
'        Dim co As New ADODB.Command
'
'    '    Dim rs As New ADODB.Recordset
'    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
'    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
'    '    i = 1
'    '    y = 1
'        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_nota_de_venta.rpt"
'        CryV01.WindowShowRefreshBtn = True
'        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'        CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
'        CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
'        iResult = CryV01.PrintReport
'        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
'     Else
'        MsgBox "No se puede Imprimir. Debe registrar datos... " & FrmDetalle.Caption, , "Atención"
'     End If
'   Else
'        MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
'   End If

End Sub

Private Sub BtnModificar_Click()
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset("estado_almacen") = "REG" Then
        accion = "MOD"
        If dtc_desc4.Text = "" Or dtc_desc11.Text = "" Or dtc_desc21.Text = "" Then
            Fra_datos.Enabled = True
        Else
            Fra_datos.Enabled = False
        End If
        FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
        FraNavega.Enabled = False
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        'Fra_datos.Enabled = True
'        Fra_Total.Visible = False
        FrmABMDet.Visible = False
        swgrabar = 2
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        'If Ado_datos.Recordset!unidad_codigo = "UALMI" Or Ado_datos.Recordset!unidad_codigo = "UALMR" Or Ado_datos.Recordset!unidad_codigo = "UALMH" Or Ado_datos.Recordset!unidad_codigo = "DADM" Then
        If VAR_ORIGEN = "UALMI" Then        'Ado_datos.Recordset!unidad_codigo
            dtc_desc3.Locked = False
            dtc_desc3.Width = 5955
            'TxtConcepto.Locked = False
        Else
            dtc_desc3.Width = 6315
            dtc_desc3.Locked = True
            'TxtConcepto.Locked = True
        End If
    Else
      MsgBox "NO se puede MODIFICAR, porque el registro ya fue Aprobado, Anulado o Cerrado.", , "Atencion"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  
  Exit Sub
UpdateErr:
    MsgBox Err.Description
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
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  If Ado_datos.Recordset("estado_codigo") = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
'    FrmABMDet2.Visible = True
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroVenta.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  swnuevo = 0
  'cmdElige.Enabled = False
  marca1 = Ado_datos.Recordset.Bookmark
'  If Ado_datos.Recordset("estado_almacen") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    FrmABMDet.Visible = True
    
'     Call AbrirDetalle
  ado_datos14.Recordset.Cancel
  Call AbrirDetalle
  'Ado_datos.Recordset.Move marca1 - 1
  accion = ""
End Sub

Private Sub BtnAnlDetalle2_Click()
 If Ado_datos.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
      db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & Ado_datos16.Recordset("cobranza_codigo") & " "
      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_deuda_bs = '0', ao_ventas_cobranza_prog.cobranza_deuda_dol = '0'  Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & ado_datos16.Recordset("cobranza_codigo") & " "

     'ado_ventas_COBRANZAS.Recordset.Delete
     'ado_ventas_COBRANZAS.Recordset.Update
     'ado_ventas_COBRANZAS.Requery
     'ado_ventas_COBRANZAS.Refresh
     ''cerea
     'ado_ventas_COBRANZAS.Refresh
   End If
  Else
    MsgBox "Los productos del registro sin Aprobar, NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  'If Ado_datos.Recordset!venta_tipo <> "E" And Ado_datos16.Recordset!estado_codigo = "REG" Then
  If Ado_datos16.Recordset!estado_codigo = "REG" And (Ado_datos.Recordset!venta_tipo = "E" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "C") Then
    marca1 = Ado_datos16.Recordset.Bookmark
    FraNavega.Enabled = False
    fraOpciones.Enabled = False
    FrmDetalle.Visible = False
'    FrmCobranza.Visible = False
    VAR_COBR1 = Ado_datos16.Recordset!cobranza_prog_codigo
    'swgrabar = 0
    swnuevo = 2
    TxtCobrador.Visible = False
    'TxtMonto.SetFocus
    'TxtNroVenta.Enabled = False
    'marca1 = ado_datos14.Recordset.BookMark
    'txt_descripcion_venta.Enabled = True
    'TxtNroVenta.Text = txt_venta.Text
    'lbltipoVenta.Caption = dtc_desc11.Text
    'lblges_gestion.Caption = Ado_datos.Recordset!ges_gestion
    SSTab1.Tab = 2
'    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    FrmCobros.Visible = True
    FrmCobros.Enabled = True
    FrmABMDet.Visible = False
'    FrmABMDet2.Visible = False
    'If Ado_datos.Recordset!estado_codigo = "APR" Then
        'sino = MsgBox("Registrará la cobranza efectiva, ahora ? ", vbYesNo, "Confirmando")
        'If sino = vbYes Then
        '    DTPFechaProg.Visible = False
        '    DTPFechaCobro.Visible = True
        '    Lbl_nombre_fac.Caption = "Factura a Nombre de:"
        '    lbl_fechas.Caption = "Fecha Efectiva de Cobranza"
        '    Txt_parche.Visible = False      '&H80000013&
        '    'dtc_desc2A.BackColor = &H80000013
        'Else
        '    DTPFechaProg.Visible = True
        '    DTPFechaCobro.Visible = False
        '    Lbl_nombre_fac.Caption = "Cliente :"
        '    lbl_fechas.Caption = "Fecha Programada de Cobranza"
        '    Txt_parche.Visible = True       '&H80000005&
        '    'dtc_desc2A.BackColor = &H80000005
        'End If
    'Else
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = False
        DTPFechaConf.Visible = True
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de Cobranza"
'        Txt_parche.Visible = True       '&H80000005&
        'dtc_desc2A.BackColor = &H80000005
    'End If
    VAR_MBS2 = Ado_datos16.Recordset!cobranza_programada_bs
    TxtMonto.SetFocus
'    Call ABRIR_TABLA_DET
'    Ado_datos16.Recordset.Move marca1 - 1
  Else
    MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos16.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Then
    If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 2
'        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmABMDet.Visible = False
'        FrmABMDet2.Visible = False
        TxtCobrador.Visible = False
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = False
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        'Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
'        Dim thisDate As Date
'        Dim thisMonth As Integer
'        thisDate = #2/12/1969#
'        thisMonth = Month(thisDate)
'        ' thisMonth now contains 2.
'
'
'        Dim thisMonth As Integer
'        Dim name As String
'        thisMonth = 4
'        ' Set Abbreviate to True to return an abbreviated name.
'        name = MonthName(thisMonth, True)
'        ' name now contains "Apr".
    Else
        MsgBox "Ya se cobró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (al Contado o Donación) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnDesAprobar_Click()
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
'
'  If rs_datos!estado_almacen = "APR" Then
'       sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'          rs_datos!estado_almacen = "REG"
'          rs_datos.UpdateBatch adAffectAll
'       End If
'    Else
'       MsgBox "No se puede HABILITAR (Desaprobar) un registro ya Habilitado o Anulado ...", vbExclamation, "Validación de Registro"
'    End If

On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_almacen = "APR" Then
       sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
'     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
         
'          rs_datos!fecha_registro = Date
'          rs_datos!usr_codigo = glusuario
'           Ado_datos.Recordset.Requery
'           Ado_datos.Refresh
          db.Execute "ap_ventas_grla 3,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
            rs_datos!estado_almacen = "REG"
           Call AbrirDetalle
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede DESPROBAR un registro Aulado(ANL) o Registrado (REG) ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede DESAPROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
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

Private Sub cmdElige_Click()
'  With ALFrmMateriales
'        .ALPrincipal
'        If .QResp Then
'            TxtCodigo.Text = .QCodigo
'            txtDesc.Text = .QItem
'        End If
'    End With
'    Txtcant_alm = 0
'    Cant_Alm = 0
'    DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
'    Txtcant_alm = Cant_Alm
'    If Cant_Alm >= TxtCantPedi Then
'        optSi = True
'    Else
'        optNo = True
'    End If
End Sub

Private Sub Contabiliza_venta()
    Call graba_proyecto
    Call graba_ingreso
  '===== Proceso para generar Asientos Contables Automáticos "DEI" y "REC"
  'sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  'If sino = vbYes Then
    ' INI CORRECCION 18-JUN-2014
    Dim i As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)

    '**** INI VERIFICAR VALIDACION REC, DES, ANI Y DVI !!! ***************
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    Select Case VAR_CODTIPO
        Case "DEI"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
              'cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
              'Subcta_deb11 = rstdestino!Subcta_cred1
              'Subcta_deb21 = rstdestino!Subcta_cred2

              'cta_credito1 = rstdestino2!cta_deb
              'Subcta_cred11 = rstdestino2!Subcta_deb1
              'Subcta_cred21 = rstdestino2!Subcta_deb2
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

        Case "REC"
            rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
            If rstdestino.RecordCount > 0 Then
                j = rstdestino.RecordCount
            Else
              MsgBox "Este comprobante no puede ser procesado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
              Exit Sub
            End If

            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select * from fo_ingresos_cabecera where ingreso_codigo = " & VAR_CODANT & " and org_codigo = '" & VAR_ORG & "' ", db, adOpenKeyset, adLockOptimistic
            If (Not rs_aux1.BOF) And (Not rs_aux1.EOF) Then
              If rs_aux1("monto_bolivianos") < rs_aux1("monto_recaudado_bolivianos") + VAR_BS2 Then
                MsgBox "El monto que está intentando recaudar en Bs. es mayor al DEVENGADO, por favor Verifique el Monto Devengado: " & CStr(rs_aux1("monto_bolivianos")) & " Solo puede recaudar :" & CStr(rs_aux1("monto_bolivianos") - rs_aux1("monto_recaudado_bolivianos")), vbOKOnly + vbCritical, "ERROR en el Monto Recaudado"
                Exit Sub
              End If
            End If
            If rs_aux1.State = 1 Then rs_aux1.Close

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

    'If DtCCta_codigo.Text <> "01" Then
    '  If rstdestino.State = 1 Then rstdestino.Close
    '  rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
    '  If Not rstFc_cuenta_bancaria.EOF Then
    '    fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
    '  Else
    '  End If
    'Else
    '    fte_codigo1 = Me.DtCFte_codigo.Text
    'End If
    'If VAR_CODTIPO = "DEI" Or VAR_CODTIPO = "DES" Then
    '  fte_codigo1 = Me.DtCFte_codigo.Text
    'End If

'    fte_codigo1 = VAR_FTE
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'
'    v_Tipo_Comp(1, 1) = VAR_CODTIPO

'    If VAR_CODTIPO = "DYR" Then
'      'j = 2
'      'v_Tipo_Comp(1, 1) = "CAD"
'      'v_Tipo_Comp(1, 2) = "CAR"
'      j = 2
'      v_Tipo_Comp(1, 1) = "DYR"
'    Else
'      j = 1
'      v_Tipo_Comp(1, 1) = IIf(VAR_CODTIPO = "DEI", "DEI", IIf(VAR_CODTIPO = "REC", "REC", IIf(VAR_CODTIPO = "DES", "DES", IIf(VAR_CODTIPO = "ANI", "ANI", ""))))
'    End If
'
'    If VAR_CODTIPO = "DVI" Then
'      j = 1
'      v_Tipo_Comp(1, 1) = "DVI"
'    End If

'    For i = 1 To j
'      If rstdestino.State = 1 Then rstdestino.Close
'      If v_Tipo_Comp(1, i) = "DEI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DVI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'        Exit Sub
'      End If

    ' INI CORRECCION 18-JUN-2014
'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' 02/07/2014 VERIFICAR
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rs_aux2.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'          Exit Sub
'        End If
'      End If
'
'      If rs_aux2.RecordCount < 1 Then
'        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'        Exit Sub
'      End If
'    Next

    'If rstdestino.State = 1 Then rstdestino.Close

    fte_codigo1 = VAR_FTE
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
      If rs_aux2.RecordCount > 0 Then
        yacontabilizo = 1
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
        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'CMBTE'", db, adOpenDynamic, adLockOptimistic
        If rstCodComp.RecordCount > 0 Then
          Var_Comp = CDbl(rstCodComp!numero_correlativo)
          Var_Comp = Var_Comp + 1
          rstCodComp!numero_correlativo = Trim(Str(Var_Comp))
          rstCodComp.Update
        End If
        If rstCodComp.State = 1 Then rstCodComp.Close
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
      rs_aux2("ges_gestion") = glGestion    'Year(Date)
      'rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
      If yacontabilizo = 0 Then
        rs_aux2("Fecha_transacion") = Date
      End If
      rs_aux2("beneficiario_codigo") = VAR_BENEF
      rs_aux2("glosa") = VAR_GLOSA
      rs_aux2("unidad_codigo") = VAR_COD4       'Ado_datos.Recordset("unidad_codigo")
      rs_aux2("solicitud_codigo") = Ado_datos.Recordset("solicitud_codigo")
      rs_aux2("tipo_moneda") = VAR_MONEDA
      rs_aux2("unidad_codigo_ant") = VAR_CITE

      rs_aux2("proceso_codigo") = "FIN"
      rs_aux2("subproceso_codigo") = "FIN-02"
      Select Case VAR_CODTIPO
        Case "DEI"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "REC"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DYR"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "DES"
            rs_aux2("etapa_codigo") = "FIN-02-01"
        Case "ANI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
        Case "DVI"
            rs_aux2("etapa_codigo") = "FIN-02-02"
      End Select

      rs_aux2("clasif_codigo") = "ADM"
      rs_aux2("doc_codigo") = "R-128"
      rs_aux2("doc_numero") = Var_Comp
      rs_aux2("pro_codigo_det") = VAR_PROY2

      rs_aux2("estado_codigo") = "APR"

      If yacontabilizo = 0 Then
        rs_aux2("usr_codigo") = glusuario
        rs_aux2("Fecha_registro") = Format(Date, "dd/mm/yyyy")
        rs_aux2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If
      rs_aux2.Update
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
'      If v_Tipo_Comp(1, i) = "DEI" Then     'Devengado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "REC" Then     'Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DYR" Then     'Devengado y Recaudado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DYR' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then     'Desafectado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DES' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then     'Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DVI" Then     'Desafectado y Anulado
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'ANI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If

'      If v_Tipo_Comp(1, i) = "DVI" Then
'        ' VERIFICAR SI SE ESTA CONTROLANDA con el DYR
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'DEI' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where Codigo_Tipo = 'REC' and rubro_codigo_I <= " & (VAR_PARTIDA) & " and rubro_codigo_F >= " & (VAR_PARTIDA) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10" Or fte_codigo1 = "20", "01", IIf(fte_codigo1 = "30", "02", IIf(fte_codigo1 = "40" Or fte_codigo1 = "50", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
'          cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
'          Subcta_deb11 = rstdestino!Subcta_cred1
'          Subcta_deb21 = rstdestino!Subcta_cred2
'
'          cta_credito1 = rstdestino2!cta_deb
'          Subcta_cred11 = rstdestino2!Subcta_deb1
'          Subcta_cred21 = rstdestino2!Subcta_deb2
'        Else
'          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''          Exit Sub
'        End If
'      End If
'
'      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'        cta_credito1 = rstdestino("cta_cred")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
'      Else
'        'MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'        'Exit Sub
'
'      End If
      '2115
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")

        cta_credito1 = rstdestino("cta_cred")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
      Else
        cta_deb1 = rstdestino!cta_cred         'rstdestino!cta_credito
        Subcta_deb11 = rstdestino!Subcta_cred1
        Subcta_deb21 = rstdestino!Subcta_cred2

        cta_credito1 = rstdestino!cta_deb
        Subcta_cred11 = rstdestino!Subcta_deb1
        Subcta_cred21 = rstdestino!Subcta_deb2
      End If

      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        d_cta_nombre_1 = rs_aux1("NombreCta")
        d_aux1_1 = rs_aux1("aux1")
        d_aux2_1 = rs_aux1("aux2")
        d_aux3_1 = rs_aux1("aux3")
      End If
      If rs_aux1.State = 1 Then rs_aux1.Close
      rs_aux1.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rs_aux1.RecordCount > 0 Then
        h_cta_nombre_1 = rs_aux1("NombreCta")
        h_aux1_1 = rs_aux1("aux1")
        h_aux2_1 = rs_aux1("aux2")
        h_aux3_1 = rs_aux1("aux3")
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
      If (VAR_CODTIPO = "DEI") Or (VAR_CODTIPO = "REC") Or (VAR_CODTIPO = "DYR") Then
        rstdestino2("D_Cuenta") = cta_deb1
        rstdestino2("D_Nombre") = Trim(d_cta_nombre_1) ' CAMPO PARA ELIMINAR
        rstdestino2("D_Subcta1") = Subcta_deb11
        rstdestino2("D_SubCta2") = Subcta_deb21
        rstdestino2("D_Aux1") = d_aux1_1
        rstdestino2("D_Aux2") = d_aux2_1
        rstdestino2("D_Aux3") = d_aux3_1
        ' para Aux1
'        Select Case d_aux1_1
'                Case "01"
'                    VAR_COD1 = VAR_BENEF
'                Case "02"
'                    VAR_COD1 = VAR_CTA
'                Case "03"
'                    VAR_COD1 = VAR_PROY2
'                Case "04"
'                    VAR_COD1 = Ado_datos.Recordset("unidad_codigo")
'                Case "05"
'                    VAR_COD1 = ""
'                Case "06"
'                    VAR_COD1 = ""
'                Case "07"
'                    VAR_COD1 = ""
'                Case "08"
'                    VAR_COD1 = ""
'                Case "09"
'                    VAR_COD1 = VAR_ORG
'                Case "10"
'                    VAR_COD1 = ""
'                Case "11"
'                    VAR_COD1 = ""
'                Case "12"
'                    VAR_COD1 = ""
'        End Select
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
                rstdestino2("D_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                rstdestino2("D_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                rstdestino2("D_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = VAR_DPTO
                rstdestino2("D_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                rstdestino2("D_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
        End Select

        Select Case d_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                rstdestino2("D_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                rstdestino2("D_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                rstdestino2("D_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = VAR_DPTO
                rstdestino2("D_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                rstdestino2("D_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
        End Select

        Select Case d_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                rstdestino2("D_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                rstdestino2("D_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                rstdestino2("D_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = VAR_DPTO
                rstdestino2("D_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                rstdestino2("D_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        ' CORREGIR MONTOS JQA 2014-JUL-08
        If j > 1 Then
            If i = 1 Then
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
        End If
        rstdestino2("D_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
        'AQUI MONEDA 02/07/01
        'rstdestino2("D_Cambio") = GlTipoCambioMercado
        'AAAAAAAAAAAAAAQQQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUIIIIIIIIIIIII JQA
        rstdestino2("H_Cuenta") = cta_credito1
        rstdestino2("H_Nombre") = Trim(h_cta_nombre_1) ' CAMPO PARA ELIMINAR
        rstdestino2("H_SubCta1") = Subcta_cred11
        rstdestino2("H_SubCta2") = Subcta_cred21
        rstdestino2("H_Aux1") = h_aux1_1
        rstdestino2("H_Aux2") = h_aux2_1
        rstdestino2("H_Aux3") = h_aux3_1
        'rstdestino2("H_Cta_Aux1") = ""
        Select Case h_aux1_1
            Case "01"
                rstdestino2("H_Cta_Aux1") = VAR_BENEF
                rstdestino2("H_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                rstdestino2("H_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                rstdestino2("H_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = VAR_DPTO
                rstdestino2("H_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                rstdestino2("H_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
        End Select

        Select Case h_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                rstdestino2("H_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                rstdestino2("H_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                rstdestino2("H_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = VAR_DPTO
                rstdestino2("H_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                rstdestino2("H_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
        End Select

        Select Case h_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                rstdestino2("H_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                rstdestino2("H_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                rstdestino2("H_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = VAR_DPTO
                rstdestino2("H_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                rstdestino2("H_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
        End Select

'        If h_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If j > 1 Then
            If i = 1 Then
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
            Else
                rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
                rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
            End If
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2))
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2))
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado    'GlTipoCambioMercado
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
                rstdestino2("D_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux1") = VAR_CTA
                rstdestino2("D_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux1") = VAR_PROY2
                rstdestino2("D_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "06"
                rstdestino2("D_Cta_Aux1") = VAR_DPTO
                rstdestino2("D_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "08"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "09"
                rstdestino2("D_Cta_Aux1") = VAR_ORG
                rstdestino2("D_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "11"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "12"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
            Case "00"
                rstdestino2("D_Cta_Aux1") = ""
                rstdestino2("D_Des_Aux1") = ""
        End Select

        Select Case h_aux2_1
            Case "01"
                rstdestino2("D_Cta_Aux2") = VAR_BENEF
                rstdestino2("D_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux2") = VAR_CTA
                rstdestino2("D_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux2") = VAR_PROY2
                rstdestino2("D_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "06"
                rstdestino2("D_Cta_Aux2") = VAR_DPTO
                rstdestino2("D_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "08"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "09"
                rstdestino2("D_Cta_Aux2") = VAR_ORG
                rstdestino2("D_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "11"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "12"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
            Case "00"
                rstdestino2("D_Cta_Aux2") = ""
                rstdestino2("D_Des_Aux2") = ""
        End Select

        Select Case h_aux3_1
            Case "01"
                rstdestino2("D_Cta_Aux3") = VAR_BENEF
                rstdestino2("D_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("D_Cta_Aux3") = VAR_CTA
                rstdestino2("D_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("D_Cta_Aux3") = VAR_PROY2
                rstdestino2("D_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("D_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("D_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "06"
                rstdestino2("D_Cta_Aux3") = VAR_DPTO
                rstdestino2("D_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "08"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "09"
                rstdestino2("D_Cta_Aux3") = VAR_ORG
                rstdestino2("D_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "11"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "12"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
            Case "00"
                rstdestino2("D_Cta_Aux3") = ""
                rstdestino2("D_Des_Aux3") = ""
        End Select
'        If h_aux1_1 = "01" Then
'          rstdestino2("D_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino2("D_Cta_Aux1") = VAR_CTA
'        End If
'        rstdestino2("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        If i = 1 Then
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("D_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("D_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
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
                rstdestino2("H_Des_Aux1") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux1") = VAR_CTA
                rstdestino2("H_Des_Aux1") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux1") = VAR_PROY2
                rstdestino2("H_Des_Aux1") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux1") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux1") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "06"
                rstdestino2("H_Cta_Aux1") = VAR_DPTO
                rstdestino2("H_Des_Aux1") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "08"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "09"
                rstdestino2("H_Cta_Aux1") = VAR_ORG
                rstdestino2("H_Des_Aux1") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "11"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "12"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
            Case "00"
                rstdestino2("H_Cta_Aux1") = ""
                rstdestino2("H_Des_Aux1") = ""
        End Select

        Select Case d_aux2_1
            Case "01"
                rstdestino2("H_Cta_Aux2") = VAR_BENEF
                rstdestino2("H_Des_Aux2") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux2") = VAR_CTA
                rstdestino2("H_Des_Aux2") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux2") = VAR_PROY2
                rstdestino2("H_Des_Aux2") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux2") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux2") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "06"
                rstdestino2("H_Cta_Aux2") = VAR_DPTO
                rstdestino2("H_Des_Aux2") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "08"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "09"
                rstdestino2("H_Cta_Aux2") = VAR_ORG
                rstdestino2("H_Des_Aux2") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "11"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "12"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
            Case "00"
                rstdestino2("H_Cta_Aux2") = ""
                rstdestino2("H_Des_Aux2") = ""
        End Select

        Select Case d_aux3_1
            Case "01"
                rstdestino2("H_Cta_Aux3") = VAR_BENEF
                rstdestino2("H_Des_Aux3") = VAR_BEND
            Case "02"
                rstdestino2("H_Cta_Aux3") = VAR_CTA
                rstdestino2("H_Des_Aux3") = VAR_CTAD
            Case "03"
                rstdestino2("H_Cta_Aux3") = VAR_PROY2
                rstdestino2("H_Des_Aux3") = VAR_EDIFD
            Case "04"
                rstdestino2("H_Cta_Aux3") = VAR_COD4        'Ado_datos.Recordset("unidad_codigo")
                rstdestino2("H_Des_Aux3") = VAR_UNID
            Case "05"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "06"
                rstdestino2("H_Cta_Aux3") = VAR_DPTO
                rstdestino2("H_Des_Aux3") = VAR_DPTOD
            Case "07"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "08"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "09"
                rstdestino2("H_Cta_Aux3") = VAR_ORG
                rstdestino2("H_Des_Aux3") = VAR_ORGD
            Case "10"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "11"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "12"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
            Case "00"
                rstdestino2("H_Cta_Aux3") = ""
                rstdestino2("H_Des_Aux3") = ""
        End Select
'        If d_aux1_1 = "01" Then
'          rstdestino2("H_Cta_Aux1") = IIf(Len(Trim(VAR_BENEF)) > 0, VAR_BENEF, "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino2("H_Cta_Aux1") = VAR_CTA
'        End If
        rstdestino2("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        If i = 1 Then
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.87
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.87
        Else
            rstdestino2("H_MontoBs") = (IIf(VAR_BS2 < 0, (VAR_BS2 * -1), VAR_BS2)) * 0.13
            rstdestino2("H_MontoDl") = (IIf(VAR_DOL2 < 0, (VAR_DOL2 * -1), VAR_DOL2)) * 0.13
        End If
        rstdestino2("H_Cambio") = GlTipoCambioMercado
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

      If yacontabilizo = 0 Then
        rstdestino2("Usr_codigo") = glusuario
        rstdestino2("Fecha_registro") = Date
        rstdestino2("Hora_registro") = Format(Time, "hh:mm:ss")
      End If

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
        VAR_CTAD = rstdestino!cta_descripcion
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
  'marca1 = Ado_datos.Recordset.Bookmark
  'rs_datos.Update
  'rs_datos.Requery
  Call OptFilGral1_Click
  'Set Ado_datos.Recordset = rs_datos
  'If rs_datos.RecordCount > 0 Then
    Ado_datos.Recordset.Move marca1 - 1
  'End If
  'db.Execute "EXEC ts_mf_ActualizaCtaBancaria"

End Sub

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
    Select Case Ado_datos.Recordset!unidad_codigo
        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
            VAR_PROY = 12
        Case "UCOM"
            VAR_PROY = 17
        Case "DVTA"
            VAR_PROY = 18

    End Select

    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & dtc_desc3.Text & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
    Else
        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
           "VALUES (" & VAR_PROY & ", '" & Ado_datos.Recordset!edif_codigo & "', '" & dtc_desc3.Text & "', '" & Ado_datos.Recordset!unidad_codigo & "', " & glGestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
    End If
    '
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
   VAR_COD4 = Ado_datos.Recordset!unidad_codigo
   VAR_CODTIPO = "DEI"
   Select Case VAR_COD4
        Case "DVTA"              'INI COMERCIAL
            VAR_ORG = "111"
            VAR_PARTIDA = "11310"
        Case "COMEX"            'INI COMEX
            VAR_ORG = "111"
            VAR_PARTIDA = "11310"
        Case "DNINS"            'INI INSTALACIONES
            VAR_ORG = "111"
            VAR_PARTIDA = "11350"
        Case "DNAJS"            'INI AJUSTE
            VAR_ORG = "113"
            VAR_PARTIDA = "11350"
        Case "DNMAN"            'INI MANTENIMIENTO
            VAR_ORG = "112"
            VAR_PARTIDA = "11320"
        Case "DNREP"            'INI REPARACIONES
            VAR_ORG = "113"
            VAR_PARTIDA = "11330"
        Case "DNMOD"            'INI MODERNIZACION
            VAR_ORG = "114"
            VAR_PARTIDA = "11340"
        Case "DNEME"            'INI EMERGENCIAS
            VAR_ORG = "113"
            VAR_PARTIDA = "11330"
        Case Else               'INI CREDITO
            VAR_ORG = "311"
            VAR_PARTIDA = "11350"
   End Select
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False

      'If v_añadir = 1 Then
        'EFECTIVO o a CREDITO
         'db.BeginTrans
         Call add_correl
         Set rstdestino = New ADODB.Recordset
         rstdestino.Open "select * from fo_ingresos_cabecera order by org_codigo, ingreso_codigo   ", db, adOpenDynamic, adLockOptimistic
         rstdestino.AddNew
         rstdestino("Ges_Gestion") = glGestion      'Year(Date)     'Ado_datos.Recordset("ges_gestion")
         rstdestino("ingreso_codigo") = correlativo1
         VAR_CODANT = correlativo1
         'CAMBIAR org_codigo
         rstdestino("org_codigo") = VAR_ORG
         'CAMBIAR org_codigo
         'CAMBIAR COD ingreso_codigo_anterior
         rstdestino("ingreso_codigo_anterior") = correlativo1
         'CAMBIAR COD ingreso_codigo_anterior
         'CAMBIAR DEI O REC
         'VAR_CODTIPO = "DEI"
         rstdestino("Codigo_tipo") = VAR_CODTIPO    '"DEI"
         'VAR_CODTIPO = "DEI"
         'CAMBIAR DEI O REC
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-01"
         rstdestino("etapa_codigo") = "FIN-01-01"
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = "R-110"
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_COD4     'Ado_datos.Recordset("unidad_codigo")
         rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos.Recordset("solicitud_codigo")
         rstdestino("solicitud_tipo") = VAR_TIPO    '"10"

         rstdestino("beneficiario_codigo") = VAR_BENEF      'Ado_datos.Recordset("beneficiario_codigo")
         'VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
         rstdestino("tipo_moneda") = "BOB"
         VAR_MONEDA = "BOB"
         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA2  'Ado_datos.Recordset("venta_descripcion")
         VAR_GLOSA = "INGRESO POR: " + VAR_GLOSA2       'Ado_datos.Recordset("venta_descripcion")
         If Ado_datos.Recordset("venta_tipo") = "E" Then
            rstdestino("tipo_comp") = "DYR"
         Else
            rstdestino("tipo_comp") = "DEI"
         End If
         'CAMBIAR FTE
         Select Case VAR_ORG
             Case "111"              'INI SERVICIOS DE PROVISION E INSTALACION
                 VAR_FTE = "10"
             Case "112"            'INI SERVICIO DE MANTENIMIENTO - MANTENIMIENTO PREVENTIVO
                 VAR_FTE = "10"
             Case "113"            'INI SERVICIO DE REPARACIONES - MANTENIMIENTO CORRECTIVO
                 VAR_FTE = "10"
             Case "114"            'INI SERVICIO DE MODERNIZACION
                 VAR_FTE = "10"
             Case "211"            'INI APORTES DE CAPITAL
                 VAR_FTE = "20"
             Case "311"            'INI BANCO MERCANTIL SANTA CRUZ
                 VAR_FTE = "30"
             Case "312"            'INI BANCO DE CREDITO
                 VAR_FTE = "30"
             Case "411"            'INI AMT - REPOSICION DE PIEZAS Y PARTES
                 VAR_FTE = "40"
             Case Else               'INI OTROS
                 VAR_FTE = "10"
        End Select
         rstdestino("fte_codigo") = VAR_FTE
         'CAMBIAR FTE
         'CAMBIAR RUBROS    'wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww ya pues
         'rstdestino("rubro_codigo") = "11200"
         'VAR_PARTIDA = "11200"
         'VAR_PARTIDA = "11320"
         rstdestino("rubro_codigo") = VAR_PARTIDA
         'CAMBIAR RUBROS
         rstdestino("cheque_o_trf") = ""
         rstdestino("Bco_codigo") = "NN"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = "NN"
         VAR_CTA = "NN"
         'CAMBIAR CTA
         rstdestino("numero_documento") = "0"
         rstdestino("unidad_codigo_ant") = VAR_CITE     'Ado_datos.Recordset("unidad_codigo_ant")
         'VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
         rstdestino("monto_dolares") = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         rstdestino("monto_bolivianos") = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         rstdestino("monto_recaudado_dolares") = 0
         rstdestino("monto_recaudado_bolivianos") = 0
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = Ado_datos.Recordset("edif_codigo")
         VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
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
'   Else
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
  'FALTAAAAA!! org_codigo JQA 2014-07-10
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '" & VAR_ORG & "' ", db, adOpenDynamic, adLockOptimistic
  'rstcorrel_ing.Open "select * from fc_organismo_financiamiento where org_codigo = '111' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "'", db, adOpenDynamic, adLockOptimistic
  If rstcorrel_ing.RecordCount = 0 Then
     rstcorrel_ing.AddNew
     rstcorrel_ing("org_codigo") = VAR_ORG
     rstcorrel_ing("ges_gestion") = glGestion       'Ado_datos.Recordset("ges_gestion")  'Trim(lblges_gestion.Caption)
     'rstcorrel_ing("correlativo") = 1
     rstcorrel_ing("correlativo_ingreso") = 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo_ingreso")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo_ingreso")
  Else
     VARG_ORGD = rstcorrel_ing!org_descripcion
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
                rstpagos("ges_gestion") = glGestion     'Ado_datos.Recordset("ges_gestion")
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

Private Sub CmdGrabaDet_Click()
On Error GoTo UpdateErr
  If dtc_codigo15 = "" Then
     MsgBox "Debe Elejir un Bien para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
      'ado_datos14.Recordset.CancelBatch
      'Call AbrirDetalle
    Exit Sub
  End If
  If TxtDescuento.Text = "" Or TxtDescuento.Text = "0" Then
    MsgBox "Debe Registrar la Cantidad Entregada, !! Vuelva a Intentar ...", vbExclamation, "Atención"
     'ado_datos14.Recordset.CancelBatch
    'Call AbrirDetalle
    Exit Sub
  End If
  
'  If dtc_desc13 = "" Then
'    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'    Exit Sub
'  End If
  'If Ado_datos.Recordset!unidad_codigo <> "DNREP" And Ado_datos.Recordset!unidad_codigo <> "UALMR" Then
    If CDbl(TxtDescuento.Text) > CDbl(IIf(Dtc_Stock13.Text = "", "0", Dtc_Stock13.Text)) Then
    '        'VAR_PARTIDA = "OK"
       MsgBox "Saldo Insuficiente en Stock (no se guardara este registro)!..."
       ' ado_datos14.Recordset.CancelBatch
       ' Call AbrirDetalle
       Exit Sub
    End If
  'End If
    
    
'    If Dtc_Stock13.Text = "10" Then
'    MsgBox ("El saldo en el almacen es de 10")
'    End If
'
'    If Dtc_Stock13.Text = "2" Then
'    MsgBox ("ATENCION El saldo en el almacen es de 2 ")
''      ado_datos14.Recordset.CancelBatch
''      Call AbrirDetalle
''       Exit Sub
'    End If
    
    'VARIABLES DE LA CABECERA
    VAR_ALMX = Ado_datos.Recordset!almacen_codigo
    correlv = Ado_datos.Recordset!venta_codigo
    VAR_PROY2 = Ado_datos.Recordset!edif_codigo
    VAR_BEN3 = Ado_datos.Recordset!beneficiario_codigo_alm
    VAR_DOC = Ado_datos.Recordset!doc_codigo_alm
    VAR_ALMD = IIf(IsNull(Ado_datos.Recordset!almacen_codigo_d), "0", Ado_datos.Recordset!almacen_codigo_d)
    'If CDbl(Dtc_Stock13.Text) >= CDbl(TxtDescuento.Text) Then
        If swnuevo = 1 Then
            Set rs_aux8 = New ADODB.Recordset
            If rs_aux8.State = 1 Then rs_aux8.Close
            rs_aux8.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & "  and bien_codigo = '" & dtc_codigo15.Text & "'", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux8.RecordCount > 0 Then
                MsgBox "Error, El bien ya fue registrado vuelva a intentar...", , "Atención"
                'ado_datos14.Recordset.CancelBatch
                'Call AbrirDetalle
                Exit Sub
            Else
                'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
                ado_datos14.Recordset!venta_codigo = Ado_datos.Recordset!venta_codigo
                ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset!ges_gestion
                ado_datos14.Recordset!estado_codigo = "APR"
                ado_datos14.Recordset!usr_codigo = glusuario
                ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
                ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
            End If
        End If
        VAR_BIEN2 = Trim(dtc_codigo15.Text)                                     'Codigo Bien (Equipo, Producto, etc)
        ado_datos14.Recordset!bien_codigo = Trim(VAR_BIEN2)                     'Codigo Bien (Equipo, Producto, etc)
            ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
            ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
            ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
    '       ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
            ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW FALTA PARAMETRIZAR WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            
            ado_datos14.Recordset!almacen_codigo = IIf(Ado_datos.Recordset!almacen_codigo = "", "2", Ado_datos.Recordset!almacen_codigo)
            ado_datos14.Recordset!doc_codigo_alm = IIf(Ado_datos.Recordset!doc_codigo_alm = "", "R-115", Ado_datos.Recordset!doc_codigo_alm)
            ado_datos14.Recordset!doc_numero_alm = IIf(Ado_datos.Recordset!doc_numero_alm = "", "1", Ado_datos.Recordset!doc_numero_alm)
            VAR_COD2 = IIf(Ado_datos.Recordset!doc_numero_alm = "", "1", Ado_datos.Recordset!doc_numero_alm)
            If TxtCantidad.Text = "" Then
                TxtCantidad.Text = TxtDescuento.Text
            End If
            Call CARGAPARAM
            ado_datos14.Recordset!almacen_tipo = VAR_ALMT                              'Tipo de Almacen I, R, H, A
            ado_datos14.Recordset!venta_det_cantidad = CDbl(IIf(TxtCantidad.Text = "", 1, TxtCantidad))  'Cantidad Solicitada
            ado_datos14.Recordset!bien_cantidad_por_empaque = CDbl(IIf(TxtDescuento = "", 1, TxtDescuento)) 'Cantidad Entregada
            ado_datos14.Recordset!estado_almacen = "REG"
            ado_datos14.Recordset.Update
            'ACTUALIZA MONTOS DEL BIEN
            If swnuevo = 1 Then
'                db.Execute "UPDATE ao_ventas_detalle SET venta_descuento_bs = '0', venta_descuento_dol = '0' WHERE venta_codigo = " & correlv & " AND bien_codigo = '" & VAR_BIEN2 & "' "

                db.Execute "UPDATE ao_ventas_detalle SET ao_ventas_detalle.venta_precio_unitario_bs  = ac_bienes.bien_precio_venta_final, ao_ventas_detalle.venta_precio_total_bs = ac_bienes.bien_precio_venta_final FROM ao_ventas_detalle INNER JOIN ac_bienes " & _
                    " ON ao_ventas_detalle.bien_codigo  = ac_bienes.bien_codigo WHERE ao_ventas_detalle.venta_codigo = " & correlv & " AND ao_ventas_detalle.bien_codigo = '" & VAR_BIEN2 & "' "
            End If
            'ACTUALIZA ao_almacen_salidas   'Copia el registro completo
                       
            Set rs_datos15 = New ADODB.Recordset
            If rs_datos15.State = 1 Then rs_datos15.Close
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'I' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
            Set ado_datos15.Recordset = rs_datos15
            ado_datos15.Refresh
            Call AbrirDetalle
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = False
            FraNavega.Enabled = True
            FrmDetalle.Enabled = True
            FrmABMDet.Visible = True
            FrmEdita.Enabled = False
            'Call OptFilGral1_Click
            swnuevo = 0
'        End If
    'Else
    '    MsgBox "Saldo Insuficiente en Stock registrado en Almacenes, verifique y luego intente nuevamente !..."
    'End If
  'Else
  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
  'End If
  accion = ""
  Exit Sub
UpdateErr:
    MsgBox Err.Description
End Sub


Private Sub BtnImprimir2_Click()
    If ado_datos14.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
        CryR01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado.rpt" '
        CryR01.WindowShowPrintSetupBtn = True
        CryR01.WindowShowRefreshBtn = True
        'CryR01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryR01.StoredProcParam(0) = ado_datos14.Recordset!bien_codigo
        CryR01.StoredProcParam(1) = Trim(Str(ado_datos14.Recordset!almacen_codigo))            'dtc_codigo1.Text
        CryR01.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryR01.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        CryR01.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
        'CryR01.Formulas(2) = "DEL_AL = '' "
        'CryR01.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
        
        iResult = CryR01.PrintReport
        If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
        Fra_reporte.Visible = False
        
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    Fra_reporte.Visible = True
End Sub

Private Sub BtnAnlDetalle_Click()
 If ado_datos14.Recordset.RecordCount > 0 Then
    If ado_datos14.Recordset("estado_almacen") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + ado_datos14.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then

            db.Execute "delete ao_ventas_detalle Where venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "' and ges_gestion = " & Ado_datos.Recordset!ges_gestion & " and bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' "

            Call AbrirDetalle

          End If
       Else
          MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       
       End If
   Else
     MsgBox "No se puede BORRAR, el registro ya fue BORRADO o APROBADO (APR), Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle_Click()
 If ado_datos14.Recordset.RecordCount > 0 Then
  If ado_datos14.Recordset!estado_almacen = "REG" Then
    FraNavega.Enabled = False
    FrmDetalle.Enabled = False
    swnuevo = 2

    marca1 = Ado_datos.Recordset.Bookmark
    TxtNroVenta.Text = Ado_datos.Recordset!venta_codigo  'txt_venta.Text
    TxtNroVenta.Locked = True
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False

    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    FrmABMDet.Visible = False
    If parametro <> Ado_datos.Recordset!unidad_codigo Then
    dtc_desc15.Locked = False
    dtc_desc15.backColor = &HC0C0C0
    Text9.Visible = True
    Else
    dtc_desc15.Locked = True
    dtc_desc15.backColor = &HFFFFFF
    Text9.Visible = False
    End If

     If ado_datos14.Recordset!par_codigo = "43340" Then
        dtc_codigo13.Text = "0"
        dtc_desc13.BoundText = dtc_codigo13.BoundText
        dtc_desc13.backColor = &H80000013
        dtc_desc13.ForeColor = &HFFFFFF
     Else
        dtc_desc13.backColor = &HFFFFFF
        dtc_desc13.ForeColor = &H80000008
        If ado_datos14.Recordset!bien_cantidad_por_empaque = "0" Then
            TxtDescuento.Text = ado_datos14.Recordset!venta_det_cantidad
     End If
    End If
    dtc_desc15.BoundText = dtc_codigo15.BoundText
    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
  Else
    MsgBox "Los registros Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
  End If
  
    Else
     MsgBox "No se puede Modificar, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub Command1_Click()
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
     If rs_datos!estado_almacen = "APR" Then
       sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
'     If ExisteReg(Ado_datos.Recordset!unidad_codigo_sol, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
         
'          rs_datos!fecha_registro = Date
'          rs_datos!usr_codigo = glusuario
'           Ado_datos.Recordset.Requery
'           Ado_datos.Refresh
          db.Execute "ap_ventas_grla 3,'" & glGestion & "', " & Ado_datos.Recordset!almacen_codigo & ", '" & Ado_datos.Recordset!doc_codigo_alm & "', " & Ado_datos.Recordset!doc_numero_alm & ", '" & ado_datos14.Recordset!bien_codigo & "', '" & Ado_datos.Recordset!edif_codigo & "'," & Ado_datos.Recordset!venta_codigo & ",'" & Ado_datos.Recordset!beneficiario_codigo_alm & "','" & Ado_datos.Recordset!fecha_verif & "'," & ado_datos14.Recordset!bien_cantidad_por_empaque & "," & precio_tot & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & glusuario & "','" & Ado_datos.Recordset!venta_descripcion & "'," & precio_uni & ""
            rs_datos!estado_almacen = "REG"
           Call AbrirDetalle
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede DESPROBAR un registro Anulado(ANL) o Registrado (REG) ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede DESAPROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub dtc_Aux11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_Aux11.BoundText
    dtc_desc11.BoundText = dtc_Aux11.BoundText
End Sub

Private Sub dtc_Aux20_Click(Area As Integer)
    dtc_desc20.BoundText = dtc_Aux20.BoundText
    dtc_codigo20.BoundText = dtc_Aux20.BoundText
End Sub

'Private Sub CmdDetCabeza_Click()
'    fraOpciones.Visible = False
'    FrmDetalle.Visible = True
'    FrmCobranza.Visible = True
'    FraNavega.Enabled = False
'    If Not (adoDetalleSolicitud.Recordset.BOF) Then adoDetalleSolicitud.Recordset.MoveFirst
'End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo20_Click(Area As Integer)
    dtc_desc20.BoundText = dtc_codigo20.BoundText
    dtc_Aux20.BoundText = dtc_codigo20.BoundText
End Sub

Private Sub dtc_codigo21_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo22_Click(Area As Integer)
    dtc_desc22.BoundText = dtc_codigo22.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
'    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc15_Change()
If accion <> "NEW" Then
If Ado_datos.Recordset!almacen_codigo <> "NULL" Then
Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux9.Close
    rs_aux9.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo = " & Ado_datos.Recordset!almacen_codigo & " and bien_codigo ='" & dtc_codigo15.Text & "'", db, adOpenStatic
   ' Set AdoAux9.Recordset = rs_aux9
   If rs_aux9.RecordCount > 0 Then
    Dtc_Stock13.Text = IIf(IsNull(rs_aux9!stock_actual), 0, rs_aux9!stock_actual)
    End If
  End If
  Else
  Dtc_Stock13.Text = "0"
End If
End Sub

Private Sub dtc_desc15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
TxtDescuento.SetFocus
End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc20_Click(Area As Integer)
    dtc_codigo20.BoundText = dtc_desc20.BoundText
    dtc_Aux20.BoundText = dtc_desc20.BoundText
    Call pDeptoD(dtc_Aux20.Text)
    dtc_desc22.Enabled = True
    TxtConcepto.Text = dtc_desc3.Text + " " + VAR_BIEN + " A " + dtc_desc20.Text
End Sub

Private Sub pDeptoD(CodigoA As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_departamento where depto_codigo  = '" & CodigoA & "'"
   
   Set dtc_codigo22.RowSource = Nothing
   Set dtc_codigo22.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo22.ReFill
   dtc_codigo22.BoundText = Empty
   
   Set dtc_desc22.RowSource = Nothing
   Set dtc_desc22.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc22.ReFill
   dtc_desc22.BoundText = Empty
End Sub

Private Sub dtc_desc21_Click(Area As Integer)
  dtc_codigo21.BoundText = dtc_desc21.BoundText
End Sub

Private Sub dtc_desc22_Click(Area As Integer)
    dtc_codigo22.BoundText = dtc_desc22.BoundText
End Sub

'Private Sub dtc_desc2_Click(Area As Integer)
'    dtc_codigo2.BoundText = dtc_desc2.BoundText
'    Dtc_aux2.BoundText = dtc_desc2.BoundText
'    Dtc_deudor2.BoundText = dtc_desc2.BoundText
'End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc3_LostFocus()
    If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
        dtc_desc20.Visible = True
        lbl_Adestino.Visible = True
        dtc_desc22.Visible = True
        lbl_Rdestino.Visible = True
        TxtConcepto.Locked = False
        TxtConcepto.Text = "TRASPASO DESDE ALMACEN DE " + VAR_BIEN + " A " + dtc_desc20.Text
    Else
        dtc_desc20.Visible = False
        lbl_Adestino.Visible = False
        dtc_desc22.Visible = False
        lbl_Rdestino.Visible = False
        TxtConcepto.Locked = False
        TxtConcepto.Text = "SALIDA DE ALMACEN DE " + VAR_BIEN
    End If
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
'    dtc_aux4.BoundText = dtc_desc4.BoundText
    VAR_BEN2 = dtc_codigo4.Text
    Call pAlmacen(dtc_codigo4.BoundText)
    dtc_desc11.Enabled = True
End Sub

Private Sub pAlmacen(CodigoA As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from ac_almacenes where beneficiario_codigo = '" & CodigoA & "'"
   
   Set dtc_codigo11.RowSource = Nothing
   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo11.ReFill
   dtc_codigo11.BoundText = Empty
   
   Set dtc_desc11.RowSource = Nothing
   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc11.ReFill
   dtc_desc11.BoundText = Empty

   Set dtc_Aux11.RowSource = Nothing
   Set dtc_Aux11.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_Aux11.ReFill
   dtc_Aux11.BoundText = Empty

End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.Text = VAR_BEN2
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
    If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
        Call pAlmacenD(dtc_codigo5.BoundText)
        dtc_desc20.Enabled = True
    End If
End Sub

Private Sub pAlmacenD(CodigoA As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from ac_almacenes where beneficiario_codigo = '" & CodigoA & "'"
   
   Set dtc_codigo20.RowSource = Nothing
   Set dtc_codigo20.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo20.ReFill
   dtc_codigo20.BoundText = Empty
   
   Set dtc_desc20.RowSource = Nothing
   Set dtc_desc20.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc20.ReFill
   dtc_desc20.BoundText = Empty

End Sub

Private Sub dtc_codigo13_Click(Area As Integer)
    dtc_desc13.BoundText = dtc_codigo13.BoundText
    Dtc_Stock13.BoundText = dtc_codigo13.BoundText
End Sub

Private Sub dtc_desc13_Click(Area As Integer)
    dtc_codigo13.BoundText = dtc_desc13.BoundText
    Dtc_Stock13.BoundText = dtc_desc13.BoundText
End Sub

Private Sub dtc_codigo2A_Click(Area As Integer)
    dtc_desc2A.BoundText = dtc_codigo2A.BoundText
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
    dtc_Aux11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
    dtc_Aux11.BoundText = dtc_desc11.BoundText
    Call pDepto(dtc_Aux11.Text)
    dtc_desc21.Enabled = True
End Sub

Private Sub pDepto(CodigoA As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_departamento where depto_codigo  = '" & CodigoA & "'"
   
   Set dtc_codigo21.RowSource = Nothing
   Set dtc_codigo21.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_codigo21.ReFill
   dtc_codigo21.BoundText = Empty
   
   Set dtc_desc21.RowSource = Nothing
   Set dtc_desc21.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_desc21.ReFill
   'dtc_desc21.BoundText = Empty
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

Private Sub dtc_codigo15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_codigo15.BoundText
    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
    Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux9.Close
    rs_aux9.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo = " & Ado_datos.Recordset!almacen_codigo & " and bien_codigo ='" & dtc_codigo15.Text & "'", db, adOpenStatic
   ' Set AdoAux9.Recordset = rs_aux9
   If rs_aux9.RecordCount > 0 Then
    Dtc_Stock13.Text = IIf(IsNull(rs_aux9!stock_actual), 0, rs_aux9!stock_actual)
  Else
  Dtc_Stock13.Text = "0"
End If

'    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
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

Private Sub dtc_subgrupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_subgrupo15.BoundText
    dtc_desc15.BoundText = dtc_subgrupo15.BoundText
    dtc_unimed15.BoundText = dtc_subgrupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_subgrupo15.BoundText
    dtc_grupo15.BoundText = dtc_subgrupo15.BoundText
    Dtc_partida15.BoundText = dtc_subgrupo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_subgrupo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_subgrupo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_subgrupo15.BoundText
End Sub

Private Sub dtc_partida15_Click(Area As Integer)
    dtc_desc15.BoundText = Dtc_partida15.BoundText
    dtc_unimed15.BoundText = Dtc_partida15.BoundText
    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
    dtc_grupo15.BoundText = Dtc_partida15.BoundText
    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
    dtc_codigo15.BoundText = Dtc_partida15.BoundText
'    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
'    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
'    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
End Sub

Private Sub dtc_desc15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_desc15.BoundText
    dtc_unimed15.BoundText = dtc_desc15.BoundText
    dtc_stocktotal15.BoundText = dtc_desc15.BoundText
    dtc_grupo15.BoundText = dtc_desc15.BoundText
    dtc_subgrupo15.BoundText = dtc_desc15.BoundText
    Dtc_partida15.BoundText = dtc_desc15.BoundText
    
    Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux9.Close
    rs_aux9.Open "SELECT * FROM ao_almacen_totales WHERE almacen_codigo = " & IIf(IsNull(Ado_datos.Recordset!almacen_codigo), 0, Ado_datos.Recordset!almacen_codigo) & " and bien_codigo ='" & dtc_codigo15.Text & "'", db, adOpenStatic
   ' Set AdoAux9.Recordset = rs_aux9
    If rs_aux9.RecordCount > 0 Then
        Dtc_Stock13.Text = IIf(IsNull(rs_aux9!stock_actual), 0, rs_aux9!stock_actual)
    Else
        Dtc_Stock13.Text = "0"
    End If
'    dtc_precioventafinal15.BoundText = dtc_desc15.BoundText
'    dtc_precioventabase15.BoundText = dtc_desc15.BoundText
'    dtc_preciocompra15.BoundText = dtc_desc15.BoundText
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

'Private Sub dtc_desc2_LostFocus()
'    'If AdoBeneficiario.Recordset!beneficiario_deudor = "SI" Then
'    If Dtc_deudor2.Text = "SI" Then
'        Dtc_deudor2.backColor = &HFF&
'    Else
'        Dtc_deudor2.backColor = &H80000010
'    End If
'
'End Sub

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
    
'    TxtDescuento.Text = "0"
'    TxtPrecioU.Text = dtc_precioventabase15.Text
'    Call AbreAlmacen
End Sub

Private Sub dtc_grupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_grupo15.BoundText
    dtc_desc15.BoundText = dtc_grupo15.BoundText
    dtc_unimed15.BoundText = dtc_grupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_grupo15.BoundText
    dtc_subgrupo15.BoundText = dtc_grupo15.BoundText
    Dtc_partida15.BoundText = dtc_grupo15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_grupo15.BoundText
'    dtc_precioventabase15.BoundText = dtc_grupo15.BoundText
'    dtc_preciocompra15.BoundText = dtc_grupo15.BoundText
End Sub

Private Sub dtc_stocktotal15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_stocktotal15.BoundText
    dtc_desc15.BoundText = dtc_stocktotal15.BoundText
    dtc_unimed15.BoundText = dtc_stocktotal15.BoundText
    dtc_grupo15.BoundText = dtc_stocktotal15.BoundText
    dtc_subgrupo15.BoundText = dtc_stocktotal15.BoundText
    Dtc_partida15.BoundText = dtc_stocktotal15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_stocktotal15.BoundText
'    dtc_precioventabase15.BoundText = dtc_stocktotal15.BoundText
'    dtc_preciocompra15.BoundText = dtc_stocktotal15.BoundText
End Sub

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

Private Sub dtc_stock13_Click(Area As Integer)
'    dtc_codigo13.BoundText = Dtc_Stock13.BoundText
'    dtc_desc13.BoundText = Dtc_Stock13.BoundText
End Sub

Private Sub dtc_unimed15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_unimed15.BoundText
    dtc_desc15.BoundText = dtc_unimed15.BoundText
    dtc_stocktotal15.BoundText = dtc_unimed15.BoundText
    dtc_grupo15.BoundText = dtc_unimed15.BoundText
    dtc_subgrupo15.BoundText = dtc_unimed15.BoundText
    Dtc_partida15.BoundText = dtc_unimed15.BoundText
'    dtc_precioventafinal15.BoundText = dtc_unimed15.BoundText
'    dtc_precioventabase15.BoundText = dtc_unimed15.BoundText
'    dtc_preciocompra15.BoundText = dtc_unimed15.BoundText
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
'    End If
'    'Ado_datos4.Refresh

'    DTPfechaIni.Value = DTPfechasol.Value
''    'validar fecha solicitud OJO JQA 31/12/2014
''    Set rs_TipoCambio = New ADODB.Recordset
''    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
''    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
''    If rs_TipoCambio.RecordCount > 0 Then
''        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
''    End If
'End Sub
Private Sub CARGAPARAM()
'    Set rs_datos12 = New ADODB.Recordset
'    If rs_datos12.State = 1 Then rs_datos12.Close
'    rs_datos12.Open "select * from gc_usuarios where usr_codigo = '" & glusuario & "'  ", db, adOpenStatic
'    If rs_datos12.RecordCount > 0 Then
'       VAR_BENEF = rs_datos12!beneficiario_codigo
'    Else
'       VAR_BENEF = "0"
'    End If
    
'    Select Case parametro
    Select Case VAR_ORIGEN
      Case "UALMI"          ', "ALMIS", "ALMIB", "ALMIC"
          VAR_BIEN = "INSUMOS"
          VAR_TIPO = "25"
          VAR_N1 = "TEC"
          VAR_N2 = "TEC-06"
          VAR_N3 = "TEC-06-02"
          VAR_POA = "3.2.8"
          VAR_ALMT = "I"
          
      Case "UALMR"          ', "ALMRS", "ALMRB", "ALMRC"
          VAR_BIEN = "REPUESTOS"
          VAR_TIPO = "26"
          VAR_N1 = "TEC"
          VAR_N2 = "TEC-07"
          VAR_N3 = "TEC-07-02"
          VAR_POA = "3.2.5"
          VAR_ALMT = "R"

      Case "UALMH"          ', "ALMHS", "ALMHB", "ALMHC"
          VAR_BIEN = "HERRAMIENTAS"
          VAR_TIPO = "27"
          VAR_N1 = "TEC"
          VAR_N2 = "TEC-08"
          VAR_N3 = "TEC-08-02"
          VAR_POA = "3.2.9"
          VAR_ALMT = "H"
      Case "GADM"
          VAR_BIEN = "ADMINISTRACION"
          VAR_TIPO = "31"
          VAR_N1 = "ADM"
          VAR_N2 = "ADM-04"
          VAR_N3 = "ADM-04-02"
          VAR_POA = "7.2.1"
          VAR_ALMT = "A"
      Case Else
          VAR_BIEN = "INSUMOS"
          VAR_TIPO = "25"
          VAR_N1 = "TEC"
          VAR_N2 = "TEC-06"
          VAR_N3 = "TEC-06-02"
          VAR_POA = "3.2.8"
          VAR_ALMT = "I"
    End Select
End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    accion = ""
    VAR_SW = ""
    lbl_cerrado = ""
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
        VAR_DA = "1.4"
    End If
    VAR_ORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            VAR_DPTO = "3"
            Select Case Aux
               Case "UALMI"    'INSUMOS
                   Aux = "ALMIB"
               Case "UALMR"    'REPUESTOS
                   Aux = "ALMRB"
               Case "UALMH"    'HERRAMIENTAS
                   Aux = "ALMHB"
               Case "GADM"    ' GENERAL
                   Aux = "DADMB"
            End Select
        Case "1.7"    'Santa Cruz
            VAR_DPTO = "7"
            Select Case Aux
               Case "UALMI"    'INSUMOS
                   Aux = "ALMIS"
               Case "UALMR"    'REPUESTOS
                   Aux = "ALMRS"
               Case "UALMH"    'HERRAMIENTAS
                   Aux = "ALMHS"
               Case "GADM"    ' GENERAL
                   Aux = "DADMS"
            End Select
        Case "1.3", "1.4"    'La Paz (Técnico=1.3)
            VAR_DPTO = "2"
            Select Case Aux
               Case "UALMI"    'INSUMOS
                   Aux = "UALMI"
               Case "UALMR"    'REPUESTOS
                   Aux = "UALMR"
               Case "UALMH"    'HERRAMIENTAS
                   Aux = "UALMH"
               Case "GADM"    ' GENERAL
                   Aux = "GADM"
            End Select
        Case "1.9"    ' Chuquisaca
            VAR_DPTO = "1"
            Select Case Aux
               Case "UALMI"    'INSUMOS
                   Aux = "ALMIC"
               Case "UALMR"    'REPUESTOS
                   Aux = "ALMRC"
               Case "UALMH"    'HERRAMIENTAS
                   Aux = "ALMHC"
               Case "GADM"    ' GENERAL
                   Aux = "DADMC"
            End Select
        Case Else    ' TODO
            VAR_DPTO = "2"
     End Select
    parametro = Aux
    VAR_R = "R-115"
    Call CARGAPARAM
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'Usuario
    lbl_cerrado.Caption = ""
    
    FrmDetalle.Caption = "DETALLE DE " + VAR_BIEN
    aw_almacen_salida.Caption = "" + VAR_BIEN
    
    mbDataChanged = False
    FrmCabecera.Enabled = False
    dg_datos.Enabled = True
    GlNombFor = "F04"

    marca1 = 1
    deta2 = 0
    swgrabar = 0
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
'    Chk_plazo.Value = 0
  
	Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'UNIDAD EJECUTORA
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'Beneficiario Personas Nat. y Juridicas
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from gc_unidad_ejecutora where estado_codigo = 'APR' AND da_codigo = '" & VAR_DA & "'", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText

    'Proyecto de Edificación
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones WHERE depto_codigo= '" & VAR_DPTO & "' AND (estado_codigo = 'APR' OR (estado_codigo = 'ANL' AND edif_tipo = 'NN')) order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    'Beneficiario Funcionario - Almacen
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' order by beneficiario_denominacion", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

    'Beneficiario Funcionario - Entregado a:
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "select * from gc_beneficiario where tipoben_codigo = '1' and estado_codigo = 'APR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    'ac_almacenes ' Origen
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "select * from ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1' ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
'    ''rs_datos11.Open "select * from ac_almacenes where depto_codigo = '" & VAR_DPTO & "' AND almacen_tipo = '" & VAR_ALMT & "' ", db, adOpenStatic
''    If VAR_BENEF = "0" Then
''        rs_datos11.Open "select * from ac_almacenes where almacen_codigo <> '1' and almacen_codigo <> '2'  ", db, adOpenStatic
''    Else
''        rs_datos11.Open "select * from ac_almacenes where beneficiario_codigo = '" & VAR_BENEF & "'  ", db, adOpenStatic
''    End If
'    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText
'    If Ado_datos11.Recordset.RecordCount > 0 Then
'       Ado_datos11.Recordset.MoveFirst
'       VAR_ALMT = rs_datos11!almacen_tipo
'       VAR_DPTO = rs_datos11!depto_codigo
'       VAR_ALMX = rs_datos11!almacen_codigo
'    Else
'       VAR_ALMT = ""
'       VAR_DPTO = ""
'       VAR_ALMX = ""
'    End If

    Set rs_datos13 = New ADODB.Recordset    'Detalle por cada Almacen
    If rs_datos13.State = 1 Then rs_datos13.Close
    'rs_datos13.Open "select * from Av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    rs_datos13.Open "select * from av_almacen_detalle", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos13.Recordset = rs_datos13
    Ado_datos13.Refresh

    'ac_almacenes - Destino
    Set rs_datos20 = New ADODB.Recordset
    If rs_datos20.State = 1 Then rs_datos20.Close
    'rs_datos20.Open "select * from ac_almacenes where beneficiario_codigo <> '" & VAR_BENEF & "'  ", db, adOpenStatic
    rs_datos20.Open "select * from ac_almacenes ", db, adOpenStatic
    Set Ado_datos20.Recordset = rs_datos20
    dtc_desc20.BoundText = dtc_codigo20.BoundText
    
    'gc_departamento - Origen
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "select * from gc_departamento   ", db, adOpenStatic
    'rs_datos21.Open "select * from gc_departamento where depto_codigo = '" & VAR_DPTO & "'  ", db, adOpenStatic      ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
    Set Ado_datos21.Recordset = rs_datos21
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    
    'gc_departamento - Destino
    Set rs_datos22 = New ADODB.Recordset
    If rs_datos22.State = 1 Then rs_datos22.Close
    rs_datos22.Open "select * from gc_departamento  ", db, adOpenStatic
    'rs_datos22.Open "select * from gc_departamento where depto_codigo <>  '" & VAR_DPTO & "'  ", db, adOpenStatic       ''4273257'    'beneficiario_codigo= '" & dtc_codigo4.Text & "'
    Set Ado_datos22.Recordset = rs_datos22
    dtc_desc22.BoundText = dtc_codigo22.BoundText
    
    'Bienes por almacen
    Set rs_datos15 = New ADODB.Recordset
    If rs_datos15.State = 1 Then rs_datos15.Close
    'rs_aux2.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'Select Case parametro
    Select Case VAR_ORIGEN
        Case "UALMI"            ', "ALMIB", "ALMIS", "ALMIC"    'INSUMOS
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'I' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
            Set ado_datos15.Recordset = rs_datos15
            ado_datos15.Refresh
'            VAR_DET = "30000"
        Case "UALMR"            ', "ALMRB", "ALMRS", "ALMRC"    'REPUESTOS
            'rs_datos15.Open "select * from ac_bienes where (par_codigo = '39810' or par_codigo = '39820')   ", db, adOpenKeyset, adLockReadOnly        'and estado_codigo = 'APR'
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'R' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
'            VAR_DET = "39800"
            Set ado_datos15.Recordset = rs_datos15
            ado_datos15.Refresh
        Case "UALMH"            ', "ALMHB", "ALMHS", "ALMHC"    'HERRAMIENTAS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'H' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
'            VAR_DET = "34800"
            Set ado_datos15.Recordset = rs_datos15
              ado_datos15.Refresh
        Case Else
            rs_datos15.Open "select * from ac_bienes where almacen_tipo = 'I' ORDER BY bien_descripcion", db, adOpenKeyset, adLockReadOnly
            Set ado_datos15.Recordset = rs_datos15
            ado_datos15.Refresh
    End Select
   'wwwwwwwwwwwwwwwwwwww
    'db.Execute "DELETE ao_ventas_cabecera where venta_codigo = 0 "
    'Call ABREVENTAS
    
    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub grabar()

  'db.BeginTrans
    If swgrabar = 1 Then
'      Dim rstdestino As New ADODB.Recordset
'      Set rstdestino = New ADODB.Recordset
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select tipo_tramite, numero_correlativo from fc_correl WHERE tipo_tramite='ventas'", db, adOpenDynamic, adLockOptimistic
'      If rstdestino.RecordCount <> 0 Then
'        Ado_datos.Recordset("venta_codigo") = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino!numero_correlativo = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino.Update
'      Else
'        Ado_datos.Recordset("venta_codigo") = 1
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      'Ado_datos.Recordset("venta_codigo") = Ado_datos.Recordset.RecordCount
'      'rstdestino.AddNew
'    End If
        Set rs_aux4 = New ADODB.Recordset
        SQL_FOR = "Select max(solicitud_codigo) as Codigo from ao_solicitud where unidad_codigo = '" & parametro & "' "
        rs_aux4.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If Not rs_aux4.EOF Then
            var_cod = IIf(IsNull(rs_aux4!Codigo), 1, rs_aux4!Codigo + 1)
        Else
            var_cod = 1
        End If
        'CREA VENTA CABECERA
        Set rs_aux6 = New ADODB.Recordset
        If rs_aux6.State = 1 Then rs_aux6.Close
        rs_aux6.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
        If Not rs_aux6.EOF Then
            var_cod5 = IIf(IsNull(rs_aux6!Codigo), 1, rs_aux6!Codigo + 1)
        End If
        VAR_R = "R-115"
'       Ado_datos.Recordset("ges_gestion") = glGestion       'CStr(Year(DTPfechasol.Value))
'       Ado_datos.Recordset("unidad_codigo") = parametro   'dtc_codigo1.Text DTPEjecucion
'       Ado_datos.Recordset("solicitud_codigo") = var_cod
'       Ado_datos.Recordset("venta_codigo") = var_cod5
'       Ado_datos.Recordset("edif_codigo") = dtc_codigo3.Text
''       Ado_datos.Recordset("depto_codigo") = Left(dtc_codigo3.Text, 1)
'       Ado_datos.Recordset("venta_fecha") = Format(IIf(IsNull(DTPFechaSol.Value), Date, DTPFechaSol.Value), "dd,mm,yyyy")
'       Ado_datos.Recordset("venta_fecha_inicio") = Format(Ado_datos.Recordset!venta_fecha, "dd,mm,yyyy")  'DTPFechaIni.Caption
'       Ado_datos.Recordset("venta_fecha_fin") = Format(Ado_datos.Recordset!venta_fecha, "dd,mm,yyyy") 'DTPFechaFin.Caption
'       Ado_datos.Recordset("venta_tipo") = "A"      'dtc_codigo11.Text                'E=Efectivo, C=Credito
'       Ado_datos.Recordset("beneficiario_codigo") = dtc_codigo2.Text        'CLIENTE
'       Ado_datos.Recordset("beneficiario_codigo_resp") = dtc_codigo4.Text   'Responsable Proceso
'
'       Ado_datos.Recordset("beneficiario_codigo_cobr") = "0"      'dtc_codigo5.Text   'Cobrador
'       Ado_datos.Recordset("venta_descripcion") = Trim(UCase(TxtConcepto.Text))
'       'CONT2 = 365 / 30 * Ado_datos.Recordset!venta_cantidad_total
'       'Ado_datos.Recordset("venta_plazo_dias_calendario") = IIf(TxtPlazo.Text = "", CONT2, TxtPlazo.Text)
'       Ado_datos.Recordset("venta_tipo_cambio") = GlTipoCambioMercado        'Val(txtTDC.Text)venta_tipo_cambio
'        'GlTipoCambioOficial As Currency        'GlTipoCambioMercado As Currency        'GlTipoCambioGestion As Currency
''       Ado_datos.Recordset("tipoben_codigo") = IIf(dtc_aux2.Text = "", "1", dtc_aux2.Text)      'Tipo de Beneficiario
'
'       Ado_datos.Recordset("unimed_codigo_cobr") = "MES"        'cmd_unimed2.Text
'       Ado_datos.Recordset("venta_cantidad_cobr") = "1"         'txtCantCobr.Text
'       Ado_datos.Recordset("mes_inicio_crono") = "ENERO"        'RTrim(cmb_mes_ini.Text)
'       VAR_MED2 = Ado_datos.Recordset!unimed_codigo_cobr
'       VAR_COBR2 = Ado_datos.Recordset!venta_cantidad_cobr
'       MControl = Ado_datos.Recordset!mes_inicio_crono
       
       VAR_MED2 = "MES"
       VAR_COBR2 = "1"
       MControl = "ENERO"
        If var_cod < 10 Then
           VAR_CITE = parametro + "-00000" + Trim(var_cod)
        End If
        If var_cod > 9 And var_cod < 100 Then
           VAR_CITE = parametro + "-0000" + Trim(var_cod)
        End If
        If var_cod > 99 And var_cod < 1000 Then
           VAR_CITE = parametro + "-000" + Trim(var_cod)
        End If
        If var_cod > 999 And var_cod < 10000 Then
           VAR_CITE = parametro + "-00" + Trim(var_cod)
        End If
        If var_cod > 9999 And var_cod < 100000 Then
           VAR_CITE = parametro + "-0" + Trim(var_cod)
        End If
        If var_cod > 99999 Then
           VAR_CITE = parametro + "-" + Trim(var_cod)
        End If
'        Ado_datos.Recordset!unidad_codigo_ant = VAR_CITE
'        Ado_datos.Recordset("proceso_codigo") = VAR_N1
'        Ado_datos.Recordset("subproceso_codigo") = VAR_N2
'        Ado_datos.Recordset("etapa_codigo") = VAR_N3
'        Ado_datos.Recordset("poa_codigo") = VAR_POA
'        Ado_datos.Recordset("clasif_codigo") = "ADM"
'        Ado_datos.Recordset("doc_numero") = "0"
'        Ado_datos.Recordset("almacen_codigo") = dtc_codigo11.Text                '2=Almacen Insumos
        If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
            VAR_R = "R-119"
            VAR_ALMD = IIf(dtc_codigo20.Text = "", "0", dtc_codigo20.Text)
            VAR_DPTOD = IIf(dtc_codigo22.Text = "", Left(dtc_codigo3.Text, 1), dtc_codigo22.Text)
        Else
            VAR_R = "R-115"
            VAR_ALMD = "0"
            VAR_DPTOD = "0"
        End If
'        Select Case VAR_ORIGEN
'            Case "UALMI"
'                Ado_datos.Recordset!doc_codigo_alm = VAR_R
'            Case "UALMR"
'                Ado_datos.Recordset!doc_codigo_almR = VAR_R
'            Case "UALMH"
'                Ado_datos.Recordset!doc_codigo_almH = VAR_R
'            Case "GADM"
'                Ado_datos.Recordset!doc_codigo = VAR_R
'            Case Else
'                Ado_datos.Recordset!doc_codigo = VAR_R
'        End Select
'        Ado_datos.Recordset!doc_codigo = VAR_R
        'Ado_datos.Recordset!doc_codigo_alm = VAR_R
        'INI ACTUALIZA CORRELATIVO POR ALMACEN
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        SQL_FOR = "select * from ac_almacenes where almacen_codigo = " & Val(dtc_codigo11.Text) & "  "
        rs_aux5.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux5.RecordCount > 0 Then
           If dtc_aux3.Text = "APR" Then
                rs_aux5!correl_sal = rs_aux5!correl_sal + 1
                VAR_NUM = rs_aux5!correl_sal
           Else
              If Ado_datos.Recordset!edif_codigo = "20101-2" Or Ado_datos.Recordset!edif_codigo = "70101-2" Or Ado_datos.Recordset!edif_codigo = "30101-2" Or Ado_datos.Recordset!edif_codigo = "10101-2" Then
                '===== ini GENERA EL CORRELATIVO POR TRANSFERENCIA DE ALMACEN ====
                Set rs_aux7 = New ADODB.Recordset
                rs_aux7.CursorLocation = adUseClient
                If rs_aux7.State = 1 Then rs_aux7.Close
                'Select Case parametro
                Select Case VAR_ORIGEN
                  Case "UALMI"          ', "ALMIB", "ALMIS", "ALMIC"
                      rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119i'", db, adOpenDynamic, adLockOptimistic
                  Case "UALMR"          ', "ALMRB", "ALMRS", "ALMRC"
                      rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119R'", db, adOpenDynamic, adLockOptimistic
                  Case "UALMH"          ', "ALMHB", "ALMHS", "ALMHC"
                      rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119H'", db, adOpenDynamic, adLockOptimistic
                  Case "GADM"
                      rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119A'", db, adOpenDynamic, adLockOptimistic
                  Case Else
                      rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119i'", db, adOpenDynamic, adLockOptimistic
                End Select
                If rs_aux7.RecordCount > 0 Then
                  VAR_NUM = CDbl(rs_aux7!numero_correlativo) + 1
                  rs_aux7!numero_correlativo = Trim(Str(VAR_NUM))
                  rs_aux7.Update
                End If
                If rs_aux7.State = 1 Then rs_aux7.Close
                '===== fin TERMINA EL CORRELATIVO POR TRANSFERENCIA DE ALMACEN ====
              Else
                rs_aux5!correl_sal = rs_aux5!correl_sal + 1
                VAR_NUM = rs_aux5!correl_sal
              End If
           End If
           rs_aux5.Update
           'db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & parametro & "'   "
        Else
           VAR_NUM = 1
        End If
        
        Select Case VAR_ORIGEN
            Case "UALMI"
                VAR_DOCI = VAR_R
                VAR_DOCR = ""
                VAR_DOCH = ""
                VAR_DOCA = ""
                VAR_BENI = IIf(dtc_codigo4.Text = "", "0", dtc_codigo4.Text)
                VAR_BENR = "0"
                VAR_BENH = "0"
                VAR_BENA = "0"
                VAR_BENDI = IIf(dtc_codigo5.Text = "", "0", dtc_codigo5.Text)
                VAR_BENDR = "0"
                VAR_BENDH = "0"
                VAR_BENDA = "0"
                VAR_NUMI = VAR_NUM
                VAR_NUMR = "0"
                VAR_NUMH = "0"
                VAR_NUMA = "0"
                
                VAR_ALMI = IIf(dtc_codigo11.Text = "", "0", dtc_codigo11.Text)
                VAR_ALMR = "0"
                VAR_ALMH = "0"
                VAR_ALMA = "0"
                
                VAR_ALMDI = IIf(dtc_codigo20.Text = "", "0", dtc_codigo20.Text)
                VAR_ALMDR = "0"
                VAR_ALMDH = "0"
                VAR_ALMDA = "0"
            Case "UALMR"
                VAR_DOCI = ""
                VAR_DOCR = VAR_R
                VAR_DOCH = ""
                VAR_DOCA = ""
                VAR_BENI = "0"
                VAR_BENR = dtc_codigo4.Text
                VAR_BENH = "0"
                VAR_BENA = "0"
                VAR_BENDI = "0"
                VAR_BENDR = dtc_codigo5.Text
                VAR_BENDH = "0"
                VAR_BENDA = "0"
                
                VAR_NUMI = "0"
                VAR_NUMR = VAR_NUM
                VAR_NUMH = "0"
                VAR_NUMA = "0"
                
                VAR_ALMI = "0"
                VAR_ALMR = dtc_codigo11.Text
                VAR_ALMH = "0"
                VAR_ALMA = "0"
                
                VAR_ALMDI = "0"
                VAR_ALMDR = dtc_codigo20.Text
                VAR_ALMDH = "0"
                VAR_ALMDA = "0"
            Case "UALMH"
                VAR_DOCI = ""
                VAR_DOCR = ""
                VAR_DOCH = VAR_R
                VAR_DOCA = ""
                VAR_BENI = "0"
                VAR_BENR = "0"
                VAR_BENH = dtc_codigo4.Text
                VAR_BENA = "0"
                VAR_BENDI = "0"
                VAR_BENDR = "0"
                VAR_BENDH = dtc_codigo5.Text
                VAR_BENDA = "0"
                
                VAR_NUMI = "0"
                VAR_NUMR = "0"
                VAR_NUMH = VAR_NUM
                VAR_NUMA = "0"
                
                VAR_ALMI = "0"
                VAR_ALMR = "0"
                VAR_ALMH = dtc_codigo11.Text
                VAR_ALMA = "0"
                
                VAR_ALMDI = "0"
                VAR_ALMDR = "0"
                VAR_ALMDH = dtc_codigo20.Text
                VAR_ALMDA = "0"
            Case "GADM"
                VAR_DOCI = ""
                VAR_DOCR = ""
                VAR_DOCH = ""
                VAR_DOCA = VAR_R
                VAR_BENI = "0"
                VAR_BENR = "0"
                VAR_BENH = "0"
                VAR_BENA = dtc_codigo4.Text
                VAR_BENDI = "0"
                VAR_BENDR = "0"
                VAR_BENDH = "0"
                VAR_BENDA = dtc_codigo5.Text
                
                VAR_NUMI = "0"
                VAR_NUMR = "0"
                VAR_NUMH = "0"
                VAR_NUMA = VAR_NUM
                
                VAR_ALMI = "0"
                VAR_ALMR = "0"
                VAR_ALMH = "0"
                VAR_ALMA = dtc_codigo11.Text
                
                VAR_ALMDI = "0"
                VAR_ALMDR = "0"
                VAR_ALMDH = "0"
                VAR_ALMDA = dtc_codigo20.Text
            Case Else
                VAR_DOCI = ""
                VAR_DOCR = VAR_R
                VAR_DOCH = ""
                VAR_DOCA = ""
                VAR_BENI = "0"
                VAR_BENR = dtc_codigo4.Text
                VAR_BENH = "0"
                VAR_BENA = "0"
                VAR_BENDI = "0"
                VAR_BENDR = dtc_codigo5.Text
                VAR_BENDH = "0"
                VAR_BENDA = "0"
                
                VAR_NUMI = "0"
                VAR_NUMR = VAR_NUM
                VAR_NUMH = "0"
                VAR_NUMA = "0"
                
                VAR_ALMI = "0"
                VAR_ALMR = dtc_codigo11.Text
                VAR_ALMH = "0"
                VAR_ALMA = "0"
                
                VAR_ALMDI = "0"
                VAR_ALMDR = dtc_codigo20.Text
                VAR_ALMDH = "0"
                VAR_ALMDA = "0"
        End Select
        FVenta = Format(IIf(IsNull(DTPfechasol.Value), Date, DTPfechasol.Value), "dd/mm/yyyy")
        Dim VAR_HORA As String
        VAR_HORA = Time
                              
        db.Execute "INSERT INTO AO_ventas_cabecera (ges_gestion, venta_codigo, depto_codigo, unidad_codigo, solicitud_codigo, edif_codigo, unidad_destino, unidad_codigo_ant, solicitud_codigo_ant, venta_fecha, venta_tipo, beneficiario_codigo, beneficiario_codigo_resp, beneficiario_codigo_cobr, beneficiario_codigo_alm, " & _
                   " beneficiario_codigo_almR, beneficiario_codigo_almH, beneficiario_codigo_tec, beneficiario_codigo_tecR, beneficiario_codigo_tecH, venta_descripcion, venta_cantidad_total, venta_monto_total_bs, venta_monto_total_dol, venta_tipo_cambio, venta_monto_cobrado_bs, venta_monto_cobrado_dol,  " & _
                   " venta_saldo_p_cobrar_bs, venta_saldo_p_cobrar_dol, venta_plazo_dias_calendario, venta_fecha_inicio, venta_fecha_fin, unimed_codigo, venta_cantidad_cobr, unimed_codigo_cobr, mes_inicio_crono, tipoben_codigo, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, poa_codigo,  " & _
                   " doc_codigo_alm, doc_numero_alm, poa_codigo_alm, almacen_codigo, almacen_codigo_d, depto_codigo_d, doc_codigo_almR, doc_numero_almR, almacen_codigoR, almacen_codigo_dR, depto_codigo_dR, doc_codigo_almH, doc_numero_almH, almacen_codigoH, depto_codigo_dH, archivo_respaldo,  " & _
                   " archivo_respaldo_cargado, correl_detalle, correl_cobro_prog, estado_cancelado, estado_alcance, estado_codigo, estado_almacen, usr_codigo, fecha_registro, hora_registro,  estado_codigo_verif, usr_codigo_verif, fecha_verif, literal_a, nro_eqp, tipo_moneda) " & _
            " values ('" & glGestion & "', " & var_cod5 & ", '" & VAR_DPTO & "', '" & parametro & "', " & var_cod & ", '" & dtc_codigo3.Text & "', '" & dtc_codigo2.Text & "', '" & VAR_CITE & "', '0', '" & FVenta & "', 'A', '" & VAR_BENA & "', '" & dtc_codigo4.Text & "', '0', '" & VAR_BENI & "', " & _
            " '" & VAR_BENR & "', '" & VAR_BENH & "', '" & VAR_BENDI & "', '" & VAR_BENDR & "', '" & VAR_BENDH & "', '" & TxtConcepto.Text & "', '0', '0', '0', '" & GlTipoCambioOficial & "', '0', '0',  " & _
            " '0', '0', '0', '" & FVenta & "', '" & FVenta & "', 'MES', '1', 'MES', 'ENERO', '1', '" & VAR_N1 & "', '" & VAR_N2 & "', '" & VAR_N3 & "', 'ADM', '" & VAR_R & "', '0', '" & VAR_POA & "', " & _
            " '" & VAR_DOCI & "', " & VAR_NUMI & ", '" & VAR_POA & "', '" & VAR_ALMI & "', '" & VAR_ALMDI & "', '" & VAR_DPTOD & "', '" & VAR_DOCR & "', " & VAR_NUMR & ", '" & VAR_ALMR & "', '" & VAR_ALMDR & "', '" & VAR_DPTOD & "', '" & VAR_DOCH & "', " & VAR_NUMH & ", '" & VAR_ALMH & "', '" & VAR_DPTOD & "', '" & VAR_CITE & "', " & _
            " 'N', '0', '1', 'N', 'N', 'APR', 'REG', '" & glusuario & "', '" & Date & "', '" & VAR_HORA & "', 'REG', '" & glusuario & "', '" & Date & "', '', '0', 'BOB' ) "

'        Ado_datos.Recordset!doc_numero_alm = VAR_NUM
'        'FIN ACTUALIZA CORRELATIVO POR ALMACEN
'        Ado_datos.Recordset!estado_codigo = "APR"
'        Ado_datos.Recordset!estado_almacen = "REG"
'        Ado_datos.Recordset!usr_codigo = glusuario
'        Ado_datos.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'        Ado_datos.Recordset!hora_registro = Format(Time, "hh/mm/ss")
'        'Ado_datos.Recordset("usuario_aprueba") = ""
'        'Ado_datos.Recordset("fecha_aprueba") = ""
    End If
    If swgrabar = 2 Then
        If Ado_datos.Recordset!doc_numero_alm = 0 Then
            'INI ACTUALIZA CORRELATIVO POR ALMACEN
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            SQL_FOR = "select * from ac_almacenes where almacen_codigo = " & Val(dtc_codigo11.Text) & "  "
            rs_aux5.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If rs_aux5.RecordCount > 0 Then
               If dtc_aux3.Text = "APR" Then
                    rs_aux5!correl_sal = rs_aux5!correl_sal + 1
                    VAR_NUM = rs_aux5!correl_sal
               Else
                    If Ado_datos.Recordset!edif_codigo = "20101-2" Or Ado_datos.Recordset!edif_codigo = "70101-2" Or Ado_datos.Recordset!edif_codigo = "30101-2" Or Ado_datos.Recordset!edif_codigo = "10101-2" Then
                      '===== ini GENERA EL CORRELATIVO POR TRANSFERENCIA DE ALMACEN ====
                        Set rs_aux7 = New ADODB.Recordset
                        rs_aux7.CursorLocation = adUseClient
                        If rs_aux7.State = 1 Then rs_aux7.Close
                        'Select Case parametro
                        Select Case VAR_ORIGEN
                          Case "UALMI"          ', "ALMIB", "ALMIS", "ALMIC"
                              rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119i'", db, adOpenDynamic, adLockOptimistic
                          Case "UALMR"          ', "ALMRB", "ALMRS", "ALMRC"
                              rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119R'", db, adOpenDynamic, adLockOptimistic
                          Case "UALMH"          ', "ALMHB", "ALMHS", "ALMHC"
                              rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119H'", db, adOpenDynamic, adLockOptimistic
                          Case "GADM"
                              rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119A'", db, adOpenDynamic, adLockOptimistic
                          Case Else
                              rs_aux7.Open "select * from fc_Correl  where tipo_tramite = 'R-119i'", db, adOpenDynamic, adLockOptimistic
                        End Select
                        If rs_aux7.RecordCount > 0 Then
                          VAR_NUM = CDbl(rs_aux7!numero_correlativo) + 1
                          rs_aux7!numero_correlativo = Trim(Str(VAR_NUM))
                          rs_aux7.Update
                        End If
                        If rs_aux7.State = 1 Then rs_aux7.Close
                        '===== fin TERMINA EL CORRELATIVO POR TRANSFERENCIA DE ALMACEN ====
                    Else
                      rs_aux5!correl_sal = rs_aux5!correl_sal + 1
                      VAR_NUM = rs_aux5!correl_sal
                    End If
               End If
               rs_aux5.Update
            Else
               VAR_NUM = 1
            End If
            'FIN ACTUALIZA CORRELATIVO POR ALMACEN
            Ado_datos.Recordset!doc_numero_alm = VAR_NUM
        Else
            VAR_NUM = Ado_datos.Recordset!doc_numero_alm
        End If
        'If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
        '    VAR_R = "R-119"
        'Else
        '    VAR_R = "R-115"
        'End If
        
        If dtc_codigo3.Text = "20101-2" Or dtc_codigo3.Text = "30101-2" Or dtc_codigo3.Text = "70101-2" Or dtc_codigo3.Text = "10101-2" Then
            VAR_R = "R-119"
            db.Execute "UPDATE ao_ventas_cabecera SET almacen_codigo_d = '" & dtc_codigo20.Text & "', depto_codigo_d = '" & dtc_codigo22.Text & "'   WHERE venta_codigo = " & var_cod5 & " "
            'db.Execute "update ao_ventas_cabecera set unidad_destino = '" & dtc_codigo2.Text & "', venta_descripcion = '" & Trim(TxtConcepto.Text) & "', beneficiario_codigo_tecR = '" & dtc_codigo5.Text & "' WHERE venta_codigo = " & var_cod5 & " "
        Else
            VAR_R = "R-115"
            'db.Execute "UPDATE ao_ventas_cabecera SET doc_codigo_almR = '" & VAR_R & "', usr_codigo_verif = '" & glusuario & "', fecha_verif = '" & DTPfechasol & "', beneficiario_codigo_almR = '" & dtc_codigo4.Text & "', beneficiario_codigo_tecR= '" & dtc_codigo5.Text & "', almacen_codigoR = " & dtc_codigo11.Text & ", depto_codigo = '" & dtc_codigo21.Text & "', doc_numero_almR = " & VAR_NUM & "   WHERE venta_codigo = " & var_cod5 & " "
            'db.Execute "update ao_ventas_cabecera set unidad_destino = '" & dtc_codigo2.Text & "', venta_descripcion = '" & Trim(TxtConcepto.Text) & "', beneficiario_codigo_tecR = '" & dtc_codigo5.Text & "' WHERE venta_codigo = " & var_cod5 & " "
        End If
        db.Execute "UPDATE ao_ventas_cabecera SET doc_codigo_alm = '" & VAR_R & "', usr_codigo_verif = '" & glusuario & "', fecha_verif = '" & DTPfechasol & "', beneficiario_codigo_alm = '" & dtc_codigo4.Text & "', beneficiario_codigo_tec= '" & dtc_codigo5.Text & "', almacen_codigo = " & dtc_codigo11.Text & ", depto_codigo = '" & dtc_codigo21.Text & "', doc_numero_alm = " & VAR_NUM & "   WHERE venta_codigo = " & var_cod5 & " "
        db.Execute "update ao_ventas_cabecera set unidad_destino = '" & dtc_codigo2.Text & "', venta_descripcion = '" & Trim(TxtConcepto.Text) & "', beneficiario_codigo_tec = '" & dtc_codigo5.Text & "' WHERE venta_codigo = " & var_cod5 & " "
    End If
'       'Entrega de Almacen  'swgrabar = 2   'modificar
'       Ado_datos.Recordset("usr_codigo_verif") = glusuario
'       Ado_datos.Recordset("fecha_verif") = Format(DTPfechasol, "dd/mm/yyyy")
'       Ado_datos.Recordset!doc_codigo_alm = VAR_R        '"R-115"
'       Ado_datos.Recordset("beneficiario_codigo_alm") = dtc_codigo4.Text        'Responsable Almacen
'       Ado_datos.Recordset("beneficiario_codigo_tec") = dtc_codigo5.Text        'Entregado a:
'       Ado_datos.Recordset("almacen_codigo") = IIf(dtc_codigo11.Text = "", "0", dtc_codigo11.Text)
'       Ado_datos.Recordset("almacen_codigo_d") = IIf(dtc_codigo20.Text = "", "0", dtc_codigo20.Text)
'       Ado_datos.Recordset("depto_codigo") = IIf(dtc_codigo21.Text = "", Left(dtc_codigo3.Text, 1), dtc_codigo21.Text)
'       Ado_datos.Recordset("depto_codigo_d") = IIf(dtc_codigo22.Text = "", Left(dtc_codigo3.Text, 1), dtc_codigo22.Text)
'       Ado_datos.Recordset("unidad_destino") = dtc_codigo2.Text
'       Ado_datos.Recordset.Update
       ' GRABA AO_SOLICITUD  ------------------------------------------
    If swgrabar = 1 Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        'SQL_FOR = "select * from ao_solicitud where edif_codigo = '" & dtc_codigo3 & "' AND unidad_codigo = '" & VAR_UNI & "' "
        SQL_FOR = "select * from ao_solicitud  WHERE solicitud_codigo = " & var_cod & " AND unidad_codigo = '" & parametro & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        
        VAR_FECHA = Format(IIf(IsNull(DTPfechasol.Value), Date, DTPfechasol.Value), "dd/mm/yyyy")
         
         'db.Execute "INSERT INTO ao_solicitud () "
         
         db.Execute "INSERT INTO ao_solicitud (ges_gestion, unidad_codigo, solicitud_codigo, solicitud_fecha_solicitud, solicitud_fecha_recepción, solicitud_tipo, edif_codigo, beneficiario_codigo, " & _
         " beneficiario_codigo_resp, beneficiario_codigo_resp2, unidad_codigo_sol, solicitud_justificacion, solicitud_observaciones, proceso_codigo, subproceso_codigo, " & _
         " etapa_codigo, etapa_codigo2, clasif_codigo, doc_codigo, doc_codigo2, doc_numero, doc_numero2, poa_codigo, ges_gestion_ant, unidad_codigo_ant, " & _
         " correl_detalle, correl_edificacion, correl_calculo, correl_persona, correl_cotiza, correl_bitacora, archivo_respaldo, archivo_respaldo_cargado, " & _
         " estado_codigo, estado_etapa2, estado_cotiza, fecha_registro, usr_codigo, usr_codigo_aprueba, fecha_aprueba, fecha_registro2, usr_codigo2 , observacion_proy )  " & _
         " VALUES ('" & glGestion & "', '" & parametro & "', " & var_cod & ", '" & VAR_FECHA & "', '" & VAR_FECHA & "', " & VAR_TIPO & ", '" & dtc_codigo3.Text & "', '" & dtc_codigo5.Text & "', " & _
         " '" & dtc_codigo4.Text & "', '0', '" & parametro & "', '" & TxtConcepto.Text & "', '" & TxtConcepto.Text & "', '" & VAR_N1 & "', '" & VAR_N2 & "', " & _
         " '" & VAR_N3 & "', '', 'ADM', '" & VAR_R & "', '', " & VAR_NUM & ", '0', '" & VAR_POA & "', '" & glGestion & "', '" & VAR_CITE & "', " & _
         " '0', '0', '0', '0', '0', '0', 'sin_nombre', 'N', 'APR', 'REG', 'REG', '" & Date & "', '" & glusuario & "', '', '" & Date & "', '" & Date & "', '', '" & dtc_desc3.Text & "' )"

        'rs_aux1!solicitud_codigo = var_cod
        'rs_aux1!estado_codigo = "APR"      'no cambia
        'rs_aux1!ges_gestion = glGestion        'Year(Date)   'no cambia
        'rs_aux1!unidad_codigo = parametro
        'rs_aux1!unidad_codigo_sol = parametro
        '' Actualiza correaltivo ...
        'rs_aux1!doc_numero = "0"    'txt_campo1.Caption
        ''rs_aux1!correl_edificacion = 0
        'rs_aux1!archivo_respaldo = "sin_nombre"
        'rs_aux1!archivo_respaldo_cargado = "N"
        ' rs_aux1!solicitud_fecha_solicitud = Format(IIf(IsNull(DTPfechasol.Value), Date, DTPfechasol.Value), "dd,mm,yyyy")
         
        ' rs_aux1!solicitud_tipo = VAR_TIPO      '"25"    'dtc_codigo2.Text
        ' rs_aux1!edif_codigo = dtc_codigo3.Text
        ' rs_aux1!beneficiario_codigo = dtc_codigo5.Text        'Entregado a:
         
         'rs_aux1!solicitud_justificacion = Trim(TxtConcepto.Text)
         'rs_aux1!proceso_codigo = VAR_N1
         'rs_aux1!subproceso_codigo = VAR_N2
         'rs_aux1!etapa_codigo = VAR_N3
         'rs_aux1!poa_codigo = VAR_POA
         'rs_aux1!clasif_codigo = "ADM"
         'rs_aux1!doc_codigo = VAR_R
         'rs_aux1!doc_numero = VAR_NUM
         'rs_aux1!solicitud_observaciones = Trim(TxtConcepto.Text)
'         rs_aux1!observacion_proy = ""  'dtc_desc3.Text
         'rs_aux1!solicitud_fecha_recepción = Format(IIf(IsNull(DTPfechasol.Value), Date, DTPfechasol.Value), "dd,mm,yyyy")
         
         'rs_aux1!beneficiario_codigo_resp = dtc_codigo4.Text        'Responsable Almacen
         'rs_aux1!beneficiario_codigo_resp2 = "0"                 'usuario2
         'rs_aux1!ges_gestion_ant = Year(Date)
         'rs_aux1!unidad_codigo_ant = Trim(VAR_CITE)
'         rs_aux1!solicitud_codigo_ant = 0
         'rs_aux1!usr_codigo_aprueba = ""
         'rs_aux1!fecha_aprueba = Date
         'rs_aux1!hora_aprueba = ""
         'rs_aux1!Foto = Date
         'rs_aux1!ARCHIVO_Foto = var_cod + ".JPG"
         'rs_aux1!archivo_foto_cargado = "N"
         'hora_registro
         'rs_aux1!beneficiario_codigo = txt_ci
         'rs_aux1!fecha_registro = Date     'no cambia
       
         'rs_aux1!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
         'rs_aux1.Update    'Batch 'adAffectAll
        'db.Execute "UPDATE ao_ventas_cabecera SET doc_codigo_alm = '" & VAR_R & "' WHERE venta_codigo = " & var_cod5 & " "
        'db.Execute "UPDATE ao_ventas_cabecera SET doc_numero_alm = " & VAR_NUM & " WHERE venta_codigo = " & var_cod5 & " "
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
'  glPersNew = "N"

End Sub

Private Sub OptFilGral1_Click()
   '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND estado_almacen = 'REG' AND LEFT(doc_codigo_alm,5) = '" & Left(VAR_R, 5) & "' "
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND estado_almacen = 'REG' AND ((almacen_tipo = '" & VAR_ALMT & "' AND unidad_codigo <> '" & parametro & "') OR unidad_codigo = '" & parametro & "') AND depto_codigo = '" & VAR_DPTO & "' "
    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND estado_almacen = 'REG' AND ((almacen_tipo = '" & VAR_ALMT & "' AND unidad_codigo <> '" & parametro & "' AND depto_codigo = '" & VAR_DPTO & "') OR unidad_codigo = '" & parametro & "')"
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_codigo_alm, almacen_codigo, doc_numero_alm"
    'rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    BtnAprobar.Visible = True
'  '===== Proceso para filtrado general de datos(registros no aprobados)
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    ''    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND estado_almacen = 'REG' AND unidad_codigo_sol = '" & parametro & "' "
'    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND estado_almacen = 'REG' AND LEFT(doc_codigo_alm,5) = '" & LEFT(VAR_R,5) & "' "
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
'    sino = rs_datos.RecordCount
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
 '===== Proceso para filtrado general de datos (todos los registros )
'    Select Case parametro
'        Case "UALMR"
'            TIPO_ALM = "R"
'        Case "UALMH"
'            TIPO_ALM = "H"
'        Case "UALMI"
'            TIPO_ALM = "I"
'        Case "GADM"
'            TIPO_ALM = "A"
'     End Select
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND (almacen_tipo = '" & VAR_ALMT & "' OR unidad_codigo = '" & parametro & "') AND depto_codigo = '" & VAR_DPTO & "' "
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND LEFT(doc_codigo_alm,5) = '" & Left(VAR_R, 5) & "' "
    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND ((almacen_tipo = '" & VAR_ALMT & "' AND unidad_codigo <> '" & parametro & "' AND depto_codigo = '" & VAR_DPTO & "') OR unidad_codigo = '" & parametro & "')"
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_codigo_alm, almacen_codigo, doc_numero_alm"
    'rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
'  '===== Proceso para filtrado general de datos (todos los registros )
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
''    queryinicial = "select * From av_ventas_cabecera WHERE estado_codigo = 'APR' "
'    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND unidad_codigo_sol = '" & parametro & "' "
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
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
  'rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as VAR_COBR2 from ao_ventas_detalle where ges_gestion = '" & ges & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot0 from ao_ventas_detalle where venta_codigo = " & Nro & " and par_codigo = '43340'", db, adOpenKeyset, adLockOptimistic
  If IsNull(rstacumdet!totbs) Then
    VAR_AUX = 0
    VAR_AUX2 = 0
    VAR_CANT = 1
  Else
    VAR_AUX = Round(rstacumdet!totbs, 2)
    VAR_AUX2 = Round(rstacumdet!totdl, 2)
    VAR_CANT = rstacumdet!cantot0
  End If

  'rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion = '" & ges & "' and estado_codigo = 'APR' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
  If IsNull(rs_datos19!totbs2) Then
    Cobrobs = 0
    VAR_COBR = 0
  Else
    Cobrobs = Round(rs_datos19!totbs2, 2)
    VAR_COBR = Round(rs_datos19!totdl2, 2)
  End If

  VAR_Bs = VAR_AUX - Cobrobs
  VAR_Dol = VAR_AUX2 - VAR_COBR
  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "

  TxtMontoBs.Text = VAR_AUX
  TxtCobrado.Text = Cobrobs
  TxtBstotal.Text = VAR_Bs

'  If IsNull(Ado_datos.Recordset!venta_monto_cobrado_bs) Then
'    Ado_datos.Recordset!venta_monto_cobrado_bs = 0
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs
'  Else
'    VAR_AUX = Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs
'  End If
'  If VAR_AUX > 0 Then
'        VAR_AUX2 = VAR_AUX / Ado_datos.Recordset!venta_tipo_cambio
'  Else
'        VAR_AUX2 = 0
'  End If
'  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas_cabecera.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas_cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas_cabecera.saldo_p_cobrar = ao_ventas_cabecera.monto_total_Bs - ao_ventas_cabecera.deuda_cobrada Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & rstacumdet!totbs & " , ao_ventas_cabecera.venta_monto_total_dol = " & rstacumdet!totdl & ", ao_ventas_cabecera.venta_cantidad_total = " & rstacumdet!cantot & ", ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_AUX & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_AUX2 & "  Where ao_ventas_cabecera.ges_gestion = '" & ges & "' And ao_ventas_cabecera.venta_codigo = " & nro & " "
'
'  TxtMontoBs = rstacumdet!totbs
'  TxtCobrado = rs_datos19!totbs2    'IIf(IsNull(Ado_datos.Recordset("venta_monto_cobrado_bs")), 0, Ado_datos.Recordset("venta_monto_cobrado_bs"))
'  If IsNull(Ado_datos.Recordset("venta_saldo_p_cobrar_bs")) Then
'    Text2 = VAR_AUX 'Ado_datos.Recordset("venta_monto_total_bs") - Ado_datos.Recordset("venta_monto_cobrado_bs")
'    Ado_datos.Recordset("venta_saldo_p_cobrar_bs") = VAR_AUX
'  Else
'    Text2 = Ado_datos.Recordset("venta_saldo_p_cobrar_bs")
'  End If

  If rstacumdet.State = 1 Then rstacumdet.Close

  'Print ado_datos14.Recordset!ges_gestion
  'Print ado_datos14.Recordset!correl_venta
  'Print ado_datos14.Recordset!venta_codigo
  'ado_datos14.Recordset!monto_Bolivianos = rstacumdet!totbs
  'ado_datos14.Recordset!monto_dolares = rstacumdet!totdl
  'ado_datos14.Recordset.Update
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & ges & "' and correl_venta = '" & corr & "' and venta_codigo = " & nro, db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount > 0 Then
'    rstdestino!monto_total_Bs = rstacumdet!totbs
'    rstdestino!monto_cobrado = rstacumdet!totbs
'    rstdestino!monto_total_Us = rstacumdet!totdl
'    rstdestino!cantidad_total_vendida = rstacumdet!cantot
'    rstdestino!saldo_p_cobrar = 0
'    rstdestino.Update
'  End If
'  'Set Ado_datos.Recordset = rstdestino
'  If rstdestino.State = 1 Then rstdestino.Close
'  If rstacumdet.State = 1 Then rstacumdet.Close
End Sub

Private Sub Option1_Click()
Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND (almacen_tipo = '" & VAR_ALMT & "' OR unidad_codigo = '" & parametro & "') and edif_descripcion LIKE 'TRASPASO%'  AND depto_codigo = '" & VAR_DPTO & "'"
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND LEFT(doc_codigo_alm,5) = '" & Left(VAR_R, 5) & "' "
    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND ((almacen_tipo = '" & VAR_ALMT & "' AND unidad_codigo <> '" & parametro & "' AND depto_codigo = '" & VAR_DPTO & "') OR unidad_codigo = '" & parametro & "') and edif_descripcion LIKE 'TRASPASO%'"
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Option2_Click()
Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND (almacen_tipo = '" & VAR_ALMT & "' OR unidad_codigo = '" & parametro & "') and  (NOT (edif_descripcion LIKE 'TRASPASO%')) AND depto_codigo = '" & VAR_DPTO & "' "
    'queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND LEFT(doc_codigo_alm,5) = '" & Left(VAR_R, 5) & "' "
    queryinicial = "select * From av_ventas_cabecera_sol_alm WHERE estado_codigo = 'APR' AND ((almacen_tipo = '" & VAR_ALMT & "' AND unidad_codigo <> '" & parametro & "' AND depto_codigo = '" & VAR_DPTO & "') OR unidad_codigo = '" & parametro & "') and  (NOT (edif_descripcion LIKE 'TRASPASO%')) "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "doc_codigo_alm, almacen_codigo, doc_numero_alm"
    'rs_datos.Sort = "doc_codigo_alm, doc_numero_alm"
    'rs_datos.Sort = "unidad_codigo, SOLICITUD_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
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

Private Sub txt_descripcion_venta_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'KeyAscii = 0
'Call CmdGrabaDet_Click
'Call BtnAddDetalle_Click
'End If
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

Private Sub txtDoc_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Then      '(KeyAscii = 8) Or '(0..9)
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtDsctoTot_LostFocus()
    If TxtDsctoTot.Text = "" Or TxtDsctoTot.Text = "0" Or TxtDsctoTot.Text = "0.00" Then
        TxtMonto.Text = "0"
    Else
        TxtMonto.Text = Round(CDbl(TxtDsctoTot.Text) * GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 46) Or (KeyAscii = 44) Then     '(KeyAscii = 8) Or
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
  '? . , 09
  ',.01234856789
End Sub

Private Sub TxtMonto_LostFocus()
    If TxtMonto.Text = "" Or TxtMonto.Text = "0" Or TxtMonto.Text = "0.00" Then
        TxtDsctoTot.Text = "0"
    Else
        TxtDsctoTot.Text = Round(CDbl(TxtMonto.Text) / GlTipoCambioMercado, 2)
    End If
End Sub

Private Sub TxtPlazo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]" Or KeyAscii = 8, KeyAscii, 0)
End Sub

Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    KeyAscii = 0
    Call CmdGrabaDet_Click
    Call BtnAddDetalle_Click
    'txt_descripcion_venta.SetFocus
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 46 Then
Exit Sub
Else
KeyAscii = 0
End If
End Sub

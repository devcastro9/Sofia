VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_identificacion_cliente 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Procesos Administrativos - Area Técnica - Identificación del Cliente"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   18585
   Icon            =   "tw_identificacion_cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraImprimeRepara 
      BackColor       =   &H80000018&
      Caption         =   "Imprime Cotizacion de Reparaciones"
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
      Height          =   6615
      Left            =   9480
      TabIndex        =   112
      Top             =   720
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton btnPanelSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   5400
         TabIndex        =   120
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton btnPanelImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   2880
         TabIndex        =   119
         Top             =   3960
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "1. Con datos del Cliente (Nombre, Cargo, Institución, etc.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   116
         Top             =   960
         Width           =   5655
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000018&
         Caption         =   "2. Sólo con Nombre de Edificio (Sin Datos del Cliente)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   115
         Top             =   1440
         Width           =   5535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000018&
         Caption         =   "3. Con datos del Cliente (Nombre, Cargo, Institución, etc.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   114
         Top             =   2760
         Width           =   5655
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000018&
         Caption         =   "4. Sólo con Nombre de Edificio (Sin Datos del Cliente)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   113
         Top             =   3240
         Width           =   5535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Imprime en Hojas Membretadas (para Emitir por Impresora)"
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
         Height          =   255
         Left            =   600
         TabIndex        =   118
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         Caption         =   "Imprime en Hojas Sin Membretes (para Emitir por PDF o Impresora)"
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
         Height          =   255
         Left            =   600
         TabIndex        =   117
         Top             =   2280
         Width           =   7215
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   580
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   20280
      TabIndex        =   54
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6720
         Picture         =   "tw_identificacion_cliente.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   76
         ToolTipText     =   "Lista Mantenimientos Gratuitos"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8040
         Picture         =   "tw_identificacion_cliente.frx":130B
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   58
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8040
         Picture         =   "tw_identificacion_cliente.frx":1B3E
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H80000015&
         Caption         =   "Ver"
         Height          =   600
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Lista Mantenimientos Gratuitos"
         Top             =   10
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_identificacion_cliente.frx":2535
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   61
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
         Picture         =   "tw_identificacion_cliente.frx":2CF4
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   60
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         DataSource      =   "Ado_datos"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "tw_identificacion_cliente.frx":3609
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   59
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "tw_identificacion_cliente.frx":3D55
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   57
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         Picture         =   "tw_identificacion_cliente.frx":450A
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   56
         ToolTipText     =   "Listado de Trámites Iniciados para Cotizacion"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "tw_identificacion_cliente.frx":4DD7
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   55
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
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12615
         TabIndex        =   63
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Height          =   8115
      Left            =   7665
      TabIndex        =   11
      Top             =   660
      Visible         =   0   'False
      Width           =   11175
      Begin VB.Frame fra_cliente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CLIENTE (Registra una de las 3 alternativas) !!!"
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
         Height          =   1605
         Left            =   120
         TabIndex        =   82
         Top             =   1560
         Width           =   10815
         Begin VB.TextBox Text2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   9120
            TabIndex        =   110
            Top             =   240
            Width           =   75
         End
         Begin VB.TextBox txt_nombre 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   3765
            TabIndex        =   86
            Top             =   300
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.TextBox txt_obs3 
            BackColor       =   &H00FFFFFF&
            DataField       =   "observaciones3"
            DataSource      =   "Ado_datos"
            Height          =   795
            Left            =   3225
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   720
            Width           =   5850
         End
         Begin VB.CommandButton BtnAux2 
            BackColor       =   &H00C0FFFF&
            Height          =   705
            Left            =   9255
            MaskColor       =   &H00FFFFFF&
            Picture         =   "tw_identificacion_cliente.frx":5599
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Registrar un NUEVO Cliente"
            Top             =   285
            Width           =   1380
         End
         Begin VB.TextBox txt_ci 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000004&
            Height          =   285
            Left            =   9240
            TabIndex        =   83
            Top             =   1245
            Visible         =   0   'False
            Width           =   1365
         End
         Begin MSDataListLib.DataCombo dtc_desc4 
            Bindings        =   "tw_identificacion_cliente.frx":6370
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3840
            TabIndex        =   109
            Top             =   285
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo4 
            Bindings        =   "tw_identificacion_cliente.frx":6389
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9240
            TabIndex        =   111
            Top             =   960
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "1. Si Existe en la Base de Datos, elija-->"
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
            Left            =   120
            TabIndex        =   88
            Top             =   285
            Width           =   3525
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "3. Destinatario de la Nota ó Datos Referenciales (Nombre, Institucion, Cargo, Telef.)"
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
            Height          =   1440
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   3120
         End
      End
      Begin VB.PictureBox FraGrabarCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   120
         ScaleHeight     =   705
         ScaleWidth      =   10920
         TabIndex        =   78
         Top             =   7320
         Visible         =   0   'False
         Width           =   10920
         Begin VB.PictureBox BtnCancelar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5715
            Picture         =   "tw_identificacion_cliente.frx":63A2
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   80
            Top             =   60
            Width           =   1455
         End
         Begin VB.PictureBox BtnGrabar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "tw_identificacion_cliente.frx":6C8E
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   79
            Top             =   60
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
            Left            =   12735
            TabIndex        =   81
            Top             =   195
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin VB.TextBox TxtPlazo 
         DataField       =   "PlazoDias"
         Height          =   285
         Left            =   9120
         TabIndex        =   75
         Text            =   "2"
         Top             =   4320
         Width           =   520
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "tw_identificacion_cliente.frx":7464
         DataField       =   "TipoContratoCodigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   25
         Top             =   4320
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "TipoContratoCodigo"
         BoundColumn     =   "TipoContratoCodigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   675
         Left            =   1515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4800
         Width           =   9405
      End
      Begin VB.TextBox txt_obs2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "observaciones2"
         DataSource      =   "Ado_datos"
         Height          =   675
         Left            =   1520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   74
         Top             =   6480
         Width           =   9405
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   645
         Left            =   1520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   5640
         Width           =   9405
      End
      Begin VB.TextBox Txt_campo3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "doc_numero2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8710
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   1080
         Width           =   2205
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "tw_identificacion_cliente.frx":747E
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4275
         TabIndex        =   18
         Top             =   5700
         Visible         =   0   'False
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
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
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7080
         TabIndex        =   41
         Top             =   570
         Width           =   285
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "tw_identificacion_cliente.frx":7497
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "tw_identificacion_cliente.frx":74B1
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   29
         Top             =   240
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
         Bindings        =   "tw_identificacion_cliente.frx":74CA
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9600
         TabIndex        =   19
         Top             =   5640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "tw_identificacion_cliente.frx":74E3
         DataField       =   "TipoContratoCodigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1875
         TabIndex        =   3
         Top             =   4320
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "DescripcionTipoContrato"
         BoundColumn     =   "TipoContratoCodigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "tw_identificacion_cliente.frx":74FD
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_identificacion_cliente.frx":7516
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   1080
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
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_identificacion_cliente.frx":752F
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_identificacion_cliente.frx":7548
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   555
         Width           =   5685
         _ExtentX        =   10028
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
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "tw_identificacion_cliente.frx":7561
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   3840
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   8715
         TabIndex        =   45
         Top             =   3720
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   109838337
         CurrentDate     =   44860
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_identificacion_cliente.frx":757B
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   900
         TabIndex        =   53
         Top             =   1095
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "tw_identificacion_cliente.frx":7594
         DataField       =   "codigo_empresa"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1860
         TabIndex        =   121
         Top             =   3360
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_empresa"
         BoundColumn     =   "codigo_empresa"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "tw_identificacion_cliente.frx":75AD
         DataField       =   "codigo_empresa"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   122
         Top             =   3360
         Visible         =   0   'False
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_empresa"
         BoundColumn     =   "codigo_empresa"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
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
         Index           =   4
         Left            =   180
         TabIndex        =   123
         Top             =   3360
         Width           =   915
      End
      Begin VB.Label LblPlazo 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Height          =   255
         Left            =   8400
         TabIndex        =   77
         Top             =   4350
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Penúltimo Párrafo. . . . . :"
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
         Height          =   480
         Left            =   180
         TabIndex        =   73
         Top             =   6600
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Justificación Técnica . . . . :"
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
         Height          =   480
         Left            =   180
         TabIndex        =   72
         Top             =   5760
         Width           =   1245
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cite TEC"
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
         Height          =   240
         Index           =   3
         Left            =   7560
         TabIndex        =   70
         Top             =   1110
         Width           =   1170
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
         Left            =   2580
         TabIndex        =   43
         Top             =   5700
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Trámite:"
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
         Index           =   1
         Left            =   180
         TabIndex        =   42
         Top             =   4350
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Adm./File.Contrato"
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
         Left            =   7680
         TabIndex        =   40
         Top             =   300
         Width           =   2070
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   7800
         TabIndex        =   39
         Top             =   555
         Width           =   1815
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
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
         Left            =   180
         TabIndex        =   36
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lbl_descripcion 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto (para Ref.). . :"
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
         Height          =   480
         Left            =   180
         TabIndex        =   35
         Top             =   4935
         Width           =   1275
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Registro ISO"
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
         Left            =   1500
         TabIndex        =   34
         Top             =   5640
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lbl_campo11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable CGI:"
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
         Left            =   180
         TabIndex        =   33
         Top             =   3870
         Width           =   1815
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
         Left            =   1725
         TabIndex        =   32
         Top             =   300
         Width           =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   11155
         Y1              =   915
         Y2              =   915
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
         Left            =   180
         TabIndex        =   27
         Top             =   555
         Width           =   1215
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
         Left            =   5460
         TabIndex        =   26
         Top             =   4380
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Doc. ISO"
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
         Left            =   3360
         TabIndex        =   21
         Top             =   5640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud:"
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
         Index           =   12
         Left            =   8700
         TabIndex        =   20
         Top             =   3390
         Width           =   1425
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10065
         TabIndex        =   4
         Top             =   555
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "#Trámite"
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
         TabIndex        =   13
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         Left            =   10155
         TabIndex        =   12
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.PictureBox BtnImprimir7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   17040
      Picture         =   "tw_identificacion_cliente.frx":75C6
      ScaleHeight     =   585
      ScaleWidth      =   1395
      TabIndex        =   69
      ToolTipText     =   "Cotización del Servicio"
      Top             =   5520
      Width           =   1400
   End
   Begin VB.PictureBox BtnImprimir4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   17040
      Picture         =   "tw_identificacion_cliente.frx":7F5C
      ScaleHeight     =   585
      ScaleWidth      =   1395
      TabIndex        =   68
      ToolTipText     =   "Cotización del Servicio"
      Top             =   3240
      Width           =   1400
   End
   Begin VB.PictureBox BtnImprimir2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   17040
      Picture         =   "tw_identificacion_cliente.frx":8933
      ScaleHeight     =   585
      ScaleWidth      =   1395
      TabIndex        =   67
      ToolTipText     =   "Cotización del Servicio"
      Top             =   960
      Width           =   1400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Herramientas"
      Height          =   635
      Left            =   17220
      Picture         =   "tw_identificacion_cliente.frx":9317
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Cotizacion y Costos del Servicio"
      Top             =   7840
      Width           =   1365
   End
   Begin VB.CommandButton BtnImprimir3 
      Caption         =   "Insumos"
      Height          =   635
      Left            =   7695
      Picture         =   "tw_identificacion_cliente.frx":AA99
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Cotizacion y Costos del Servicio"
      Top             =   7840
      Width           =   1365
   End
   Begin VB.Frame FraDet7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SERVICIO TECNICO EXTERNO O INTERNO"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   9435
      TabIndex        =   51
      Top             =   5280
      Width           =   9480
      Begin VB.PictureBox FrmABMDet7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   9225
         TabIndex        =   97
         Top             =   240
         Width           =   9255
         Begin VB.PictureBox BtnAddDetalle7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   50
            Picture         =   "tw_identificacion_cliente.frx":C21B
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   100
            Top             =   40
            Width           =   1220
         End
         Begin VB.PictureBox BtnModDetalle7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   1320
            Picture         =   "tw_identificacion_cliente.frx":C9DA
            ScaleHeight     =   585
            ScaleWidth      =   1425
            TabIndex        =   99
            Top             =   40
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   3045
            Picture         =   "tw_identificacion_cliente.frx":D2EF
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   98
            Top             =   0
            Width           =   1220
         End
      End
      Begin MSDataGridLib.DataGrid dg_det7 
         Bindings        =   "tw_identificacion_cliente.frx":DA3B
         Height          =   1215
         Left            =   75
         TabIndex        =   52
         Top             =   945
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12632319
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.Refer."
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
            Caption         =   "Descripcion Servicio"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
         BeginProperty Column06 
            DataField       =   "observacion"
            Caption         =   "Descripcion_para_Cliente"
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
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2475.213
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1995.024
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SOLICITUD DE HERRAMIENTAS (COSTOS)"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   9330
      TabIndex        =   47
      Top             =   7590
      Width           =   9600
      Begin VB.PictureBox FrmABMDet6 
         Appearance      =   0  'Flat
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   9345
         TabIndex        =   101
         Top             =   240
         Width           =   9375
         Begin VB.CommandButton BtnAnlDetalle6 
            Height          =   640
            Left            =   3120
            Picture         =   "tw_identificacion_cliente.frx":DA56
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Anula Producto Elegido"
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton BtnModDetalle6 
            Height          =   640
            Left            =   1440
            Picture         =   "tw_identificacion_cliente.frx":E24E
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Modifica Producto Elegido"
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton BtnAddDetalle6 
            Height          =   640
            Left            =   50
            Picture         =   "tw_identificacion_cliente.frx":EC63
            Style           =   1  'Graphical
            TabIndex        =   102
            ToolTipText     =   "Adiciona Producto"
            Top             =   0
            Width           =   1365
         End
      End
      Begin MSDataGridLib.DataGrid dg_det6 
         Bindings        =   "tw_identificacion_cliente.frx":F513
         Height          =   1215
         Left            =   60
         TabIndex        =   48
         Top             =   960
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.Refer."
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
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SOLICITUD DE REPUESTOS (COSTOS)"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   9450
      TabIndex        =   46
      Top             =   2940
      Width           =   9480
      Begin VB.PictureBox FrmABMDet5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   9225
         TabIndex        =   93
         Top             =   240
         Width           =   9255
         Begin VB.PictureBox BtnAddDetalle5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   50
            Picture         =   "tw_identificacion_cliente.frx":F52E
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   96
            Top             =   40
            Width           =   1220
         End
         Begin VB.PictureBox BtnModDetalle5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   1420
            Picture         =   "tw_identificacion_cliente.frx":FCED
            ScaleHeight     =   585
            ScaleWidth      =   1425
            TabIndex        =   95
            Top             =   40
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   3000
            Picture         =   "tw_identificacion_cliente.frx":10602
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   94
            Top             =   0
            Width           =   1220
         End
      End
      Begin MSDataGridLib.DataGrid dg_det5 
         Bindings        =   "tw_identificacion_cliente.frx":10D4E
         Height          =   1215
         Left            =   60
         TabIndex        =   49
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.BOB."
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
            DataField       =   "bien_total_compra"
            Caption         =   "Precio.USD"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "observacion"
            Caption         =   "Descripcion_para_Cliente"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   2459.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2009.764
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SOLICITUD DE INSUMOS (COSTOS)"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   75
      TabIndex        =   28
      Top             =   7590
      Width           =   9240
      Begin VB.PictureBox FrmABMDet3 
         Appearance      =   0  'Flat
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   8985
         TabIndex        =   105
         Top             =   240
         Width           =   9015
         Begin VB.CommandButton BtnAddDetalle3 
            Height          =   640
            Left            =   50
            Picture         =   "tw_identificacion_cliente.frx":10D69
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Adiciona Producto"
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton BtnModDetalle3 
            Height          =   640
            Left            =   1440
            Picture         =   "tw_identificacion_cliente.frx":11619
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Modifica Producto Elegido"
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton BtnAnlDetalle3 
            Height          =   640
            Left            =   2805
            Picture         =   "tw_identificacion_cliente.frx":1202E
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Anula Producto Elegido"
            Top             =   0
            Width           =   1365
         End
      End
      Begin MSDataGridLib.DataGrid dg_det3 
         Bindings        =   "tw_identificacion_cliente.frx":12826
         Height          =   1215
         Left            =   60
         TabIndex        =   44
         Top             =   960
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2143
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.Refer."
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
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
         BeginProperty Column06 
            DataField       =   "observacion"
            Caption         =   "Descripcion para el Cliente"
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
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   2324.977
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1800
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_det4 
         Height          =   855
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EQUIPOS QUE INTERVIENEN EN EL SERVICIO"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   9435
      TabIndex        =   22
      Top             =   660
      Width           =   9480
      Begin VB.PictureBox FrmABMDet2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   9225
         TabIndex        =   89
         Top             =   240
         Width           =   9255
         Begin VB.PictureBox BtnAddDetalle2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   50
            Picture         =   "tw_identificacion_cliente.frx":12841
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   92
            Top             =   40
            Width           =   1220
         End
         Begin VB.PictureBox BtnModDetalle2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   1400
            Picture         =   "tw_identificacion_cliente.frx":13000
            ScaleHeight     =   585
            ScaleWidth      =   1425
            TabIndex        =   91
            Top             =   40
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   2925
            Picture         =   "tw_identificacion_cliente.frx":13915
            ScaleHeight     =   585
            ScaleWidth      =   1215
            TabIndex        =   90
            Top             =   0
            Width           =   1220
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "tw_identificacion_cliente.frx":14061
         Height          =   1215
         Left            =   60
         TabIndex        =   23
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo "
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
            DataField       =   "bien_codigo_anterior"
            Caption         =   "Nro.Eqp."
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.Servicio"
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
         BeginProperty Column03 
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Hrs.X Día"
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
            DataField       =   "marca_codigo"
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
         BeginProperty Column05 
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo"
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
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas/Identificacion.Ubicacion"
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
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3179.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00800000&
      Height          =   6825
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   9255
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "tw_identificacion_cliente.frx":1407C
         Height          =   6135
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   10821
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "#Trámite"
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
            DataField       =   "observacion_proy"
            Caption         =   "Nombre.Edificio"
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
            DataField       =   "doc_numero2"
            Caption         =   "No.CiteTec"
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
            DataField       =   "solicitud_fecha_solicitud"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3525.166
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
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
         Left            =   2280
         TabIndex        =   37
         Top             =   6465
         Value           =   -1  'True
         Width           =   1455
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
         Left            =   5640
         TabIndex        =   38
         Top             =   6465
         Width           =   1155
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   6405
         Width           =   8985
         _ExtentX        =   15849
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
      ScaleWidth      =   20160
      TabIndex        =   5
      Top             =   10935
      Width           =   20160
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
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
      Left            =   9840
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_datos2 
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
      Left            =   4680
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
      Left            =   6960
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
      Left            =   9240
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
      Left            =   11520
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
      Left            =   13800
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
      Left            =   2400
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
      Left            =   4680
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
      Left            =   11520
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
      Left            =   13800
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
      Left            =   6960
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
   Begin MSAdodcLib.Adodc Ado_detalle7 
      Height          =   330
      Left            =   9240
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
      Caption         =   "Ado_detalle7"
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
      Left            =   120
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
   Begin MSAdodcLib.Adodc Ado_detalle4 
      Height          =   330
      Left            =   2400
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
      Caption         =   "Ado_detalle4"
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
   Begin MSAdodcLib.Adodc Ado_detalle5 
      Height          =   330
      Left            =   4680
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
      Caption         =   "Ado_detalle5"
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
   Begin MSAdodcLib.Adodc Ado_detalle6 
      Height          =   330
      Left            =   6960
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
      Caption         =   "Ado_detalle6"
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
      Left            =   10320
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "Propuesta (Página 1)"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   690
      WindowHeight    =   370
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   10800
      Top             =   10320
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
   Begin Crystal.CrystalReport CR00 
      Left            =   9360
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
   Begin Crystal.CrystalReport CR05 
      Left            =   11760
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   0
      WindowTop       =   370
      WindowWidth     =   690
      WindowHeight    =   370
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR06 
      Left            =   12240
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   690
      WindowTop       =   370
      WindowWidth     =   690
      WindowHeight    =   370
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR04 
      Left            =   11280
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   690
      WindowTop       =   0
      WindowWidth     =   690
      WindowHeight    =   370
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR07 
      Left            =   12720
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
End
Attribute VB_Name = "tw_identificacion_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset
Dim rs_datos12 As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset
Dim rs_det6 As New ADODB.Recordset
Dim rs_det7 As New ADODB.Recordset

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

Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod, VAR_DET As String
Dim VAR_VAL, VAR_SUBP As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_DA, VAR_UORIGEN As String
Dim VAR_DPTO, VAR_DPTOC As String
Dim VAR_TIT, VAR_SUBT As String
Dim var_literal As String
Dim VAR_PLAZO As String

Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_TOTBS As Double

Dim VAR_TIPO, VAR_SOL As Integer
Dim iResult, VAR_CITES As Integer

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle2_Click()
  
   VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
   VAR_SUBP = Ado_datos.Recordset!subproceso_codigo
  If rs_datos!estado_codigo = "REG" Then
  
    swnuevo = 1
    fraOpciones.Visible = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Select Case VAR_SUBP        'dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "COM-01"    '4. COMPRA-VENTA DE EQUIPOS
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
'            mw_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
'            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Show vbModal
        Case "COM-02"    '3. VENTA DE SERVICIOS (PROVISION)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Show vbModal
        Case "COM-03"    '4. VENTA DE SERVICIOS (INSTALACIONES)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Show vbModal
        Case "COM-04"    '5. VENTA DE SERVICIOS (AJUSTE)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Show vbModal
        Case "TEC-01"    '6. VENTA DE SERVICIOS (MANTENIMIENTO GRATUITO)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.Show vbModal

        Case "TEC-02"    '10. VENTA DE SERVICIOS MANTENIMIENTO PREVENTIVO
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.Show vbModal
        Case "TEC-03"    '7. VENTA DE SERVICIOS REPARACION
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.Show vbModal
        Case "TEC-04"    '8. VENTA DE SERVICIOS (EMERGENCIAS)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.Show vbModal
        Case "TEC-05"    '9. SERVICIO MODERNIZACION    End Select
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            tw_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes.lbl_det.Caption = "43340"
            tw_solicitud_bienes.Txt_estado.Caption = "REG"
            tw_solicitud_bienes.Show vbModal
        End Select
        
    


       If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
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
    swnuevo = 0
    fraOpciones.Visible = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAddDetalle3_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        VAR_DET = "30000"
        sino = MsgBox("Desea cargar el Kit de Insumos ? ", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            Call CARGAR_KIT3
        Else
            Call NuevoDetalle
        End If
        VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
    
         If OptFilGral1.Value = True Then
            Call OptFilGral1_Click        'Pendientes
         Else
            Call OptFilGral2_Click        'TODOS
         End If
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
        ''grupo_codigo = '30000' and (par_codigo <> '39800' and par_codigo <> '34800')
    Else
        MsgBox "NO puede adicionar un NUEVO INSUMO, porque previamente debe registrar... " + FraDet2.Caption, vbExclamation
    End If
End Sub

Private Sub CARGAR_KIT3()
    'gc_unidad_ejecutora
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "Select * from ac_bienes where kit = 'I01' ", db, adOpenStatic
    If rs_datos12.RecordCount > 0 Then
        rs_datos12.MoveFirst
        While Not rs_datos12.EOF
            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & dtc_codigo1.Text & "' and solicitud_codigo = " & txt_codigo.Caption & "   ", db, adOpenKeyset, adLockOptimistic
            rs_aux7.AddNew
            rs_aux7("ges_gestion").Value = glGestion
            rs_aux7("unidad_codigo").Value = dtc_codigo1.Text
            rs_aux7("solicitud_codigo").Value = txt_codigo.Caption
            rs_aux7("estado_codigo").Value = "REG"
            rs_aux7("venta_o_compra").Value = "V"
          
            rs_aux7("bien_codigo").Value = IIf(rs_datos12!bien_codigo = "", "NN", rs_datos12!bien_codigo)
            rs_aux7("marca_codigo").Value = IIf(rs_datos12!marca_codigo = "", "S/M", rs_datos12!marca_codigo)
            rs_aux7("modelo_codigo").Value = IIf(rs_datos12!modelo_codigo = "", "S/M", rs_datos12!modelo_codigo)
                
            rs_aux7("grupo_codigo").Value = IIf(rs_datos12!grupo_codigo = "", "30000", rs_datos12!grupo_codigo)
            rs_aux7("subgrupo_codigo").Value = IIf(rs_datos12!subgrupo_codigo = "", "34000", rs_datos12!subgrupo_codigo)
            rs_aux7("par_codigo").Value = IIf(rs_datos12!par_codigo = "", "34800", rs_datos12!par_codigo)
            
            rs_aux7("bien_precio_venta_base").Value = IIf(rs_datos12!bien_precio_venta_final = "", 0, rs_datos12!bien_precio_venta_final)
            rs_aux7("unimed_codigo").Value = IIf(rs_datos12!unimed_codigo = "", "PZA", rs_datos12!unimed_codigo)
            rs_aux7("bien_cantidad").Value = 1      'IIf(Txt_campo16 = "", 1, Txt_campo16)
            rs_aux7("bien_total_venta").Value = IIf(rs_datos12!bien_precio_venta_final = "", 0, rs_datos12!bien_precio_venta_final)
            rs_aux7("bien_precio_compra").Value = 0
            rs_aux7("bien_total_compra").Value = 0
            
            rs_aux7("fosa_dimension_frente").Value = 7      'IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            rs_aux7("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            rs_aux7("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            rs_aux7("usr_codigo").Value = glusuario
            rs_aux7.UpdateBatch adAffectAll
            
         rs_datos12.MoveNext
         Wend
    End If
    
End Sub

Private Sub NuevoDetalle()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Visible = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Call ABRIR_TABLA_DET
    Select Case VAR_DET
        Case "30000"
        'If VAR_DET = "30000" Then
            Ado_detalle3.Recordset.AddNew
            tw_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes3.lbl_det.Caption = VAR_DET     '"34110"
            tw_solicitud_bienes3.Txt_estado.Caption = "REG"
            GlExtension = Ado_detalle2.Recordset!bien_codigo
            tw_solicitud_bienes3.Show vbModal
        'End If
        Case "39800"
        'If VAR_DET = "39800" Then
            Ado_detalle5.Recordset.AddNew
            tw_solicitud_bienes5.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes5.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes5.lbl_det.Caption = VAR_DET     '"34110"
            tw_solicitud_bienes5.Txt_estado.Caption = "REG"
            GlExtension = Ado_detalle2.Recordset!bien_codigo
            tw_solicitud_bienes5.Show vbModal
        'End If
        Case "34800"
        'If VAR_DET = "34800" Then
            Ado_detalle6.Recordset.AddNew
            tw_solicitud_bienes6.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes6.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes6.lbl_det.Caption = VAR_DET     '"34110"
            tw_solicitud_bienes6.Txt_estado.Caption = "REG"
            tw_solicitud_bienes6.Show vbModal
        'End If
        Case "24300"
        'If VAR_DET = "24300" Then
            Ado_detalle7.Recordset.AddNew
            tw_solicitud_bienes7.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes7.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes7.lbl_det.Caption = VAR_DET     '"34110"
            tw_solicitud_bienes7.Txt_estado.Caption = "REG"
            GlExtension = Ado_detalle2.Recordset!bien_codigo
            tw_solicitud_bienes7.Show vbModal
        'End If
        Case Else
            Ado_detalle5.Recordset.AddNew
            tw_solicitud_bienes5.txt_codigo.Caption = Me.txt_codigo.Caption
            tw_solicitud_bienes5.Txt_campo1.Caption = Me.dtc_codigo1.Text
            tw_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
            tw_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
            tw_solicitud_bienes5.lbl_det.Caption = VAR_DET     '"34110"
            tw_solicitud_bienes5.Txt_estado.Caption = "REG"
            GlExtension = Ado_detalle2.Recordset!bien_codigo
            tw_solicitud_bienes5.Show vbModal
    End Select
            
    swnuevo = 0
    Call ABRIR_TABLA_DET
    fraOpciones.Visible = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
   
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnAddDetalle5_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        VAR_DET = "39800"
        swnuevo = "1"
        Call NuevoDetalle
        VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
         If OptFilGral1.Value = True Then
            Call OptFilGral1_Click        'Pendientes
         Else
            Call OptFilGral2_Click        'TODOS
         End If
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
    Else
        MsgBox "NO puede adicionar un NUEVO REPUESTO, porque previamente debe registrar... " + FraDet2.Caption, vbExclamation
    End If
End Sub

Private Sub BtnAddDetalle6_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        VAR_DET = "34800"
        sino = MsgBox("Desea cargar el Kit de Herramientas ? ", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            Call CARGAR_KIT6
        Else
            Call NuevoDetalle
        End If
        VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
    
         If OptFilGral1.Value = True Then
            Call OptFilGral1_Click        'Pendientes
         Else
            Call OptFilGral2_Click        'TODOS
         End If
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
    Else
        MsgBox "NO puede adicionar una NUEVA HERRAMIENTA, porque previamente debe registrar... " + FraDet2.Caption, vbExclamation
    End If
End Sub

Private Sub CARGAR_KIT6()
    'gc_unidad_ejecutora
    Set rs_datos12 = New ADODB.Recordset
    If rs_datos12.State = 1 Then rs_datos12.Close
    rs_datos12.Open "Select * from ac_bienes where kit = 'H01' ", db, adOpenStatic
    If rs_datos12.RecordCount > 0 Then
        rs_datos12.MoveFirst
        While Not rs_datos12.EOF
            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & dtc_codigo1.Text & "' and solicitud_codigo = " & txt_codigo.Caption & "   ", db, adOpenKeyset, adLockOptimistic
            rs_aux7.AddNew
            rs_aux7("ges_gestion").Value = glGestion
            rs_aux7("unidad_codigo").Value = dtc_codigo1.Text
            rs_aux7("solicitud_codigo").Value = txt_codigo.Caption
            rs_aux7("estado_codigo").Value = "REG"
            rs_aux7("venta_o_compra").Value = "V"
          
            rs_aux7("bien_codigo").Value = IIf(rs_datos12!bien_codigo = "", "NN", rs_datos12!bien_codigo)
            rs_aux7("marca_codigo").Value = IIf(rs_datos12!marca_codigo = "", "S/M", rs_datos12!marca_codigo)
            rs_aux7("modelo_codigo").Value = IIf(rs_datos12!modelo_codigo = "", "S/M", rs_datos12!modelo_codigo)
                
            rs_aux7("grupo_codigo").Value = IIf(rs_datos12!grupo_codigo = "", "30000", rs_datos12!grupo_codigo)
            rs_aux7("subgrupo_codigo").Value = IIf(rs_datos12!subgrupo_codigo = "", "34000", rs_datos12!subgrupo_codigo)
            rs_aux7("par_codigo").Value = IIf(rs_datos12!par_codigo = "", "34800", rs_datos12!par_codigo)
            
            rs_aux7("bien_precio_venta_base").Value = IIf(rs_datos12!bien_precio_venta_final = "", 0, rs_datos12!bien_precio_venta_final)
            rs_aux7("unimed_codigo").Value = IIf(rs_datos12!unimed_codigo = "", "PZA", rs_datos12!unimed_codigo)
            rs_aux7("bien_cantidad").Value = 1      'IIf(Txt_campo16 = "", 1, Txt_campo16)
            rs_aux7("bien_total_venta").Value = IIf(rs_datos12!bien_precio_venta_final = "", 0, rs_datos12!bien_precio_venta_final)
            rs_aux7("bien_precio_compra").Value = 0
            rs_aux7("bien_total_compra").Value = 0
            
            rs_aux7("fosa_dimension_frente").Value = 7      'IIf(dtc_codigo5.Text = "", 7, dtc_codigo5.Text)
            rs_aux7("fosa_dimension_fondo").Value = 0  'Txt_campo15.Text
            'mw_solicitud.Ado_detalle1.Recordset("archivo_foto").Value = Trim(dtc_codigo1.Text) + "-A.JPG"
            'mw_solicitud.Ado_detalle1.Recordset("archivo_plano").Value = Trim(dtc_codigo1.Text) + "-B.JPG"
            rs_aux7("fecha_registro").Value = Date
            'mw_solicitud.Ado_detalle1.Recordset("hora_registro").Value = Date
            rs_aux7("usr_codigo").Value = glusuario
            rs_aux7.UpdateBatch adAffectAll
            
         rs_datos12.MoveNext
         Wend
    End If
    
End Sub

Private Sub BtnAddDetalle7_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        VAR_DET = "24300"
        Call NuevoDetalle
        VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
    
         If OptFilGral1.Value = True Then
            Call OptFilGral1_Click        'Pendientes
         Else
            Call OptFilGral2_Click        'TODOS
         End If
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
    Else
        MsgBox "NO puede adicionar un NUEVO SERVICIO TECNICO, porque previamente debe registrar... " + FraDet2.Caption, vbExclamation
    End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   If Ado_detalle2.Recordset.RecordCount > 0 Then
       If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle2.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede BORRAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle3_Click()
   If Ado_detalle3.Recordset.RecordCount > 0 Then
       If Ado_detalle3.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle3.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle3.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnAnlDetalle5_Click()
   If Ado_detalle5.Recordset.RecordCount > 0 Then
       If Ado_detalle5.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle5.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle5.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle6_Click()
   If Ado_detalle6.Recordset.RecordCount > 0 Then
       If Ado_detalle6.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle6.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle6.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle7_Click()
   If Ado_detalle7.Recordset.RecordCount > 0 Then
       If Ado_detalle7.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle7.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle7.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Then         'Or glusuario = "LNAVA"
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo UpdateErr
'  If Ado_datos.Recordset.RecordCount > 0 Then
'   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
'        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
'        Exit Sub
'   End If
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
   If rs_datos!estado_codigo = "REG" Then       'And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "update ao_solicitud_bienes set ao_solicitud_bienes.almacen_tipo = ac_bienes.almacen_tipo from ac_bienes where ac_bienes.bien_codigo = ao_solicitud_bienes.bien_codigo"
        
        db.Execute "UPDATE ac_bienes SET unimed_codigo_empaque = unimed_codigo where (unimed_codigo_empaque Is Null) "
        db.Execute "UPDATE ao_solicitud_bienes SET unimed_codigo_empaque = unimed_codigo where (unimed_codigo_empaque Is Null) "
        
        db.Execute "UPDATE ac_bienes SET almacen_tipo ='Q' WHERE (par_codigo ='43340' AND almacen_tipo IS NULL) "
        db.Execute "UPDATE ao_solicitud_bienes SET almacen_tipo ='Q' WHERE (par_codigo ='43340' AND almacen_tipo IS NULL) "
        VAR_UNI = Ado_datos.Recordset!unidad_codigo
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        VAR_SUBP = Ado_datos.Recordset!subproceso_codigo
        Select Case VAR_SUBP        'dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "TEC-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                'If rs_aux1.RecordCount > 0 Then
                '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                '    var_cod = 0
                '    Exit Sub
                'Else
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & VAR_UNI & "' ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = glGestion
                    rs_aux1!unidad_codigo = VAR_UNI
                    rs_aux1!solicitud_codigo = VAR_SOL
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!trafico_codigo = var_cod
                   ' rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
                'End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
            
            'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Case "TEC-02", "TEC-03", "TEC-04", "TEC-05"     '10. SERVICIO MANTENIMIENTO Y REPARACIONES
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               If VAR_SUBP = "TEC-02" Then
                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               Else
                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, SUM(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               End If
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
                    db.Execute "delete ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "' "
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = glGestion
                    rs_aux1!unidad_codigo = VAR_UNI
                    rs_aux1!solicitud_codigo = VAR_SOL
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!depto_codigo = Left(Ado_datos.Recordset!edif_codigo, 1)
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
                    'If rs_aux3.RecordCount > 0 Then
                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    'Else
                        'db.Execute "delete ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "' "
                        VAR_AUX = rs_aux3.RecordCount + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion         'glGestion
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = VAR_AUX
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = IIf(IsNull(rs_aux5!bien_precio_eur), 0, rs_aux5!bien_precio_eur)   'EUR
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = IIf(IsNull(rs_aux5!bien_total_eur), 0, rs_aux5!bien_total_eur)    'EUR
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!observaciones = rs_aux5!observacion
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!bien_codigo_padre = rs_aux5!bien_codigo_padre
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                        rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique los datos de Bienes..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        Case "COM-03"    '3. SERVICIO INSTALACION
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = glGestion
                    rs_aux1!unidad_codigo = VAR_UNI
                    rs_aux1!solicitud_codigo = VAR_SOL
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   '
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
                    'If rs_aux3.RecordCount > 0 Then
                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    'Else
                        VAR_AUX = rs_aux3.RecordCount + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = VAR_AUX
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = IIf(IsNull(rs_aux5!bien_precio_eur), 0, rs_aux5!bien_precio_eur)   'EUR
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = IIf(IsNull(rs_aux5!bien_total_eur), 0, rs_aux5!bien_total_eur)    'EUR
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!observaciones = rs_aux5!observacion
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                     rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique la Venta de Productos..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        Case "COM-04"    '4. SERVICIO AJUSTE
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & VAR_UNI & "' and solicitud_codigo =" & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = glGestion
                    rs_aux1!unidad_codigo = VAR_UNI
                    rs_aux1!solicitud_codigo = VAR_SOL
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   ", db, adOpenStatic
                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & glGestion & "'   ", db, adOpenKeyset, adLockOptimistic
                    'If rs_aux3.RecordCount > 0 Then
                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    'Else
                        VAR_AUX = rs_aux3.RecordCount + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = VAR_AUX
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = IIf(IsNull(rs_aux5!bien_precio_eur), 0, rs_aux5!bien_precio_eur)   'EUR
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = IIf(IsNull(rs_aux5!bien_total_eur), 0, rs_aux5!bien_total_eur)    'EUR
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                     rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique la Venta de Productos..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        End Select
        If rs_datos!unidad_codigo = "DNMAN" Then
            db.Execute "update ao_solicitud set estado_cotiza = 'APR' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
            'rs_datos!estado_cotiza = "APR"
        End If
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            Txt_campo1.Caption = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        db.Execute "update ao_solicitud set doc_numero = " & Txt_campo1.Caption & " where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        'rs_datos!doc_numero = txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        VAR_ARCH = "TEC_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
        db.Execute "update ao_solicitud set archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        db.Execute "update ao_solicitud set archivo_respaldo_cargado = 'N' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        db.Execute "update ao_solicitud set estado_codigo = 'APR' where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        db.Execute "update ao_solicitud set fecha_aprueba = '" & Date & "'  where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        db.Execute "update ao_solicitud set usr_codigo_aprueba = '" & glusuario & "'  where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & VAR_SOL & "   "
        'rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
        'rs_datos!archivo_respaldo_cargado = "N"
        'rs_datos!estado_codigo = "APR"
        'rs_datos!fecha_aprueba = Date
        'rs_datos!usr_codigo_aprueba = glusuario
        'rs_datos.UpdateBatch adAffectAll
        db.Execute "update ao_ventas_detalle set ao_ventas_detalle.almacen_tipo = ac_bienes.almacen_tipo from ac_bienes where ac_bienes.bien_codigo = ao_ventas_detalle.bien_codigo "
        
        
    VAR_COD2 = Ado_datos.Recordset!solicitud_codigo
    OptFilGral2_Click
    
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
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description


End Sub

Private Sub BtnAux2_Click()
    glPersNew = "NEWF"
    txt_nombre.Visible = False
    gw_p_gc_beneficiario_aux.Show vbModal
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Call OptFilGral2_Click
        buscados = 1
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
'        If OptFilGral1.Value = True Then
'            MsgBox "Esta Buscando los Registros... " + OptFilGral1.Caption, vbInformation, "Atención!"
'        Else
'            MsgBox "Esta Buscando... " + OptFilGral2.Caption + " los Registros.", vbInformation, "Atención!"
'        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
        '        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
      OptFilGral1.Visible = True
      OptFilGral2.Visible = True
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.Cancel
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        mbDataChanged = False
        Fra_datos.Visible = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        dg_det2.Visible = True
        dg_det3.Visible = True
        dg_det5.Visible = True
        dg_det6.Visible = True
        dg_det7.Visible = True
        
        FrmABMDet2.Enabled = True
        FrmABMDet5.Enabled = True
        FrmABMDet3.Visible = True
        FrmABMDet6.Enabled = True
        FrmABMDet7.Enabled = True
        'txt_codigo.Enabled = True
        If rs_datos!solicitud_codigo <> "" Then
        VAR_SOLA = rs_datos!solicitud_codigo
        End If
        
     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
        rs_datos.Find "solicitud_codigo = " & VAR_SOLA & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
        rs_datos.MoveLast
     End If
     
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub btnEliminar_Click()
    If glusuario = "CCRUZ" Then         'Or glusuario = "LNAVA"
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    'If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If ExisteReg(Ado_datos.Recordset!unidad_codigo, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ANL"
          rs_datos!fecha_registro = Date
          rs_datos!usr_codigo = glusuario
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
        rs_datos!estado_codigo = "ERR"
        rs_datos!fecha_registro = Date
        rs_datos!usr_codigo = glusuario
        rs_datos.UpdateBatch adAffectAll
       'MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "Select max(solicitud_codigo) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        'If rs_aux1.RecordCount > 0 Then
        If Not rs_aux1.EOF Then
            var_cod = IIf(IsNull(rs_aux1!Codigo), 1, rs_aux1!Codigo + 1)
        Else
            var_cod = 1
        End If
'        rs_aux11
        'CORRELATIVO CITE TEC- REPARACION
        If VAR_UNI = "DNREP" Or VAR_UNI = "DREPS" Or VAR_UNI = "DREPB" Or VAR_UNI = "DREPC" Then
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            rs_aux11.Open "Select correl_negocia as Codigo from  gc_unidad_ejecutora where unidad_codigo = '" & VAR_UNI & "' ", db, adOpenKeyset, adLockOptimistic
            If Not rs_aux11.EOF Then
                VAR_CITES = IIf(IsNull(rs_aux11!Codigo), 1, rs_aux11!Codigo + 1)
            Else
                VAR_CITES = 1
            End If
            'Actualiza correaltivo Cite Cotiza ...
            db.Execute "Update gc_unidad_ejecutora Set correl_negocia = " & VAR_CITES & " Where unidad_codigo = '" & VAR_UNI & "'   "
        Else
            'CORRELATIVO CITE-TEC MANTENIMIENTO Y OTROS
            Set rs_aux10 = New ADODB.Recordset
            If rs_aux10.State = 1 Then rs_aux10.Close
            SQL_FOR = "Select max(doc_numero2) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
            rs_aux10.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If Not rs_aux10.EOF Then
                VAR_CITES = IIf(IsNull(rs_aux10!Codigo), 1, rs_aux10!Codigo + 1)
            Else
                VAR_CITES = 1
            End If
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        Txt_campo3.Text = VAR_CITES
        ' Guardar con INSERT
        'ges_gestion, unidad_codigo, solicitud_codigo, solicitud_fecha_solicitud, solicitud_fecha_recepción, solicitud_tipo, edif_codigo, beneficiario_codigo,
'                    beneficiario_codigo_resp, beneficiario_codigo_resp2, unidad_codigo_sol, solicitud_justificacion, solicitud_observaciones, proceso_codigo, subproceso_codigo,
'                      etapa_codigo, etapa_codigo2, clasif_codigo, doc_codigo, doc_codigo2, doc_numero, doc_numero2, poa_codigo, ges_gestion_ant, unidad_codigo_ant,
'                      solicitud_codigo_ant, correl_detalle, correl_edificacion, correl_calculo, correl_persona, correl_cotiza, correl_bitacora, archivo_respaldo, archivo_respaldo_cargado,
'                      estado_codigo, estado_etapa2, estado_cotiza, fecha_registro, hora_registro, usr_codigo, usr_codigo_aprueba, fecha_aprueba, hora_aprueba, fecha_registro2,
'                      usr_codigo2 , observacion_proy, mes_codigo
                      
        'db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " &
        '"VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"

        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
        rs_datos!doc_numero2 = IIf(Txt_campo3.Text = "", "0", Txt_campo3.Text)
     End If
     If VAR_SW = "MOD" Then
        VAR_UNI = rs_datos!unidad_codigo
        var_cod = rs_datos!solicitud_codigo
     End If
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     'rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        VAR_BENEF = IIf(txt_ci.Text = "", "0", txt_ci.Text)
        'rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        VAR_BENEF = dtc_codigo4.Text
        'rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     rs_datos!beneficiario_codigo = VAR_BENEF
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     If var_cod < 10 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
     End If
     If var_cod > 9 And var_cod < 100 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
     End If
     If var_cod > 99 And var_cod < 1000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
     End If
     If var_cod > 999 And var_cod < 10000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
     End If
     If var_cod > 9999 And var_cod < 100000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
     End If
     If var_cod > 99999 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
     End If

     'rs_datos!poa_codigo = IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
     'Select Case dtc_codigo2.Text
     Select Case dtc_codigo1.Text
'        Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL - Case "1"    'SOLO COMPRAS BB y SS
'            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
'            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
'            rs_datos!etapa_codigo = "COM-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
'            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
'            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
'            rs_datos!poa_codigo = "3.1.1"
'        Case "CMX-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'            rs_datos!solicitud_tipo = "3"
'            rs_datos!proceso_codigo = "CMX"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
'            rs_datos!subproceso_codigo = "CMX-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
'            rs_datos!etapa_codigo = "CMX-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
'            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
'            rs_datos!doc_codigo = "R-XXX"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
'            rs_datos!poa_codigo = "4.1.1"
'        Case "COM-02"    '3. COMPRA-VENTA BB Y SS - COMERCIAL -         'VENTAS NUEVAS
        Case "DVTA", "DCOMS", "DCOMB", "DCOMC"    '3. VENTAS NUEVAS
            rs_datos!solicitud_tipo = "3"
            rs_datos!proceso_codigo = "COM"
            rs_datos!subproceso_codigo = "COM-02"
            rs_datos!etapa_codigo = "COM-02-01"
            rs_datos!clasif_codigo = "COM"
            rs_datos!doc_codigo = "R-234"
            rs_datos!poa_codigo = "3.1.1"
'        Case "COM-03"    'VENTA DE SERVICIOS INSTALACIONES
        Case "DNINS", "DINSS", "DINSB", "DINSC"    '4. SERVICIOS INSTALACIONES
            rs_datos!solicitud_tipo = "4"
            rs_datos!proceso_codigo = "COM"
            rs_datos!subproceso_codigo = "COM-03"
            rs_datos!etapa_codigo = "COM-03-01"
            rs_datos!clasif_codigo = "TEC"
            rs_datos!doc_codigo = "R-362"
            rs_datos!poa_codigo = "3.2.2"
'        Case "COM-04" '5       'VENTA DE SERVICIOS AJUSTE
        Case "DNAJS", "DAJSS", "DAJSB", "DAJSC"    '5. SERVICIOS AJUSTE
            rs_datos!solicitud_tipo = "5"
            rs_datos!proceso_codigo = "COM"
            rs_datos!subproceso_codigo = "COM-04"
            rs_datos!etapa_codigo = "COM-04-01"
            rs_datos!clasif_codigo = "TEC"
            rs_datos!doc_codigo = "R-362"
            rs_datos!poa_codigo = "3.2.6"
'        Case "TEC-01"    '6. SERVICIO MANTENIMIENTO GRATUITO
'            rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
'            rs_datos!subproceso_codigo = "TEC-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
'            rs_datos!etapa_codigo = "TEC-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
'            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
'            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
'            rs_datos!poa_codigo = "3.2.3"           'IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
'        Case "TEC-02"    '10. SERVICIO MANTENIMIENTO PREVENTIVO
        Case "DNMAN", "DMANS", "DMANB", "DMANC"    '10. SERVICIO MANTENIMIENTO INTEGRAL
            rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = "TEC"       'Left(dtc_codigo2.Text, 3)
            rs_datos!subproceso_codigo = "TEC-02"       'IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = "TEC-02-01"        'Trim(dtc_codigo2.Text) + "-01"  '
            rs_datos!clasif_codigo = "TEC"           'Left(dtc_codigo2.Text, 3)  '
            rs_datos!doc_codigo = "R-355"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            rs_datos!poa_codigo = "3.2.3"           'IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
            'COD.ADM. o CODIGO DE CONTRATO
            rs_datos!unidad_codigo_ant = Trim(Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)) + "-" + Trim(CStr(glGestion))
            'End If
'        Case "TEC-03" '10 REPARACION    If VAR_UNI = "DNIREP" Then
        Case "DNREP", "DREPS", "DREPB", "DREPC"    '7. SERVICIO DE REPARACIONES
            rs_datos!solicitud_tipo = "7"
            rs_datos!proceso_codigo = "TEC"
            rs_datos!subproceso_codigo = "TEC-03"
            rs_datos!etapa_codigo = "TEC-03-01"
            rs_datos!clasif_codigo = "TEC"
            rs_datos!doc_codigo = "R-362"
            rs_datos!poa_codigo = "3.2.4"
'        Case "TEC-04" '10 EMERGENCIAS   If VAR_UNI = "DNEME" Then
        Case "DNEME", "DEMES", "DEMEB", "DEMEC"    '8. SERVICIO DE EMERGENCIAS
            rs_datos!solicitud_tipo = "8"
            rs_datos!proceso_codigo = "TEC"
            rs_datos!subproceso_codigo = "TEC-04"
            rs_datos!etapa_codigo = "TEC-04-01"
            rs_datos!clasif_codigo = "TEC"
            rs_datos!doc_codigo = "R-362"
            rs_datos!poa_codigo = "3.2.1"
'        Case "TEC-05"    '5. SERVICIO MODERNIZACION -If VAR_UNI = "DNMOD" Then
        Case "DNMOD", "DMODS", "DMODB", "DMODC"    '9. SERVICIO DE MODERNIZACIONES
            rs_datos!solicitud_tipo = "9"
            rs_datos!proceso_codigo = "TEC"
            rs_datos!subproceso_codigo = "TEC-05"
            rs_datos!etapa_codigo = "TES-05-01"
            rs_datos!clasif_codigo = "TEC"
            rs_datos!doc_codigo = "R-362"
            rs_datos!poa_codigo = "3.2.7"
        Case Else   '10. SERVICIO MANTENIMIENTO PREVENTIVO
            'If VAR_UNI = "DNMAN" Then
            rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = "TEC"             'Left(dtc_codigo2.Text, 3)     ' "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "TEC-02"       'IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = "TEC-02-01"         'Trim(dtc_codigo2.Text) + "-01"  'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"              'Left(dtc_codigo2.Text, 3)      'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-355"
            rs_datos!poa_codigo = "3.2.3"
     End Select
     rs_datos!TipoContratoCodigo = IIf(dtc_codigo10.Text = "", "0", dtc_codigo10.Text)
     rs_datos!PlazoDias = IIf(TxtPlazo.Text = "", "15", TxtPlazo.Text)
     rs_datos!solicitud_observaciones = IIf(txt_obs.Text = "", "", txt_obs.Text)
     rs_datos!observaciones2 = IIf(txt_obs2.Text = "", "", txt_obs2.Text)
     rs_datos!observaciones3 = IIf(txt_obs3.Text = "", "", txt_obs3.Text)
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     rs_datos!observacion_proy = dtc_desc3.Text
     rs_datos!trans_codigo = IIf(dtc_codigo10.Text = "", "0", dtc_codigo10.Text)
     rs_datos!ges_gestion_ant = glGestion       'glGestion
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     rs_datos!codigo_empresa = dtc_codigo8.Text
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     VAR_SOLA = rs_datos!solicitud_codigo
'     If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Call OptFilGral1_Click
'     Else
'        Call OptFilGral2_Click
'     End If
'     rs_datos.MoveLast

     If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
     VAR_SW = ""
        rs_datos.Find "solicitud_codigo = " & VAR_SOLA & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     VAR_SW = ""
        rs_datos.MoveLast
     End If
    
     mbDataChanged = False
      
     Fra_datos.Visible = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
    dg_datos.Enabled = True
    dg_det2.Visible = True
    dg_det3.Visible = True
    dg_det5.Visible = True
    dg_det6.Visible = True
    dg_det7.Visible = True
    
    FrmABMDet2.Enabled = True
    FrmABMDet5.Enabled = True
    FrmABMDet3.Visible = True
    FrmABMDet6.Enabled = True
    FrmABMDet7.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True
      
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If VAR_SW = "ADD" Then
'    If (Year(DTPfecha1.Value) = Year(Date)) Then
'      MsgBox "Debe registrar una fecha de la Gestión Actual ... ", vbCritical + vbExclamation, "Validación de datos"
'      VAR_VAL = "ERR"
'      Exit Sub
'    End If
'  End If
  
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo10.Text = "") Then
    MsgBox "Debe registrar Tipo de Tramite... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If (Ado_detalle1.Recordset.RecordCount > 0) And (Ado_datos.Recordset!estado_codigo <> "ANL") Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CR00.ReportFileName = App.Path & "\Reportes\comercial\ar_solicitud_cotizacion.rpt"
        CR00.ReportFileName = App.Path & "\Reportes\tecnico\tr_lista_solicitud_tecnico.rpt"
        CR00.WindowShowPrintSetupBtn = True
        CR00.WindowShowRefreshBtn = True
        'CR00.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
        CR00.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      CR00.Formulas(3) = "titulo = '" & var_titulo & "' "
      CR00.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "

        iResult = CR00.PrintReport
        If iResult <> 0 Then MsgBox CR00.LastErrorNumber & " : " & CR00.LastErrorString, vbCritical, "Error de impresión"
        CR00.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir2_Click()
    Select Case parametro
        Case "DNINS"            'INI GRABA INSTALACIONES
            'dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            'dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN", "DMANS", "DMANB", "DMANC"            'MANTENIMIENTO PREVENTIVO
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  'Dim iResult As Integer
                  'Dim co As New ADODB.Command
                  Set rs_datos5 = New ADODB.Recordset
                  If rs_datos5.State = 1 Then rs_datos5.Close
                  rs_datos5.Open "Select * from av_acumula_total_bienes_subgrupo where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " ", db, adOpenStatic
                  If rs_datos5.RecordCount > 0 Then
                        var_literal = Literal(CStr(rs_datos5!bien_total)) + " BOLIVIANOS"
                  Else
                        var_literal = "CERO 00/100 BOLIVIANOS"
                  End If
                  db.Execute "update ao_solicitud set literal = '" & var_literal & "' where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " "
                  '
                  'VAR_LITERAL = CantidadConLetra(CStr(Ado_detalle2.Recordset!bien_total_venta)) + " BOLIVIANOS"
                  'db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = av_acumula_ventas_detalle.venta_det_cantidad from ac_bienes, av_acumula_ventas_detalle Where ac_bienes.grupo_codigo = av_acumula_ventas_detalle.grupo_codigo And ac_bienes.subgrupo_codigo = av_acumula_ventas_detalle.subgrupo_codigo And ac_bienes.bien_codigo = av_acumula_ventas_detalle.bien_codigo"
                  'FraImprimeMantenimiento.Visible = True
                  '-----------------------------------------------------------------PAGINA 1
                  sino = MsgBox("La Cotización, Imprimirá con datos del Cliente (Nombre, Cargo, Institución, etc.) ? ", vbYesNo + vbQuestion, "Atención")
                    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
                    '-----------------------------------------------------------------PAGINA 1
                    If sino = vbYes Then
                      If Ado_datos.Recordset!codigo_empresa = 2 Then
                        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1_cliente_CGE.rpt"
                      Else
                        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1_cliente.rpt"
                      End If
                    Else
                      If Ado_datos.Recordset!codigo_empresa = 2 Then
                        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1_CGE.rpt"
                      Else
                        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1.rpt"
                      End If
                    End If
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta1.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA SERVICIO DE MANTENIMIENTO INTEGRAL"
                  CR02.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR02.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR02.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  'CR02.WindowState = crptMaximized
                  '-----------------------------------------------------------------PAGINA 2
                  If Ado_datos.Recordset!codigo_empresa = 2 Then
                    CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta2_CGE.rpt"
                  Else
                    CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta2.rpt"
                  End If
                  CR04.WindowShowPrintSetupBtn = True
                  CR04.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA TÉCNICA"
                  CR04.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR04.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR04.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "

                  CR04.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR04.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR04.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR04.PrintReport
                  If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
                  'CR04.WindowState = crptMaximized
                  '-----------------------------------------------------------------PAGINA 3
                  If Ado_datos.Recordset!codigo_empresa = 2 Then
                    CR05.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta3_CGE.rpt"
                  Else
                    CR05.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta3.rpt"
                  End If
                  CR05.WindowShowPrintSetupBtn = True
                  CR05.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA TÉCNICA"
                  CR05.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR05.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR05.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "

                  CR05.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR05.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR05.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR05.PrintReport
                  If iResult <> 0 Then MsgBox CR05.LastErrorNumber & " : " & CR05.LastErrorString, vbCritical, "Error de impresión"
                  'CR05.WindowState = crptMaximized
                  '-----------------------------------------------------------------PAGINA 4
                  If Ado_datos.Recordset!codigo_empresa = 2 Then
                    CR06.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta4_CGE.rpt"
                  Else
                    CR06.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta4.rpt"
                  End If

                  CR06.WindowShowPrintSetupBtn = True
                  CR06.WindowShowRefreshBtn = True

                  VAR_TIT = "GERENCIA TECNICA"
                  VAR_SUBT = "PROPUESTA ECONÓMICA"
                  CR06.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
                  CR06.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
                  CR06.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
                  CR06.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR06.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR06.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR06.PrintReport
                  If iResult <> 0 Then MsgBox CR06.LastErrorNumber & " : " & CR06.LastErrorString, vbCritical, "Error de impresión"
                  'CR06.WindowState = crptMaximized

                  '-----------------------------------------------------------------
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
        Case "DNREP", "DREPS", "DREPB", "DREPC"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
          If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                'INI LITERAL
                Set rs_datos5 = New ADODB.Recordset
                If rs_datos5.State = 1 Then rs_datos5.Close
                rs_datos5.Open "Select SUM(bien_total) AS bien_total from av_acumula_total_bienes_subgrupo where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " AND (subgrupo_codigo = '39000' OR subgrupo_codigo ='24000' OR subgrupo_codigo ='34000') ", db, adOpenStatic
                If rs_datos5.RecordCount > 0 Then
                      var_literal = Literal(CStr(rs_datos5!bien_total)) + " BOLIVIANOS"
                Else
                      var_literal = "CERO 00/100 BOLIVIANOS"
                End If
                db.Execute "update ao_solicitud set literal = '" & var_literal & "' where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " "
                'FIN LITERAL
                'Option5.Value = True
                FraImprimeRepara.Visible = True
                fraOpciones.Visible = False
                FraNavega.Enabled = False
                FraDet3.Visible = False
                FraDet6.Visible = False
'                sino = MsgBox("La Cotización, Imprimirá con datos del Cliente (Nombre, Cargo, Institución, etc.) ? ", vbYesNo + vbQuestion, "Atención")
'                  'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'                  '-----------------------------------------------------------------PAGINA 1 - OPCION 1 y 2
'                  If sino = vbYes Then
'                    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_cliente.rpt"
'                  Else
'                    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag1.rpt"
'                  End If
'                  CR02.WindowShowPrintSetupBtn = True
'                  CR02.WindowShowRefreshBtn = True
'
'                  CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
'                  CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
'
'                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
'                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
'                  iResult = CR02.PrintReport
'                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
'                  CR02.WindowState = crptMaximized
'                  '-----------------------------------------------------------------PAGINA 2
'                  CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag2.rpt"
'                  CR04.WindowShowPrintSetupBtn = True
'                  CR04.WindowShowRefreshBtn = True
'
'                  CR04.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
'                  CR04.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
'
'                  CR04.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'                  CR04.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
'                  CR04.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
'                  iResult = CR04.PrintReport
'                  If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
'                  CR04.WindowState = crptMaximized
'                  'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If

        Case "DNEME"            'EMERGENCIAS
            'dtc_codigo2.Text = "TEC-04" '10
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  
                  'Dim co As New ADODB.Command
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_solicitud_cotizacion.rpt"
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True
                  'MsgBox rs.RecordCount
                    CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                    CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                  'Call CREAVISTAF11          'JQA JUN-2008
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
        Case "DNMOD"            'MODERNIZACION
            'dtc_codigo2.Text = "TEC-05" '10
        Case Else
            'dtc_codigo2.Text = "TEC-01"   '3
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  
                  'Dim co As New ADODB.Command
                  'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_solicitud_cotizacion.rpt"
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True
                  'MsgBox rs.RecordCount
                    CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                    CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                  'Call CREAVISTAF11          'JQA JUN-2008
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
    End Select
  
End Sub

Private Sub BtnImprimir3_Click()
 
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle3.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
'        VAR_TIT = "MODULO TECNICO"
'                  VAR_SUBT = "PROPUESTA DE SERVICIO MANTENIMIENTO PREVENTIVO"
'                  CR02.Formulas(0) = "Titulo = '" & VAR_TIT & "' "
'                  CR02.Formulas(1) = "Subtitulo = '" & VAR_SUBT & "' "
'                  CR02.Formulas(2) = "Subtitulo2 = '" & lbl_titulo.Caption & "' "
                  
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet3.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet3.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir4_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle5.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet5.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet5.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir5_Click()

    Set rs_aux12 = New ADODB.Recordset
    If rs_aux12.State = 1 Then rs_aux12.Close
    'AND (unidad_destino IS NOT NULL)
    rs_aux12.Open "Select * from ao_ventas_cabecera WHERE ((unidad_codigo = 'DVTA') OR (unidad_codigo = 'DCOMS') OR (unidad_codigo = 'DCOMB') OR (unidad_codigo = 'DCOMC')) AND (estado_codigo = 'APR') AND (unidad_destino IS NULL) ", db, adOpenStatic
    If rs_aux12.RecordCount > 0 Then
        rs_aux12.MoveFirst
        While Not rs_aux12.EOF
            Set rs_aux13 = New ADODB.Recordset
            If rs_aux13.State = 1 Then rs_aux13.Close
            rs_aux13.Open "Select * from ao_ventas_cabecera WHERE ((unidad_codigo = 'DNMAN') OR (unidad_codigo = 'DMANS') OR (unidad_codigo = 'DMANB') OR (unidad_codigo = 'DMANC')) AND (estado_codigo = 'APR') AND (edif_codigo = '" & rs_aux12!edif_codigo & "') ", db, adOpenStatic
            If rs_aux13.RecordCount > 0 Then
                'rs_aux12!unidad_destino = rs_aux13!unidad_CODIGO
                db.Execute "UPDATE ao_ventas_cabecera SET unidad_destino = '" & rs_aux13!unidad_codigo & "' WHERE (venta_codigo = " & rs_aux12!venta_codigo & " ) "
                db.Execute "UPDATE ao_ventas_alcance SET estado_mantenimiento = 'APR' WHERE (venta_codigo = " & rs_aux12!venta_codigo & " AND solicitud_tipo ='6') "
            End If
            rs_aux12.MoveNext
        Wend
    End If
'        VAR_GESTION2 = rs_aux1!ges_gestion
'        VAR_UNIDAD2 = rs_aux1!unidad_codigo
'        VAR_SOL2 = rs_aux1!solicitud_codigo
'        VER_EDIF2 = rs_aux1!edif_codigo
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR07.ReportFileName = App.Path & "\reportes\comercial\ar_lista_actas_entrega_definitiva_ok.rpt"
        CR07.WindowShowPrintSetupBtn = True
        CR07.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR07.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
'        CR07.StoredProcParam(0) = VAR_GESTION2
'        CR07.StoredProcParam(1) = VAR_UNIDAD2
'        CR07.StoredProcParam(2) = VAR_SOL2
'        CR07.StoredProcParam(3) = VER_EDIF2
'        CR07.StoredProcParam(4) = "1"           'Me.Ado_datos.Recordset!cotiza_codigo
        iResult = CR07.PrintReport
        If iResult <> 0 Then MsgBox CR07.LastErrorNumber & " : " & CR07.LastErrorString, vbCritical, "Error de impresión"
'    'Else
'    '    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
'    'End If
        CR07.WindowState = crptMaximized
'    Else
'        MsgBox "No Existe el Equipo registrado en Ventas Nuevas... "
'    End If

End Sub

Private Sub BtnImprimir7_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle7.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet7.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet7.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnModDetalle2_Click()
  If Ado_detalle2.Recordset.RecordCount > 0 Then
      If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
        marca1 = Ado_detalle2.Recordset.Bookmark
        swnuevo = 2
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FraDet2.Enabled = False
        FrmABMDet2.Enabled = False
        FraDet3.Enabled = False
        FrmABMDet3.Enabled = False
        Fra_datos.Enabled = False
    
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'                Call ABRIR_TABLA_DET
'                mw_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'                mw_solicitud_edificacion.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
'                'mw_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'                'mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'                mw_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
'                mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'
'                mw_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
'                mw_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
'                mw_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
'                mw_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
'                mw_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
'                mw_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
'                mw_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
'                mw_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
'                mw_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
'                mw_solicitud_edificacion.Txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
'                mw_solicitud_edificacion.Txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
'                mw_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
'
'                mw_solicitud_edificacion.Show vbModal
            Case "COM-03"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
                Call ABRIR_TABLA_DET
                tw_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                'tw_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle2.Recordset("bitacora_codigo")
                'tw_solicitud_bienes.Txt_estado.Caption = "REG"
                tw_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                tw_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                tw_solicitud_bienes.dtc_desc1.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux1.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux2.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux3.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo2.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo3.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo4.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                
                tw_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                tw_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                tw_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                tw_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                tw_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                tw_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                tw_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                tw_solicitud_bienes.dtc_desc2.BoundText = tw_solicitud_bienes.dtc_codigo2.BoundText
                tw_solicitud_bienes.lbl_det.Caption = "43340"
                tw_solicitud_bienes.Show vbModal
                
            Case "TEC-05"    '5. SERVICIO MODERNIZACION
            Case "TEC-01"    '6. SERVICIO DE MANTENIMIENTO GRATUITO
                Call ABRIR_TABLA_DET
                tw_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                'tw_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle2.Recordset("bitacora_codigo")
                'tw_solicitud_bienes.Txt_estado.Caption = "REG"
                tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                tw_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                tw_solicitud_bienes.dtc_desc1.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux1.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux2.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.dtc_aux3.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo2.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo3.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                tw_solicitud_bienes.Txt_campo4.BoundText = tw_solicitud_bienes.dtc_codigo1.BoundText
                
                tw_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                tw_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                tw_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                tw_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                tw_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                tw_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                tw_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                tw_solicitud_bienes.dtc_desc2.BoundText = tw_solicitud_bienes.dtc_codigo2.BoundText
                tw_solicitud_bienes.lbl_det.Caption = "43340"
                tw_solicitud_bienes.Show vbModal
            Case "TEC-02"    '10. VENTA DE SERVICIO DE MANTENIMIENTO PREVENTIVO
                'Call ABRIR_TABLA_DET
                tw_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                
                tw_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                tw_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                tw_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                tw_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                tw_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                tw_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                tw_solicitud_bienes.Txt_campo19.Text = Me.Ado_detalle2.Recordset("bien_cantidad_por_empaque")
                
                tw_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                tw_solicitud_bienes.Txt_campo15.Text = "10" 'Me.Ado_detalle2.Recordset("fosa_dimension_frente")
                
                tw_solicitud_bienes.lbl_det.Caption = "43340"
                tw_solicitud_bienes.Show vbModal
            Case "TEC-03"    '7. VENTA DE SERVICIOS REPARACION
                tw_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                tw_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                    
                tw_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                tw_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                tw_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                tw_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                tw_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                tw_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                
                tw_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
    '            tw_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
    '            tw_solicitud_bienes.dtc_desc2.BoundText = tw_solicitud_bienes.dtc_codigo2.BoundText
                tw_solicitud_bienes.lbl_det.Caption = "43340"
                tw_solicitud_bienes.Show vbModal
            
        End Select
        swnuevo = 0
        fraOpciones.Enabled = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Enabled = True
        FraDet3.Enabled = True
        FrmABMDet3.Enabled = True
    '    Fra_datos.Enabled = True
        Call ABRIR_TABLA_DET
        Ado_detalle2.Recordset.Move marca1 - 1
      Else
        MsgBox "No se puede MODIFICAR, porque ya está APROBADO o ANULADO, Verifique por favor!! ", vbExclamation
      End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub ModifDetalle()
  
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False

            If VAR_DET = "30000" Then
                'marca1 = Ado_detalle3.Recordset.Bookmark
                tw_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes3.Txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
                
                tw_solicitud_bienes3.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
                tw_solicitud_bienes3.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes3.Txt_campo8.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!marca_codigo), "S/M", Me.Ado_detalle3.Recordset!marca_codigo)
                tw_solicitud_bienes3.Txt_campo9.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!modelo_codigo), "S/M", Me.Ado_detalle3.Recordset!modelo_codigo)
                
                tw_solicitud_bienes3.Txt_campo16.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_cantidad), "1", Me.Ado_detalle3.Recordset!bien_cantidad)
                tw_solicitud_bienes3.Txt_campo10.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_precio_venta_base), "0", Me.Ado_detalle3.Recordset!bien_precio_venta_base)
                tw_solicitud_bienes3.Txt_campo11.Caption = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_total_venta), "0", Me.Ado_detalle3.Recordset!bien_total_venta)
    
                tw_solicitud_bienes3.Txt_campo14.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
                tw_solicitud_bienes3.Txt_campo15.Text = Me.Ado_detalle3.Recordset("fosa_dimension_frente")

                tw_solicitud_bienes3.lbl_det.Caption = VAR_DET
                GlExtension = Ado_detalle2.Recordset!bien_codigo            'Equipo Padre
                tw_solicitud_bienes3.Show vbModal
                'Ado_detalle3.Recordset.Move marca1 - 1
            End If
            If VAR_DET = "39800" Then
                tw_solicitud_bienes5.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes5.txt_codigo.Caption = Me.Ado_detalle5.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes5.Txt_campo1.Caption = Me.Ado_detalle5.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes5.Txt_campo5.Text = Me.Ado_detalle5.Recordset("bien_codigo")
                
                tw_solicitud_bienes5.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion), "-", Me.Ado_detalle5.Recordset!bien_descripcion)
                tw_solicitud_bienes5.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle5.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes5.Txt_campo8.Text = Me.Ado_detalle5.Recordset("marca_codigo")
                tw_solicitud_bienes5.Txt_campo9.Text = Me.Ado_detalle5.Recordset("modelo_codigo")
                
                tw_solicitud_bienes5.Txt_campo16.Text = Me.Ado_detalle5.Recordset("bien_cantidad")
                tw_solicitud_bienes5.Txt_campo10.Text = Me.Ado_detalle5.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes5.Txt_campo11.Text = Me.Ado_detalle5.Recordset("bien_total_venta")
                
                tw_solicitud_bienes5.TxtObservacion.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!observacion), "", Me.Ado_detalle5.Recordset!observacion)
                
                tw_solicitud_bienes5.Txt_campo14.Text = Me.Ado_detalle5.Recordset("unimed_codigo")
                tw_solicitud_bienes5.Txt_campo15.Text = Me.Ado_detalle5.Recordset("fosa_dimension_frente")
                tw_solicitud_bienes5.dtc_codigo2.BoundText = Me.Ado_detalle5.Recordset("unimed_codigo")
                tw_solicitud_bienes5.dtc_desc2.BoundText = tw_solicitud_bienes5.dtc_codigo2.BoundText
                'tw_solicitud_bienes5.dtc_desc2.BoundText = Me.Ado_detalle5.Recordset("unimed_codigo")
                GlExtension = Ado_detalle2.Recordset!bien_codigo
                tw_solicitud_bienes5.Show vbModal
            End If
            If VAR_DET = "34800" Then
                 tw_solicitud_bienes6.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes6.txt_codigo.Caption = Me.Ado_detalle6.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes6.Txt_campo1.Caption = Me.Ado_detalle6.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes6.Txt_campo5.Text = Me.Ado_detalle6.Recordset("bien_codigo")
                
                'tw_solicitud_bienes6.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
'                tw_solicitud_bienes6.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes6.Txt_campo8.Text = Me.Ado_detalle6.Recordset("marca_codigo")
                tw_solicitud_bienes6.Txt_campo9.Text = Me.Ado_detalle6.Recordset("modelo_codigo")
                
                tw_solicitud_bienes6.Txt_campo16.Text = Me.Ado_detalle6.Recordset("bien_cantidad")
                tw_solicitud_bienes6.Txt_campo10.Text = Me.Ado_detalle6.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes6.Txt_campo11.Caption = Me.Ado_detalle6.Recordset("bien_total_venta")
                
                tw_solicitud_bienes6.Txt_campo14.Text = Me.Ado_detalle6.Recordset("unimed_codigo")
                tw_solicitud_bienes6.Txt_campo15.Text = Me.Ado_detalle6.Recordset("fosa_dimension_frente")
                
                tw_solicitud_bienes6.lbl_det.Caption = VAR_DET
                tw_solicitud_bienes6.Show vbModal
            End If

            If VAR_DET = "24300" Then
                tw_solicitud_bienes7.txt_codigo.Caption = Me.Ado_detalle7.Recordset("solicitud_codigo")  'cod_cabecera
                tw_solicitud_bienes7.Txt_campo1.Caption = Me.Ado_detalle7.Recordset("unidad_codigo")  'Unidad
                tw_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                tw_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
                tw_solicitud_bienes7.Txt_campo5.Text = Me.Ado_detalle7.Recordset("bien_codigo")
                
                tw_solicitud_bienes7.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion), "-", Me.Ado_detalle7.Recordset!bien_descripcion)
                tw_solicitud_bienes7.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle7.Recordset!bien_descripcion_anterior)
                tw_solicitud_bienes7.Txt_campo8.Text = Me.Ado_detalle7.Recordset("marca_codigo")
                tw_solicitud_bienes7.Txt_campo9.Text = Me.Ado_detalle7.Recordset("modelo_codigo")
                
                tw_solicitud_bienes7.Txt_campo16.Text = Me.Ado_detalle7.Recordset("bien_cantidad")
                tw_solicitud_bienes7.Txt_campo10.Text = Me.Ado_detalle7.Recordset("bien_precio_venta_base")
                tw_solicitud_bienes7.Txt_campo11.Caption = Me.Ado_detalle7.Recordset("bien_total_venta")
                
                tw_solicitud_bienes7.Txt_campo14.Text = Me.Ado_detalle7.Recordset("unimed_codigo")
                tw_solicitud_bienes7.Txt_campo15.Text = Me.Ado_detalle7.Recordset("fosa_dimension_frente")
                
                'tw_solicitud_bienes7.TxtObservacion.Text = Me.Ado_detalle7.Recordset!observacion
                tw_solicitud_bienes7.TxtObservacion.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!observacion), "", Me.Ado_detalle7.Recordset!observacion)
                tw_solicitud_bienes7.lbl_det.Caption = VAR_DET
                GlExtension = Ado_detalle2.Recordset!bien_codigo            'Equipo Padre
                tw_solicitud_bienes7.Show vbModal
            End If
'    End Select
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
    Call ABRIR_TABLA_DET
'    Ado_detalle3.Recordset.Move marca1 - 1
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModDetalle3_Click()
   If Ado_detalle3.Recordset.RecordCount > 0 Then
       If Ado_detalle3.Recordset("estado_codigo") = "REG" Then
       marca1 = Ado_detalle3.Recordset.Bookmark
          VAR_DET = "30000"
          Call ModifDetalle
          Call ABRIR_TABLA_DET
        Ado_detalle3.Recordset.Move marca1 - 1
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle5_Click()
   If Ado_detalle5.Recordset.RecordCount > 0 Then
       If Ado_detalle5.Recordset("estado_codigo") = "REG" Then
         marca1 = Ado_detalle5.Recordset.Bookmark
          VAR_DET = "39800"
          swnuevo = "2"
          Call ModifDetalle
        Call ABRIR_TABLA_DET
        Ado_detalle5.Recordset.Move marca1 - 1
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle6_Click()
   If Ado_detalle6.Recordset.RecordCount > 0 Then
       If Ado_detalle6.Recordset("estado_codigo") = "REG" Then
       marca1 = Ado_detalle6.Recordset.Bookmark
          VAR_DET = "34800"
          Call ModifDetalle
           Call ABRIR_TABLA_DET
        Ado_detalle6.Recordset.Move marca1 - 1
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle7_Click()
   If Ado_detalle7.Recordset.RecordCount > 0 Then
       If Ado_detalle7.Recordset("estado_codigo") = "REG" Then
       marca1 = Ado_detalle7.Recordset.Bookmark
          VAR_DET = "24300"
          Call ModifDetalle
          Call ABRIR_TABLA_DET
        Ado_detalle7.Recordset.Move marca1 - 1
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModificar_Click()
    If glusuario = "CCRUZ" Then     'Or glusuario = "LNAVA"
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        'marca1 = Ado_datos.Recordset.Bookmark
        Fra_datos.Visible = True
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        FrmABMDet3.Visible = False
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc4.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
        'Call OptFilGral1_Click
        'Ado_datos.Recordset.Move marca1 - 1
        Select Case parametro
            Case "DVTA"             'INI COMERCIAL
                dtc_codigo2.Text = "COM-01"   '3
            Case "COMEX"            'INI COMEX
                dtc_codigo2.Text = "CMX-01"   '3
            Case "DNINS"            'INI GRABA INSTALACIONES
                dtc_codigo2.Text = "COM-03" '4
            Case "DNAJS"            'AJUSTE
                dtc_codigo2.Text = "COM-04" '5
            Case "DNMAN", "DMANB", "DMANS", "DMANC"            'MANTENIMIENTO PREVENTIVO
                dtc_codigo2.Text = "TEC-02" '10
            Case "DNREP", "DREPB", "DREPS", "DREPC"         'MANTENIMIENTO CORRECTIVO / REPARACIONES
                dtc_codigo2.Text = "TEC-03" '10
            Case "DNEME"           'EMERGENCIAS
                dtc_codigo2.Text = "TEC-04" '10
            Case "DNMOD"            'MODERNIZACION
                dtc_codigo2.Text = "TEC-05" '10
            Case Else
                dtc_codigo2.Text = "TEC-01"   '3
        End Select
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub btnPanelImprimir_Click()
    Dim iResult As Integer
    If Option1.Value = True Then
        ' 1. Con datos del Cliente (Nombre, Cargo, Institución, etc.)   -    - OPCION 1
        '-----------------------------------------------------------------     PAGINA 1
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_cliente.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True

        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
        Call CotizaRep_Pag1Op1_2
    ElseIf Option2.Value = True Then
        '2. Sólo con Nombre de Edificio (Sin Datos del Cliente)         - OPCION 2
        '-----------------------------------------------------------------PAGINA 1
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag1.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True

        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
        Call CotizaRep_Pag1Op1_2
    ElseIf Option3.Value = True Then
        ' 3. Con datos del Cliente (Nombre, Cargo, Institución, etc.)   -    - OPCION 3
        '-----------------------------------------------------------------     PAGINA 1
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_cliente_SM.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True

        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
        Call CotizaRep_Pag1Op1_2SM
    ElseIf Option4.Value = True Then
        '4. Sólo con Nombre de Edificio (Sin Datos del Cliente)         - OPCION 4
        '-----------------------------------------------------------------PAGINA 1
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag1_SM.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True

        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

        CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
        Call CotizaRep_Pag1Op1_2SM
    Else
        MsgBox "Debe Elegir una de las opciones", vbExclamation, "Atencion"
    End If
End Sub

Private Sub btnPanelSalir_Click()
    FraImprimeRepara.Visible = False
    fraOpciones.Visible = True
    FraNavega.Enabled = True
    FraDet3.Visible = True
    FraDet6.Visible = True
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
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
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
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
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
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

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
    'Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   
   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub
  
Private Sub pnivel11(codigo2 As String)
    Select Case codigo2
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"            'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
'    Dim strConsultaF As String
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
End Sub

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
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub
 
Private Sub dtc_desc3_LostFocus()
    dtc_codigo4.Text = dtc_aux3.Text
    'Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text
    Select Case parametro
        Case "DNMAN", "DMANS", "DMANB", "DMANC"
            Txt_descripcion.Text = "Propuesta de Servicio de MANTENIMIENTO INTEGRAL. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
        Case "DNREP", "DREPS", "DREPB", "DREPC"
            Txt_descripcion.Text = "Servicio de REPARACIONES. Edificio: " + dtc_desc3.Text
        Case "DNMOD", "DMODS", "DMODB", "DMODC"
            Txt_descripcion.Text = "Propuesta de MODERNIZACION de equipos. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
        Case "DNINS", "DINSS", "DINSB", "DINSC"
            Txt_descripcion.Text = "Servicio de INSTALACION de equipos. Edificio: " + dtc_desc3.Text
        Case "DNEME", "DEMES", "DEMEB", "DEMEC"
            Txt_descripcion.Text = "Atención de EMERGENCIAS. Edificio: " + dtc_desc3.Text + ". Cod.ADM.: " + Mid(dtc_codigo3.Text, 7, Len(dtc_codigo3.Text) - 6)
            Set rs_aux9 = New ADODB.Recordset
            If rs_aux9.State = 1 Then rs_aux9.Close
            rs_aux9.Open "Select * from tv_zona_piloto_edif_resp ", db, adOpenStatic
            If rs_aux9.RecordCount > 0 Then
                dtc_codigo11.Text = rs_aux9!beneficiario_codigo
                dtc_desc11.BoundText = dtc_codigo11.BoundText
                'dtc_desc11.Text = rs_aux9!beneficiario_denominacion
            Else
                dtc_codigo11.Text = "4245046"
                dtc_desc11.BoundText = dtc_codigo11.BoundText
                'dtc_desc11.Text = "ORAQUENI QUITO JAVIER"
            End If
        Case Else
    End Select
    dtc_desc4.BoundText = dtc_codigo4.BoundText
  
    'Call pnivel1(dtc_codigo1.BoundText)
    'dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
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
    buscados = 0
    swnuevo = 0
    VAR_SW = ""
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux8.RecordCount > 0 Then
        usuario2 = rs_aux8!beneficiario_codigo
        VAR_DA = rs_aux8!da_codigo
        VAR_DPTOC = rs_aux8!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    VAR_UORIGEN = Aux
    If Aux = "DNMAN" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DMANB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DMANS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNMAN"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DMANC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNMAN"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 10
         VAR_PLAZO = "Validez             días"
     End If
     If Aux = "DNREP" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DREPB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DREPS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNREP"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DREPC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNREP"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 7
         VAR_PLAZO = "Plazo:               días"
     End If
     If Aux = "DNINS" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DINSB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DINSS"
                'VAR_DPTOC = "7"
            Case "1.3", "1.2"    'La Paz - Tecnico
                Aux = "DNINS"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DINSC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNINS"
                'VAR_DPTOC = "0"
         End Select
         VAR_TIPO = 4
         VAR_PLAZO = "Plazo:               días"
     End If
    If Aux = "DNEME" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DMANB"
                'VAR_DPTOC = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DMANS"
                'VAR_DPTOC = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNEME"
                'VAR_DPTOC = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DMANC"
                'VAR_DPTOC = "1"
            Case "0"    ' TODO
                Aux = "DNEME"
                'VAR_DPTOC = "2"
         End Select
         VAR_TIPO = 8
         VAR_PLAZO = "Plazo:               dias"
     End If
    LblPlazo.Caption = VAR_PLAZO
    parametro = Aux
    db.Execute "UPDATE ao_solicitud SET ao_solicitud.observacion_proy = gc_edificaciones.edif_descripcion from ao_solicitud inner join gc_edificaciones on ao_solicitud.edif_codigo = gc_edificaciones.edif_codigo WHERE (ao_solicitud.unidad_codigo = '" & parametro & "')"
    'db.Execute "UPDATE ao_solicitud SET ao_solicitud.observacion_proy = gc_edificaciones.edif_descripcion from ao_solicitud inner join gc_edificaciones on ao_solicitud.edif_codigo = gc_edificaciones.edif_codigo where ao_solicitud.edif_codigo <> '0' and ao_solicitud.observacion_proy is null"
    'parametro = "estado_codigo" + " = " + "'REG'"
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    
    'gc_proceso_nivel2
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If parametro = "DNINS" Or parametro = "DNAJS" Then
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'COM' order by subproceso_descripcion", db, adOpenStatic
    Else
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'TEC' order by subproceso_descripcion", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones WHERE (estado_codigo = 'APR') OR (edif_codigo_corto='6') order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'EMPRESA
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_empresas order by codigo_empresa", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    'gc_ContratoTipo  -
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_ContratoTipo WHERE solicitud_tipo = " & VAR_TIPO & " order by TipoContratoCodigo", db, adOpenStatic
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

Private Sub ABRIR_TABLA_DET()
    'BITACORA
'    Set rs_det1 = New ADODB.Recordset
'    If rs_det1.State = 1 Then rs_det1.Close
'    'rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle1.Recordset = rs_det1
'    If rs_det1.RecordCount > 0 Then
'        dg_det1.Visible = True
'        Set dg_det1.DataSource = Ado_detalle1.Recordset
'    Else
'        dg_det1.Visible = False
'        'Set Ado_detalle1.Recordset = rsNada
'        Set dg_det1.DataSource = rsNada
'    End If
    
    'EQUIPOS par_codigo = '43340'
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    'rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and (par_codigo = '43340' ) ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and estado_codigo = 'APR'
    rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " and (par_codigo = '43340' ) ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and estado_codigo = 'APR'
    Set Ado_detalle2.Recordset = rs_det2
    If rs_det2.RecordCount > 0 Then
        dg_det2.Visible = True
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det2.Visible = False
        'Set Ado_detalle2.Recordset = rsNada
        Set dg_det2.DataSource = rsNada
    End If
    
    'INSUMOS y materiales par_codigo = '43340'
    Set rs_det3 = New Recordset
    If rs_det3.State = 1 Then rs_det3.Close
    'rs_det3.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    rs_det3.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    Set Ado_detalle3.Recordset = rs_det3.DataSource
    If rs_det3.RecordCount > 0 Then
        dg_det3.Visible = True
        Set dg_det3.DataSource = Ado_detalle3.Recordset
    Else
        dg_det3.Visible = False
        Set dg_det3.DataSource = rsNada
    End If

    'REPUESTOS par_codigo = '39800'
    Set rs_det5 = New Recordset
    If rs_det5.State = 1 Then rs_det5.Close
    'rs_det5.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  AND (almacen_tipo = 'R')  ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    rs_det5.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  AND (almacen_tipo = 'R')  ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    Set Ado_detalle5.Recordset = rs_det5.DataSource
    If rs_det5.RecordCount > 0 Then
        dg_det5.Visible = True
        Set dg_det5.DataSource = Ado_detalle5.Recordset
    Else
        dg_det5.Visible = False
        Set dg_det5.DataSource = rsNada
    End If

    'HERRAMIENTAS par_codigo = '43700' - par_codigo = '34800'
    Set rs_det6 = New Recordset
    If rs_det6.State = 1 Then rs_det6.Close
    'rs_det6.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText     'and estado_codigo = 'APR'
    rs_det6.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText     'and estado_codigo = 'APR'
    Set Ado_detalle6.Recordset = rs_det6.DataSource
    If rs_det6.RecordCount > 0 Then
        dg_det6.Visible = True
        Set dg_det6.DataSource = Ado_detalle6.Recordset
    Else
        dg_det6.Visible = False
        Set dg_det6.DataSource = rsNada
    End If
    
    Set rs_det4 = New Recordset
    If rs_det4.State = 1 Then rs_det4.Close
    'rs_det4.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det4.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle4.Recordset = rs_det4.DataSource
    Set dg_det4.DataSource = Ado_detalle4.Recordset
    
    'REPUESTOS par_codigo = '24000'
    Set rs_det7 = New Recordset
    If rs_det7.State = 1 Then rs_det7.Close
    'rs_det7.Open "select * from av_solicitud_bienes7 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and par_codigo = '24300'   ", db, adOpenKeyset, adLockOptimistic, adCmdText      '
    rs_det7.Open "select * from av_solicitud_bienes7 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & VAR_SOL & " and par_codigo = '24300'   ", db, adOpenKeyset, adLockOptimistic, adCmdText      '
    Set Ado_detalle7.Recordset = rs_det7.DataSource
    If rs_det7.RecordCount > 0 Then
        dg_det7.Visible = True
        Set dg_det7.DataSource = Ado_detalle7.Recordset
    Else
        dg_det7.Visible = False
        Set dg_det7.DataSource = rsNada
    End If
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
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
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    If VAR_SW <> "ADD" Then
        'Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
'        If VAR_SOL = 0 Then
        'If VAR_SW <> "" Then
        If Not (Ado_datos.Recordset.EOF) Then   'And Not (Ado_datos.Recordset.BOF)
            VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        End If
        Call ABRIR_TABLA_DET
        'VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        'Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
'    FraDet1.Caption = "BITÁCORA " + dtc_desc2.Text
    'FraDet1.Caption = "BITÁCORA DE " + lbl_titulo
'    txt_aux9.Text = dtc_desc9.Text
    If Not (Ado_datos.Recordset.EOF) Then
        If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet2.Visible = False
            FrmABMDet3.Visible = False
            BtnAprobar.Visible = False
            If glusuario = "ADMIN" Or glusuario = "ADMINSTC" Or glusuario = "ADMINCBB" Or glusuario = "ADMINCHQ" Or glusuario = "CSALINAS" Then
                BtnDesAprobar.Visible = True
            Else
                BtnDesAprobar.Visible = False
            End If
        Else
            If Ado_datos.Recordset!estado_codigo = "REG" Then
                BtnAprobar.Visible = True
                BtnDesAprobar.Visible = False
            End If
            FrmABMDet2.Visible = True
            FrmABMDet3.Visible = True
        End If
    End If
  Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        Set dg_det3.DataSource = rsNada
        Set dg_det5.DataSource = rsNada
        Set dg_det6.DataSource = rsNada
        Set dg_det7.DataSource = rsNada
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
  End If
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
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

Private Sub BtnAñadir_Click()
    If glusuario = "CCRUZ" Then     'Or glusuario = "LNAVA"
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo AddErr
    VAR_SW = "ADD"
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Visible = True
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    dg_det2.Visible = False
    dg_det3.Visible = False
    dg_det5.Visible = False
    dg_det6.Visible = False
    dg_det7.Visible = False
    FrmABMDet2.Enabled = False
    FrmABMDet5.Enabled = False
    FrmABMDet3.Visible = False
    FrmABMDet6.Enabled = False
    FrmABMDet7.Enabled = False
    'txt_codigo.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
    Ado_datos.Recordset.AddNew
    dtc_desc11.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_desc2.Locked = True
    Select Case parametro
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN", "DMANB", "DMANS", "DMANC"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP", "DREPB", "DREPS", "DREPC"         'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"           'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    dtc_desc2.BoundText = dtc_codigo2.BoundText
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
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String, Codigo As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_ventas_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & Codigo & " and estado_codigo = 'APR'   "
'    <> 'ANL'
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

'Private Function ExisteReg(Unidad As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE edif_codigo = '" & Unidad & "'"
'    rs.Open GlSqlAux, db, adOpenStatic
'    ExisteReg = rs!Cuantos > 0
'End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "2"
            queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') "
        Case "7"
            queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR   Left(edif_codigo, 1) = '8' OR   Left(edif_codigo, 1) = '9' OR   Left(edif_codigo, 1) = '1' )) "
        Case "3"
            queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
        Case "1"
            queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
        Case Else
            queryinicial = "Select * from ao_solicitud where (estado_codigo = 'REG' AND Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "')) "
    End Select
            'queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
            'queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo='" & VAR_UORIGEN & "' AND left(edif_codigo,1) = '" & VAR_DPTO & "')) "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    If rs_datos.RecordCount > 0 Then
    rs_datos.MoveFirst
    End If
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "2"
            queryinicial = "Select * from ao_solicitud WHERE (unidad_codigo = '" & parametro & "') "
        Case "7"
            queryinicial = "Select * from ao_solicitud where ((unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR   Left(edif_codigo, 1) = '8' OR   Left(edif_codigo, 1) = '9' OR   Left(edif_codigo, 1) = '1' )) "
            'queryinicial = "Select * from ao_solicitud where ((Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR   Left(edif_codigo, 1) = '8' OR   Left(edif_codigo, 1) = '9' OR   Left(edif_codigo, 1) = '1' )) "
        Case "3"
            queryinicial = "Select * from ao_solicitud where ((unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
            'queryinicial = "Select * from ao_solicitud where ((Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '4' )) "
        Case "1"
            queryinicial = "Select * from ao_solicitud where ((unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "') AND (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
            'queryinicial = "Select * from ao_solicitud where ((Left(edif_codigo, 1) = '" & VAR_DPTOC & "' OR  Left(edif_codigo, 1) = '5' OR  Left(edif_codigo, 1) = '6' )) "
        Case Else
            queryinicial = "Select * from ao_solicitud where (Left(edif_codigo, 1) = '" & VAR_DPTOC & "' AND (unidad_codigo = '" & parametro & "' OR unidad_codigo = '" & VAR_UORIGEN & "'))"
    End Select
    'queryinicial = "Select * from ao_solicitud where unidad_codigo = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "unidad_codigo, solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
     If rs_datos.RecordCount > 0 Then
    rs_datos.MoveFirst
    End If
End Sub

'Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txt_obs_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub CotizaRep_Pag1Op1_2()
    Dim iResult As Integer
    '-----------------------------------------------------------------PAGINA 2 - OPCION 1 y 2
    CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag2.rpt"
    CR04.WindowShowPrintSetupBtn = True
    CR04.WindowShowRefreshBtn = True

    CR04.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
    CR04.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

    CR04.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR04.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR04.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    iResult = CR04.PrintReport
    If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
    CR04.WindowState = crptMaximized
    FraImprimeRepara.Visible = False
    fraOpciones.Visible = True
    FraNavega.Enabled = True
    FraDet3.Visible = True
    FraDet6.Visible = True
End Sub

Private Sub CotizaRep_Pag1Op1_2SM()
    Dim iResult As Integer
    '-----------------------------------------------------------------PAGINA 2 - OPCION 1 y 2
    CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion_pag2_SM.rpt"
    CR04.WindowShowPrintSetupBtn = True
    CR04.WindowShowRefreshBtn = True

    CR04.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
    CR04.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "

    CR04.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR04.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR04.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    iResult = CR04.PrintReport
    If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
    CR04.WindowState = crptMaximized
    FraImprimeRepara.Visible = False
    fraOpciones.Visible = True
    FraNavega.Enabled = True
    FraDet3.Visible = True
    FraDet6.Visible = True
End Sub

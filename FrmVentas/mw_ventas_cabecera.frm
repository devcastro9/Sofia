VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form mw_ventas_cabecera 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Administrativos - Ventas - Proceso de Ventas"
   ClientHeight    =   10740
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   16845
   Icon            =   "mw_ventas_cabecera.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7.28916e6
   ScaleMode       =   0  'User
   ScaleWidth      =   2.23075e9
   WindowState     =   2  'Maximized
   Begin VB.Frame FraZona 
      BackColor       =   &H00404040&
      Caption         =   "Elija una Zona Piloto..."
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
      Height          =   2535
      Left            =   8760
      TabIndex        =   214
      Top             =   4800
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   4080
         MaskColor       =   &H00000000&
         Picture         =   "mw_ventas_cabecera.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   216
         ToolTipText     =   "Cancelar"
         Top             =   1560
         Width           =   1365
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   2040
         Picture         =   "mw_ventas_cabecera.frx":12EE
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   1560
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "mw_ventas_cabecera.frx":1ADC
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   218
         Top             =   660
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux7 
         Bindings        =   "mw_ventas_cabecera.frx":1AF5
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   219
         Top             =   300
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "mes_par_impar"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "mw_ventas_cabecera.frx":1B0E
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1080
         TabIndex        =   220
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "0"
      End
      Begin VB.Label Label29 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   217
         Top             =   1425
         Width           =   2025
      End
   End
   Begin VB.Frame frm_benef 
      BackColor       =   &H00404040&
      Caption         =   "Registra Datos del Cliente"
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
      Height          =   2535
      Left            =   8760
      TabIndex        =   207
      Top             =   4800
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton BtnGrabarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   2040
         Picture         =   "mw_ventas_cabecera.frx":1B27
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelarBen 
         BackColor       =   &H00C0FFFF&
         Height          =   635
         Left            =   4080
         MaskColor       =   &H00000000&
         Picture         =   "mw_ventas_cabecera.frx":2315
         Style           =   1  'Graphical
         TabIndex        =   210
         ToolTipText     =   "Cancelar"
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox TxtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos16"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   209
         Text            =   "0"
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox TxtCelular 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Ado_datos16"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   208
         Text            =   "0"
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono Celular:"
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
         Left            =   720
         TabIndex        =   213
         Top             =   960
         Width           =   2025
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico:"
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
         Left            =   720
         TabIndex        =   212
         Top             =   465
         Width           =   2025
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   138
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10920
         Picture         =   "mw_ventas_cabecera.frx":2C01
         ScaleHeight     =   615
         ScaleWidth      =   1575
         TabIndex        =   195
         ToolTipText     =   "Registra Adenda o Modificación al Contrato"
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18000
         Picture         =   "mw_ventas_cabecera.frx":34B8
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   139
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnVer3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8040
         Picture         =   "mw_ventas_cabecera.frx":3C7A
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   161
         ToolTipText     =   "Cambiar Contrato a Provisional o Viceversa"
         Top             =   20
         Visible         =   0   'False
         Width           =   1430
      End
      Begin VB.PictureBox BtnAprobar1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         Picture         =   "mw_ventas_cabecera.frx":47C4
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   160
         ToolTipText     =   "Verifica el Contrato"
         Top             =   20
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6600
         Picture         =   "mw_ventas_cabecera.frx":4FFC
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   159
         ToolTipText     =   "Cerrar Tramite y Archivarlo (cuando ya no tiene nada pendiente)"
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
         Picture         =   "mw_ventas_cabecera.frx":5AB6
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   140
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   0
         Width           =   1400
      End
      Begin VB.CommandButton BtnDesAprobar 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   18120
         Picture         =   "mw_ventas_cabecera.frx":6383
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         DownPicture     =   "mw_ventas_cabecera.frx":658D
         Height          =   705
         Left            =   19800
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   145
         ToolTipText     =   "Nuevo"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   -15
         Picture         =   "mw_ventas_cabecera.frx":6D4C
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   144
         ToolTipText     =   "Modifica datos del Contrato"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1440
         Picture         =   "mw_ventas_cabecera.frx":7661
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   143
         ToolTipText     =   "Anula Contrato elegido"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9600
         Picture         =   "mw_ventas_cabecera.frx":7DAD
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   142
         ToolTipText     =   "Aprueba el Contrato Elegido (ya NO podrá ser modificado)"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "mw_ventas_cabecera.frx":85E0
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   141
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
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
         Left            =   14280
         TabIndex        =   147
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   1875
      TabIndex        =   115
      Top             =   6960
      Width           =   1935
      Begin VB.PictureBox BtnImprimir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":8D95
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   181
         ToolTipText     =   "Imprime el Plan de Cuotas a cobrar"
         Top             =   2400
         Width           =   1400
      End
      Begin VB.PictureBox BtnAprobar2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":9662
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   180
         ToolTipText     =   "Aprueba Cuota y Envia a Facturación u Orden de Cobro"
         Top             =   1800
         Width           =   1320
      End
      Begin VB.PictureBox BtnAnlDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":9E95
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   179
         ToolTipText     =   "Anula Cuota elegida"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox BtnModDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":A5E1
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   178
         ToolTipText     =   "Modifica datos de la Cuota elegida"
         Top             =   600
         Width           =   1430
      End
      Begin VB.PictureBox BtnAddDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":AEF6
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   177
         ToolTipText     =   "Registra Nueva Cuota"
         Top             =   0
         Width           =   1430
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000018&
      FillColor       =   &H00FFFFFF&
      Height          =   1260
      Left            =   120
      Negotiate       =   -1  'True
      ScaleHeight     =   5
      ScaleMode       =   4  'Character
      ScaleWidth      =   15.625
      TabIndex        =   113
      Top             =   5655
      Width           =   1935
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modifica Equipo"
         Height          =   720
         Left            =   240
         Picture         =   "mw_ventas_cabecera.frx":B6B5
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Modifica datos del Equipo"
         Top             =   240
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4770
      Left            =   6480
      TabIndex        =   18
      Top             =   765
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8414
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "REGISTRO DE VENTAS"
      TabPicture(0)   =   "mw_ventas_cabecera.frx":BAF7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DETALLE BIENES (Equipos)"
      TabPicture(1)   =   "mw_ventas_cabecera.frx":BB13
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "REGISTRO PLAN DE CUOTAS"
      TabPicture(2)   =   "mw_ventas_cabecera.frx":BB2F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrmCobros"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "ALCANCE DEL CONTRATO"
      TabPicture(3)   =   "mw_ventas_cabecera.frx":BB4B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmAlcance"
      Tab(3).Control(1)=   "FraGrabarCancelar1"
      Tab(3).Control(2)=   "FrmABMDet1"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox FrmABMDet1 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -74880
         Negotiate       =   -1  'True
         ScaleHeight     =   3.313
         ScaleMode       =   4  'Character
         ScaleWidth      =   98.625
         TabIndex        =   174
         Top             =   3720
         Width           =   11895
         Begin VB.PictureBox BtnModDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5760
            Picture         =   "mw_ventas_cabecera.frx":BB67
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   187
            ToolTipText     =   "Modifica datos del Alcance del Contrato"
            Top             =   120
            Width           =   1430
         End
         Begin VB.PictureBox BtnAddDetalle1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "mw_ventas_cabecera.frx":C47C
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   182
            ToolTipText     =   "Genera nuevos items del Alcance del Conrato"
            Top             =   120
            Width           =   1430
         End
      End
      Begin VB.PictureBox FraGrabarCancelar1 
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   -74880
         ScaleHeight     =   705
         ScaleWidth      =   11835
         TabIndex        =   190
         Top             =   3840
         Visible         =   0   'False
         Width           =   11895
         Begin VB.PictureBox BtnFlecha01 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   590
            Left            =   7200
            Picture         =   "mw_ventas_cabecera.frx":CC3B
            ScaleHeight     =   585
            ScaleWidth      =   825
            TabIndex        =   193
            ToolTipText     =   "Genera nuevos items del Alcance del Conrato"
            Top             =   0
            Width           =   825
         End
         Begin VB.PictureBox BtnGrabar1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   10200
            Picture         =   "mw_ventas_cabecera.frx":D361
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   191
            Top             =   50
            Width           =   1280
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "... Cuando termine -->"
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
            Left            =   8040
            TabIndex        =   194
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label LblAyuda01 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Para Registrar: ""Fecha.Inicio y Fecha.Fin"", modifique las fechas ..."
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
            Left            =   960
            TabIndex        =   192
            Top             =   240
            Width           =   6075
         End
      End
      Begin VB.Frame FrmAlcance 
         BackColor       =   &H00C0C0C0&
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   175
         Top             =   360
         Width           =   11895
         Begin MSDataGridLib.DataGrid DtgAlcance 
            Bindings        =   "mw_ventas_cabecera.frx":DB4F
            Height          =   2985
            Left            =   120
            Negotiate       =   -1  'True
            TabIndex        =   176
            Top             =   240
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   5265
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "venta_codigo"
               Caption         =   "Cod_Venta"
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
               DataField       =   "solicitud_tipo"
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
            BeginProperty Column02 
               DataField       =   "solicitud_tipo_descripcion"
               Caption         =   "Descripcion del Alcance"
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
               DataField       =   "unidad_codigo_tec"
               Caption         =   "Unidad.Ejecutora"
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
               DataField       =   "fecha_inicio_alcance"
               Caption         =   "Fecha.Inicio"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "fecha_fin_alcance"
               Caption         =   "Fecha.Fin"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "venta_tiempo_dias"
               Caption         =   "Dias Calendario"
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
                  ColumnWidth     =   585.071
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   4500.284
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  Locked          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  DividerStyle    =   1
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  DividerStyle    =   1
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column07 
                  Alignment       =   2
                  ColumnWidth     =   615.118
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame FrmCobros 
         BackColor       =   &H00C0C0C0&
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
         Height          =   4350
         Left            =   -74960
         TabIndex        =   53
         Top             =   380
         Width           =   12015
         Begin VB.CommandButton CmdEmail 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11340
            Picture         =   "mw_ventas_cabecera.frx":DB68
            Style           =   1  'Graphical
            TabIndex        =   206
            Top             =   2175
            Width           =   375
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   11020
            TabIndex        =   74
            Top             =   2190
            Width           =   270
         End
         Begin MSDataListLib.DataCombo dtc_email2A 
            Bindings        =   "mw_ventas_cabecera.frx":E56A
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   8565
            TabIndex        =   196
            Top             =   2175
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_email"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.ComboBox Txt_liquida 
            DataField       =   "es_liquidacion"
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
            ForeColor       =   &H00000040&
            Height          =   315
            ItemData        =   "mw_ventas_cabecera.frx":E583
            Left            =   11040
            List            =   "mw_ventas_cabecera.frx":E58D
            TabIndex        =   156
            Text            =   "NO"
            Top             =   1575
            Width           =   675
         End
         Begin VB.ComboBox cmd_fac 
            Height          =   315
            ItemData        =   "mw_ventas_cabecera.frx":E599
            Left            =   240
            List            =   "mw_ventas_cabecera.frx":E5A3
            TabIndex        =   154
            Text            =   "FACTURA"
            Top             =   1040
            Width           =   1995
         End
         Begin MSComCtl2.DTPicker DTPFechaProg 
            DataField       =   "cobranza_fecha_prog"
            DataSource      =   "Ado_datos16"
            Height          =   285
            Left            =   10005
            TabIndex        =   123
            Top             =   195
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   119537665
            CurrentDate     =   44713
            MinDate         =   32874
         End
         Begin VB.CheckBox Chk_plazo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Es requisito para el Plazo de entrega ?"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   7215
            TabIndex        =   121
            Top             =   3240
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txt_plazo 
            CausesValidation=   0   'False
            DataField       =   "cobranza_concepto_plazo"
            DataSource      =   "Ado_datos16"
            Height          =   465
            Left            =   2280
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   119
            Top             =   3120
            Width           =   9420
         End
         Begin VB.PictureBox Frame7 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   680
            Left            =   0
            ScaleHeight     =   615
            ScaleWidth      =   12000
            TabIndex        =   108
            Top             =   3675
            Width           =   12060
            Begin VB.PictureBox CmdCancelaCobro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   6240
               Picture         =   "mw_ventas_cabecera.frx":E5C0
               ScaleHeight     =   615
               ScaleWidth      =   1455
               TabIndex        =   189
               Top             =   10
               Width           =   1455
            End
            Begin VB.PictureBox CmdGrabaCobro 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   4440
               Picture         =   "mw_ventas_cabecera.frx":EEAC
               ScaleHeight     =   615
               ScaleWidth      =   1275
               TabIndex        =   188
               Top             =   10
               Width           =   1280
            End
         End
         Begin VB.TextBox TxtCobrador 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            CausesValidation=   0   'False
            DataField       =   "nombre_cobrador"
            DataSource      =   "Ado_datos16"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   60
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   1575
            Width           =   4650
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8520
            TabIndex        =   75
            Top             =   1590
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            Bindings        =   "mw_ventas_cabecera.frx":F682
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   6960
            TabIndex        =   97
            Top             =   1575
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox TxtDsctoTot 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cobranza_programada_dol"
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   10005
            TabIndex        =   55
            Text            =   "0"
            Top             =   1040
            Width           =   1680
         End
         Begin VB.TextBox TxtDscto 
            Alignment       =   2  'Center
            DataField       =   "cobranza_total_bs"
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
            Height          =   285
            Left            =   8280
            TabIndex        =   12
            Text            =   "0"
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtMontoDol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "cobranza_total_dol"
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
            Enabled         =   0   'False
            Height          =   285
            Left            =   10320
            TabIndex        =   54
            Text            =   "0"
            Top             =   1320
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtMonto 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   7815
            TabIndex        =   11
            Text            =   "0"
            Top             =   1040
            Width           =   1575
         End
         Begin VB.TextBox TxtObs 
            CausesValidation=   0   'False
            DataField       =   "cobranza_observaciones"
            DataSource      =   "Ado_datos16"
            Height          =   465
            Left            =   2280
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   2595
            Width           =   9420
         End
         Begin MSDataListLib.DataCombo dtc_codigo2A 
            Bindings        =   "mw_ventas_cabecera.frx":F69C
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   6960
            TabIndex        =   94
            Top             =   2175
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   16777215
            ListField       =   "beneficiario_nit"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc2A 
            Bindings        =   "mw_ventas_cabecera.frx":F6B5
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   2280
            TabIndex        =   95
            Top             =   2175
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc4A 
            Bindings        =   "mw_ventas_cabecera.frx":F6CE
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   2280
            TabIndex        =   96
            Top             =   1575
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin MSComCtl2.DTPicker DTPFechaCobro 
            DataField       =   "cobranza_fecha_cobro"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   2925
            TabIndex        =   112
            Top             =   1035
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
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
            CalendarBackColor=   16777215
            Format          =   119537667
            CurrentDate     =   44600
            MaxDate         =   109939
            MinDate         =   36526
         End
         Begin MSDataListLib.DataCombo dtc_benef2A 
            Bindings        =   "mw_ventas_cabecera.frx":F6E8
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos16"
            Height          =   315
            Left            =   9840
            TabIndex        =   199
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NIT"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7080
            TabIndex        =   198
            Top             =   1965
            Width           =   270
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8640
            TabIndex        =   197
            Top             =   1965
            Width           =   435
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Estimada a Cobrar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            TabIndex        =   162
            Top             =   750
            Width           =   2430
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Es Liquidación?"
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
            Height          =   195
            Left            =   9600
            TabIndex        =   157
            Top             =   1605
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "Documento a Emitir:"
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
            TabIndex        =   155
            Top             =   750
            Width           =   1785
         End
         Begin VB.Label TxtNroVentaC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "venta_codigo"
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1440
            TabIndex        =   122
            Top             =   195
            Width           =   1365
         End
         Begin VB.Label lbl_plazo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto p/ Factura  o p/Orden de Cobro:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   240
            TabIndex        =   120
            Top             =   3080
            Width           =   1920
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(USD) $us."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   10290
            TabIndex        =   118
            Top             =   750
            Width           =   960
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFFC0&
            X1              =   -240
            X2              =   11160
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Label Txt_cod_cobro 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataField       =   "cobranza_prog_codigo"
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5115
            TabIndex        =   117
            Top             =   195
            Width           =   1005
         End
         Begin VB.Label Lbl_nombre_fac 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Factura u Orden de Cobro a nombres de:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   225
            TabIndex        =   116
            Top             =   2025
            Width           =   1950
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Cuota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   3960
            TabIndex        =   111
            Top             =   195
            Width           =   1050
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9315
            TabIndex        =   110
            Top             =   1275
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   9165
            TabIndex        =   109
            Top             =   915
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lbl_fechas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Programada de la Cuota:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   6840
            TabIndex        =   62
            Top             =   195
            Width           =   3075
         End
         Begin VB.Label Lbl_Cobrador 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Encargao de Cobrar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   225
            TabIndex        =   61
            Top             =   1605
            Width           =   1875
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            BackStyle       =   0  'Transparent
            Caption         =   "(BOB) Bs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   8040
            TabIndex        =   60
            Top             =   750
            Width           =   885
         End
         Begin VB.Label lbl_monto 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "IMPORTE DE LA CUOTA -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   5175
            TabIndex        =   59
            Top             =   1065
            Width           =   2550
         End
         Begin VB.Label lbl_obs 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto de la Cuota:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   225
            TabIndex        =   58
            Top             =   2660
            Width           =   2040
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   225
            TabIndex        =   57
            Top             =   195
            Width           =   1050
         End
      End
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
         Height          =   4350
         Left            =   -74960
         TabIndex        =   36
         Top             =   380
         Width           =   12015
         Begin VB.Frame Fra_Monto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad --------------- Precio Unitario Usd ----------- Descuento ------------  Total Dolares (Usd)"
            Enabled         =   0   'False
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
            Height          =   780
            Left            =   285
            TabIndex        =   130
            Top             =   2720
            Width           =   8295
            Begin VB.TextBox TxtPrecioU 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               DataField       =   "venta_precio_unitario_dol"
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
               Height          =   285
               Left            =   2040
               TabIndex        =   134
               Text            =   "0"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_precio_total_dol"
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
               Height          =   285
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   133
               Text            =   "0"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox TxtDescuento 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_descuento_dol"
               DataSource      =   "ado_datos14"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4200
               TabIndex        =   132
               Text            =   "0"
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox TxtCantidad 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_det_cantidad"
               DataSource      =   "ado_datos14"
               Height          =   285
               Left            =   120
               TabIndex        =   131
               Text            =   "0"
               Top             =   360
               Width           =   1215
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
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   3840
               TabIndex        =   137
               Top             =   360
               Width           =   225
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
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   1560
               TabIndex        =   136
               Top             =   405
               Width           =   240
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
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   5880
               TabIndex        =   135
               Top             =   360
               Width           =   285
            End
         End
         Begin VB.TextBox Txt_modelo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4560
            TabIndex        =   106
            Text            =   "0"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Txt_modelo3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   6345
            TabIndex        =   103
            Text            =   "0"
            Top             =   2340
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txt_modelo2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
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
            Left            =   4440
            TabIndex        =   102
            Text            =   "0"
            Top             =   2340
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox Txt_modelo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1920
            TabIndex        =   100
            Text            =   "0"
            Top             =   2340
            Width           =   2175
         End
         Begin VB.OptionButton OpMod3 
            BackColor       =   &H00404040&
            Caption         =   "3"
            Height          =   285
            Left            =   7800
            TabIndex        =   8
            Top             =   2340
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton OpMod2 
            BackColor       =   &H00404040&
            Caption         =   "2"
            Height          =   285
            Left            =   5760
            TabIndex        =   7
            Top             =   2340
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.OptionButton OpMod1 
            BackColor       =   &H00404040&
            Caption         =   "1"
            Height          =   285
            Left            =   4095
            TabIndex        =   6
            Top             =   2340
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.PictureBox FraGrabarDet 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   780
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   12000
            TabIndex        =   99
            Top             =   0
            Width           =   12060
            Begin VB.CommandButton BtnAnlDetalle 
               BackColor       =   &H80000018&
               Caption         =   "Anular-->"
               Height          =   640
               Left            =   1560
               Picture         =   "mw_ventas_cabecera.frx":F701
               Style           =   1  'Graphical
               TabIndex        =   186
               ToolTipText     =   "Anula la Cobranza Identificada"
               Top             =   0
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.PictureBox CmdCancelaDet 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   6240
               Picture         =   "mw_ventas_cabecera.frx":103CB
               ScaleHeight     =   615
               ScaleWidth      =   1455
               TabIndex        =   185
               Top             =   60
               Width           =   1455
            End
            Begin VB.PictureBox CmdGrabaDet 
               Appearance      =   0  'Flat
               BackColor       =   &H80000006&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   4440
               Picture         =   "mw_ventas_cabecera.frx":10CB7
               ScaleHeight     =   615
               ScaleWidth      =   1275
               TabIndex        =   184
               Top             =   60
               Width           =   1280
            End
            Begin VB.CommandButton BtnAddDetalle 
               BackColor       =   &H80000018&
               Caption         =   "Codificar"
               Height          =   640
               Left            =   120
               Picture         =   "mw_ventas_cabecera.frx":1148D
               Style           =   1  'Graphical
               TabIndex        =   183
               ToolTipText     =   "Codifica Equipos"
               Top             =   0
               Visible         =   0   'False
               Width           =   1365
            End
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   11220
            TabIndex        =   72
            Top             =   2550
            Width           =   255
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   8640
            TabIndex        =   71
            Top             =   1820
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   11220
            TabIndex        =   70
            Top             =   3510
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   11220
            TabIndex        =   69
            Top             =   1815
            Width           =   255
         End
         Begin MSDataListLib.DataCombo dtc_preciocompra15 
            Bindings        =   "mw_ventas_cabecera.frx":118CF
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3600
            TabIndex        =   63
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
            Bindings        =   "mw_ventas_cabecera.frx":118E9
            CausesValidation=   0   'False
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6000
            TabIndex        =   48
            Top             =   2160
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
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
            Bindings        =   "mw_ventas_cabecera.frx":11903
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   3480
            TabIndex        =   47
            Top             =   2160
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
         Begin VB.TextBox txt_descripcion_venta 
            CausesValidation=   0   'False
            DataField       =   "concepto_venta"
            DataSource      =   "ado_datos14"
            Height          =   340
            Left            =   225
            MaxLength       =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   3975
            Width           =   9465
         End
         Begin VB.TextBox TxtNroVenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   900
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo dtc_precioventafinal15 
            Bindings        =   "mw_ventas_cabecera.frx":1191D
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6045
            TabIndex        =   37
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
            Bindings        =   "mw_ventas_cabecera.frx":11937
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   6960
            TabIndex        =   38
            Top             =   1800
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_codigo"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc15 
            Bindings        =   "mw_ventas_cabecera.frx":11951
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   360
            TabIndex        =   13
            Top             =   1800
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "bien_descripcion"
            BoundColumn     =   "bien_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc12 
            Bindings        =   "mw_ventas_cabecera.frx":1196B
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   9000
            TabIndex        =   14
            Top             =   1200
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   -2147483644
            ListField       =   "tipoben_descripcion"
            BoundColumn     =   "tipoben_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_aux12 
            Bindings        =   "mw_ventas_cabecera.frx":11985
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10080
            TabIndex        =   40
            Top             =   840
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
            Bindings        =   "mw_ventas_cabecera.frx":1199F
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10320
            TabIndex        =   5
            Top             =   1200
            Visible         =   0   'False
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "almacen_descripcion"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_unimed15 
            Bindings        =   "mw_ventas_cabecera.frx":119B9
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10200
            TabIndex        =   49
            Top             =   1800
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
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
            Bindings        =   "mw_ventas_cabecera.frx":119D3
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10200
            TabIndex        =   51
            Top             =   3495
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
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
            Bindings        =   "mw_ventas_cabecera.frx":119ED
            DataField       =   "tipoben_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   9360
            TabIndex        =   64
            Top             =   840
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
            Bindings        =   "mw_ventas_cabecera.frx":11A07
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10920
            TabIndex        =   66
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
            Bindings        =   "mw_ventas_cabecera.frx":11A21
            DataField       =   "almacen_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   10200
            TabIndex        =   68
            Top             =   2535
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "stock_actual"
            BoundColumn     =   "almacen_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_partida15 
            Bindings        =   "mw_ventas_cabecera.frx":11A3B
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   960
            TabIndex        =   73
            Top             =   2160
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
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
            Bindings        =   "mw_ventas_cabecera.frx":11A55
            DataField       =   "bien_codigo"
            DataSource      =   "ado_datos14"
            Height          =   315
            Left            =   1080
            TabIndex        =   105
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Al ""Grabar"" este registro, se Genera un NUEVO código de Equipo..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   3240
            TabIndex        =   104
            Top             =   915
            Width           =   6240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFF80&
            X1              =   9840
            X2              =   9840
            Y1              =   795
            Y2              =   4320
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Si el equipo ya existe, entonces elige de ""Codigo del Equipo""..."
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
            Height          =   240
            Left            =   3255
            TabIndex        =   65
            Top             =   1200
            Width           =   5625
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   600
            Left            =   10275
            TabIndex        =   52
            Top             =   3000
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Medida"
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
            Left            =   10155
            TabIndex        =   50
            Top             =   1530
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Venta:"
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
            Left            =   360
            TabIndex        =   46
            Top             =   975
            Width           =   1170
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   45
            Top             =   3675
            Width           =   4245
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Equipo"
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
            Left            =   360
            TabIndex        =   44
            Top             =   1530
            Width           =   2100
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código del Equipo"
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
            Left            =   6960
            TabIndex        =   43
            Top             =   1530
            Width           =   1680
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo del Equipo"
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
            Left            =   360
            TabIndex        =   42
            Top             =   2355
            Width           =   1710
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10140
            TabIndex        =   41
            Top             =   2280
            Visible         =   0   'False
            Width           =   1425
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
         Height          =   4350
         Left            =   40
         TabIndex        =   23
         Top             =   380
         Width           =   12015
         Begin VB.TextBox Txt_campo2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   7440
            TabIndex        =   171
            Text            =   "0"
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox Text13 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   6840
            TabIndex        =   107
            Top             =   370
            Width           =   350
         End
         Begin VB.CommandButton Cmd_Cliente 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cliente"
            Height          =   735
            Left            =   7560
            MaskColor       =   &H00C0FFFF&
            Picture         =   "mw_ventas_cabecera.frx":11A6F
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Nuevo Personal"
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text10 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   8880
            TabIndex        =   93
            Top             =   795
            Width           =   330
         End
         Begin MSDataListLib.DataCombo Dtc_deudor2 
            Bindings        =   "mw_ventas_cabecera.frx":11D79
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2685
            TabIndex        =   87
            Top             =   1080
            Visible         =   0   'False
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   255
            ForeColor       =   0
            ListField       =   "beneficiario_deudor"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "mw_ventas_cabecera.frx":11D92
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   3420
            TabIndex        =   86
            Top             =   1020
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
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
            Caption         =   "-- Fecha de Venta --------- Tipo de Venta ------------------------------------------------ Ejecutivo de Ventas"
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
            Height          =   1365
            Left            =   120
            TabIndex        =   32
            Top             =   1635
            Width           =   11775
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
               Left            =   8640
               TabIndex        =   173
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   735
            End
            Begin MSDataListLib.DataCombo dtc_desc11 
               Bindings        =   "mw_ventas_cabecera.frx":11DAB
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   2280
               TabIndex        =   1
               Top             =   270
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "venta_tipo_descripcion"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_desc4 
               Bindings        =   "mw_ventas_cabecera.frx":11DC5
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   6480
               TabIndex        =   3
               Top             =   270
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_denominacion"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "Todos"
            End
            Begin VB.TextBox TxtPlazo 
               Alignment       =   2  'Center
               DataField       =   "venta_plazo_dias_calendario"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   840
               TabIndex        =   2
               Text            =   "0"
               Top             =   510
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtConcepto 
               DataField       =   "venta_descripcion"
               DataSource      =   "Ado_datos"
               Height          =   525
               Left            =   2400
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   705
               Width           =   9135
            End
            Begin MSComCtl2.DTPicker DTPfechasol 
               DataField       =   "venta_fecha"
               DataSource      =   "Ado_datos"
               Height          =   285
               Left            =   240
               TabIndex        =   0
               Top             =   270
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   119537665
               CurrentDate     =   44228
               MinDate         =   32874
            End
            Begin MSDataListLib.DataCombo dtc_codigo11 
               Bindings        =   "mw_ventas_cabecera.frx":11DDE
               DataField       =   "venta_tipo"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   5880
               TabIndex        =   67
               Top             =   255
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "venta_tipo"
               BoundColumn     =   "venta_tipo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtc_codigo4 
               Bindings        =   "mw_ventas_cabecera.frx":11DF8
               DataField       =   "beneficiario_codigo_resp"
               DataSource      =   "Ado_datos"
               Height          =   315
               Left            =   9960
               TabIndex        =   98
               Top             =   480
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "beneficiario_codigo"
               BoundColumn     =   "beneficiario_codigo"
               Text            =   "0"
            End
            Begin VB.Label lbl_concepto 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Concepto de la Venta:"
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
               Height          =   360
               Left            =   240
               TabIndex        =   33
               Top             =   795
               Width           =   2100
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Fra_Total 
            BackColor       =   &H00C0C0C0&
            Caption         =   $"mw_ventas_cabecera.frx":11E11
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
            Height          =   1215
            Left            =   120
            TabIndex        =   25
            Top             =   3060
            Width           =   11775
            Begin VB.TextBox txtTDC 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               DataField       =   "venta_tipo_cambio"
               DataSource      =   "Ado_datos"
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   201
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox TxtAdendaUsd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_adenda_dol"
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
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   168
               Text            =   "0"
               Top             =   405
               Width           =   1425
            End
            Begin VB.TextBox TxtOrigenUsd 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataField       =   "venta_monto_origen_dol"
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   2160
               TabIndex        =   167
               Text            =   "0"
               Top             =   405
               Width           =   1545
            End
            Begin VB.TextBox TxtAdendaBs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_adenda_bs"
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
               Left            =   4200
               Locked          =   -1  'True
               TabIndex        =   164
               Text            =   "0"
               Top             =   795
               Width           =   1425
            End
            Begin VB.TextBox TxtOrigenBs 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               DataField       =   "venta_monto_origen_bs"
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   163
               Text            =   "0"
               Top             =   795
               Width           =   1545
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
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   127
               Text            =   "0"
               Top             =   405
               Width           =   1425
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
               Left            =   8040
               Locked          =   -1  'True
               TabIndex        =   126
               Text            =   "0"
               Top             =   405
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
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   125
               Text            =   "0"
               Top             =   405
               Width           =   1545
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
               Left            =   8040
               Locked          =   -1  'True
               TabIndex        =   28
               Text            =   "0"
               Top             =   795
               Width           =   1545
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
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   6000
               Locked          =   -1  'True
               TabIndex        =   27
               Text            =   "0"
               Top             =   795
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
               Left            =   10125
               Locked          =   -1  'True
               TabIndex        =   26
               Text            =   "0"
               Top             =   795
               Width           =   1425
            End
            Begin VB.Label lbl_campo4 
               Alignment       =   2  'Center
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Cambio:"
               ForeColor       =   &H00400000&
               Height          =   480
               Left            =   120
               TabIndex        =   202
               Top             =   240
               Width           =   780
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00400000&
               X1              =   960
               X2              =   960
               Y1              =   1220
               Y2              =   120
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   5685
               TabIndex        =   170
               Top             =   480
               Width           =   285
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "+/-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3765
               TabIndex        =   169
               Top             =   480
               Width           =   405
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "="
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   5685
               TabIndex        =   166
               Top             =   840
               Width           =   285
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "+/-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3765
               TabIndex        =   165
               Top             =   840
               Width           =   405
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
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   9645
               TabIndex        =   129
               Top             =   435
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   7680
               TabIndex        =   128
               Top             =   435
               Width           =   285
            End
            Begin VB.Label Label7 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "BOB (Bs.) :"
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
               Left            =   1080
               TabIndex        =   124
               Top             =   810
               Width           =   1095
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00400000&
               X1              =   7635
               X2              =   7635
               Y1              =   1220
               Y2              =   120
            End
            Begin VB.Label lbl_totalBs 
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "USD ($US):"
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
               Left            =   1080
               TabIndex        =   31
               Top             =   405
               Width           =   1095
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
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   7695
               TabIndex        =   30
               Top             =   765
               Width           =   285
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
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   9645
               TabIndex        =   29
               Top             =   765
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   9585
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "mw_ventas_cabecera.frx":11EA6
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   960
            TabIndex        =   80
            Top             =   1260
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "mw_ventas_cabecera.frx":11EBF
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4440
            TabIndex        =   83
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
            Bindings        =   "mw_ventas_cabecera.frx":11ED9
            DataField       =   "unidad_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1515
            TabIndex        =   84
            Top             =   360
            Width           =   5685
            _ExtentX        =   10028
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
            Bindings        =   "mw_ventas_cabecera.frx":11EF2
            DataField       =   "beneficiario_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1680
            TabIndex        =   88
            Top             =   1080
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   -2147483624
            ListField       =   "tipoben_codigo"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_aux3 
            Bindings        =   "mw_ventas_cabecera.frx":11F0B
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   9300
            TabIndex        =   90
            Top             =   720
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "codigo5"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo3 
            Bindings        =   "mw_ventas_cabecera.frx":11F24
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7620
            TabIndex        =   91
            Top             =   780
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "codigo"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc3 
            Bindings        =   "mw_ventas_cabecera.frx":11F3D
            DataField       =   "edif_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   960
            TabIndex        =   92
            Top             =   780
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            BackColor       =   12632256
            ForeColor       =   0
            ListField       =   "descripcion"
            BoundColumn     =   "codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo8 
            Bindings        =   "mw_ventas_cabecera.frx":11F56
            DataField       =   "codigo_empresa"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   10680
            TabIndex        =   203
            Top             =   960
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "codigo_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc8 
            Bindings        =   "mw_ventas_cabecera.frx":11F6F
            DataField       =   "codigo_empresa"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   6240
            TabIndex        =   204
            Top             =   1260
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "denominacion_empresa"
            BoundColumn     =   "codigo_empresa"
            Text            =   ""
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "EMPRESA:"
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
            Left            =   5160
            TabIndex        =   205
            Top             =   1260
            Width           =   1035
         End
         Begin VB.Label LblEmpresa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            DataSource      =   "Ado_datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   540
            Left            =   11040
            TabIndex        =   200
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cite de Contrato"
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
            Height          =   285
            Left            =   7440
            TabIndex        =   172
            Top             =   75
            Width           =   1725
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
            Left            =   3120
            TabIndex        =   158
            Top             =   0
            Width           =   4875
         End
         Begin VB.Label txt_campo1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   11160
            TabIndex        =   153
            Top             =   720
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label txt_codigo1 
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   10440
            TabIndex        =   152
            Top             =   720
            Visible         =   0   'False
            Width           =   765
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
            TabIndex        =   89
            Top             =   780
            Width           =   705
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
            TabIndex        =   85
            Top             =   345
            Width           =   1215
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
            TabIndex        =   82
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
            Left            =   1545
            TabIndex        =   81
            Top             =   75
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   9540
            TabIndex        =   35
            Top             =   75
            Width           =   1245
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
            Left            =   180
            TabIndex        =   34
            Top             =   1260
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
      TabIndex        =   76
      Top             =   720
      Width           =   6345
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
         Left            =   3960
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   4520
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "mw_ventas_cabecera.frx":11F88
         Height          =   4170
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   6120
         _ExtentX        =   10795
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
         ColumnCount     =   8
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
            DataField       =   "estado_codigo_verif"
            Caption         =   "Verificado"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
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
         Width           =   6120
         _ExtentX        =   10795
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
      Caption         =   "DETALLE DE EQUIPOS / BIENES"
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
      Height          =   1335
      Left            =   2160
      TabIndex        =   21
      Top             =   5580
      Width           =   16455
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "mw_ventas_cabecera.frx":11FA0
         Height          =   1065
         Left            =   240
         TabIndex        =   22
         Top             =   225
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   1879
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
            Caption         =   "Modelo.Elegido"
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
   Begin VB.Frame FrmCobranza 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PLAN DE CUOTAS PARA COBROS AL CLIENTE "
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
      Height          =   3150
      Left            =   2160
      TabIndex        =   19
      Top             =   6930
      Width           =   16455
      Begin MSDataGridLib.DataGrid DtgCobro 
         Bindings        =   "mw_ventas_cabecera.frx":11FBA
         Height          =   2820
         Left            =   255
         TabIndex        =   20
         Top             =   240
         Width           =   16110
         _ExtentX        =   28416
         _ExtentY        =   4974
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
         ColumnCount     =   12
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
            Caption         =   "Fecha.de.la.Cuota"
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
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto a Pagar Bs."
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
         BeginProperty Column04 
            DataField       =   "doc_codigo_fac"
            Caption         =   "Fac/Recibo"
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
            Caption         =   "Concepto para Factura u O.C."
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
            DataField       =   "cobranza_codigo"
            Caption         =   "Nro.Cobranza"
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
         BeginProperty Column11 
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
               Locked          =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   3179.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2775.118
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
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
      Left            =   2280
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
      TabIndex        =   148
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
         Picture         =   "mw_ventas_cabecera.frx":11FD4
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   150
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
         Picture         =   "mw_ventas_cabecera.frx":127AA
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   149
         Top             =   0
         Width           =   1455
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
         TabIndex        =   151
         Top             =   180
         Width           =   885
      End
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   0
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
      Left            =   6840
      Top             =   10920
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   9120
      Top             =   10920
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "mw_ventas_cabecera"
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
Dim rs_aux14 As New ADODB.Recordset
Dim rs_aux15 As New ADODB.Recordset
Dim rs_aux16 As New ADODB.Recordset
Dim rs_aux17 As New ADODB.Recordset
Dim rs_aux18 As New ADODB.Recordset
Dim rs_aux19 As New ADODB.Recordset
Dim rs_aux20 As New ADODB.Recordset

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
Dim VAR_PARADA As Integer
Dim VAR_ZONA, VAR_ZPILOTO As Integer
Dim VAR_DIA, VAR_MES As Integer
Dim VAR_IDTAREA, VAR_NRODIAS As Integer

'contabilidad
Dim VAR_EMPRESA, VAR_TIPOCOMPID, VAR_MONEDAID As Integer
Dim VAR_TIPOCAMBIO As Double
Dim EntregadoA, VAR_CONCEPTO As String

Dim VAR_COMPM, VAR_PLANID As Long

Dim VAR_FECHAINI, VAR_FECHAFIN As Date
Dim VAR_FCTRLINI, VAR_FCTRLFIN, VAR_FECHACTRL As Date

Dim VAR_DCORR, VAR_HCORR As String

Dim Cobrobs, VAR_COBR, VAR_AUX, VAR_AUX2 As Double
Dim VAR_Bs, VAR_Dol, VAR_BS2, VAR_DOL2, VAR_MBS2, VAR_MDOL2 As Double
Dim VAR_AUX4, VAR_AUX5 As Double
Dim VAR_RECORRIDO, VAR_VELOCIDAD As Double
Dim VAR_PASAJEROS, VAR_PARADAS As String

Dim gestion0, var_literal, VAR_PROY2, VAR_CITE, VAR_CTA As String
Dim VAR_CODTIPO, VAR_ORG, VAR_FTE, VAR_BENEF, VAR_GLOSA, VAR_GLOSA2, VAR_MONEDA As String
Dim VAR_BEND, VAR_EDIFD, VARG_ORGD, VAR_CTAD, VAR_UNID, VAR_DPTO, VAR_DPTOD As String
Dim VAR_COD1, VAR_COD2, VAR_COD3, VAR_UNIDCOD As String
Dim VAR_TIPOV, VAR_UNIMED As String
Dim VAR_COBR0, VAR_OA, VAR_OA2, VAR_NEW As String
Dim VAR_PAIS, VAR_EQP, VAR_TIPOEQP As String
Dim VAR_DA, VAR_UORIGEN As String
Dim VAR_NOMD, VAR_NOMH As String
Dim VAR_JQ, VAR_VAL As String
Dim VAR_TRAMITE As String
Dim VAR_COMPRA, VAR_DOCFAC As String
Dim VAR_DESTAREA, VAR_BIEN As String
Dim VAR_BENINST, VAR_BENAJST, VAR_AUX1 As String
    
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
   If Not IsNull(Ado_datos.Recordset!venta_codigo) Then
        If Ado_datos.Recordset!codigo_empresa = "2" Then
            LblEmpresa.Caption = "CGE"
        Else
            LblEmpresa.Caption = "CGI"
        End If
        nroventa = Ado_datos.Recordset!venta_codigo
        lbl_cerrado.Caption = ""
        If (Ado_datos.Recordset!estado_codigo = "REG") Then
            If glusuario = "DTERCEROS" Or glusuario = "CPLATA" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CPAREDES" Or glusuario = "KGARCIA" Or glusuario = "EVILLALOBOS" Or glusuario = "LVEDIA" Or glusuario = "JCASTRO" Or glusuario = "ADMIN" Or glusuario = "CSALINAS" Then
                BtnAprobar.Visible = True   'APROBAR Tramite
                BtnEliminar.Visible = True  'ANULAR Tramite
                BtnAñadir.Visible = True    'CERRAR Tramite
                BtnVer.Visible = False
            Else
                BtnAprobar.Visible = False   'APROBAR Tramite
                BtnEliminar.Visible = False  'ANULAR Tramite
                BtnAñadir.Visible = False    'CERRAR Tramite
                BtnVer.Visible = True
            End If
            BtnAprobar1.Visible = True
            BtnDesAprobar.Visible = False
            BtnModificar.Visible = True
            BtnModDetalle1.Visible = True
'            BtnAprobar.Visible = False   'APROBAR Tramite
'            BtnEliminar.Visible = False  'ANULAR Tramite
            If IsNull(Ado_datos.Recordset!venta_tipo) Then
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
                    If Ado_datos.Recordset!estado_codigo = "ANL" Then
                        lbl_cerrado.Caption = "TRAMITE ANULADO !!"
                        FrmABMDet2.Visible = False
                        BtnAñadir.Visible = False   'Cerrar Tramite
                        BtnVer3.Visible = False     'Provisional
                        FrmABMDet.Visible = False
                        FrmABMDet1.Visible = False
'                        BtnAprobar.Visible = False   'APROBAR Tramite
'                        BtnEliminar.Visible = False  'ANULAR Tramite
                    Else
                        lbl_cerrado.Caption = ""
                        FrmABMDet2.Visible = True
'                        BtnAprobar.Visible = False
                        BtnDesAprobar.Visible = True
                        BtnModificar.Visible = False
                        BtnModDetalle1.Visible = False
                        FrmABMDet.Visible = False
                        FrmABMDet1.Visible = False
                        BtnVer.Visible = True
                    End If
            End Select
'            BtnEliminar.Visible = False
'            BtnVer.Visible = True
            FrmABMDet2.Visible = True
            FrmCobranza.Visible = True
            FrmAlcance.Visible = True
            If (Ado_datos.Recordset!estado_codigo = "APR") Then
                'CRONOGRAMA COMPRA SERVICIO
                'Compra Cabecera Funcionario - Vendedor
                Set rs_datos8 = New ADODB.Recordset
                If rs_datos8.State = 1 Then rs_datos8.Close
                rs_datos8.Open "select * from ao_compra_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenStatic
                'Set Ado_datos4.Recordset = rs_datos8
                'Compra Adjudica
                Set rs_det2 = New ADODB.Recordset
                If rs_det2.State = 1 Then rs_det2.Close
                rs_det2.Open "select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
                Set Ado_detalle2.Recordset = rs_det2
                If Ado_detalle2.Recordset.RecordCount > 0 Then
                    Set rs_det3 = New ADODB.Recordset
                    If rs_det3.State = 1 Then rs_det3.Close
                    rs_det3.Open "select * from ao_compra_planilla_pagos where compra_codigo = " & rs_det2!compra_codigo & " and adjudica_codigo = " & rs_det2!adjudica_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
                    Set Ado_detalle3.Recordset = rs_det3
'                    If Ado_detalle3.Recordset.RecordCount > 0 Then
'                        dg_det3.Visible = True
'                        Set dg_det3.DataSource = Ado_detalle3.Recordset
'                    Else
'                        dg_det3.Visible = False
'                        Set dg_det3.DataSource = rsNada
'                    End If
                Else
'                    dg_det3.Visible = False
'                    Set dg_det3.DataSource = rsNada
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
'            TxtPlazo.Visible = True
            BtnAddDetalle2.Visible = True
        Else
'            TxtPlazo.Visible = False
            If Ado_datos.Recordset("venta_tipo") = "E" Then
                BtnAddDetalle2.Visible = False
            End If
        End If
        
        Call ABRIR_TABLA_DET
'        FrmDetalle.Caption = "BIENES DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
'        FrmCobranza.Caption = "CRONOGRAMA DE COBRANZAS DE LA VENTA NRO. " + Str((Ado_datos.Recordset("venta_codigo")))
        
        FrmDetalle.Caption = "BIENES DEL TRAMITE NRO. " + Str((Ado_datos.Recordset("solicitud_codigo")))
        FrmCobranza.Caption = "CRONOGRAMA DE COBRANZAS DE TRAMITE NRO. " + Str((Ado_datos.Recordset("solicitud_codigo")))

        End If
        GlEdificio = Ado_datos.Recordset!edif_codigo
        FrmDetalle.Visible = True
        FrmCobranza.Visible = True
'        FrmAlcance.Visible = True
    Else
        FrmABMDet.Visible = False
        FrmABMDet1.Visible = False
        FrmABMDet2.Visible = False
'        FrmAlcance.Visible = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
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

Private Sub BtnAddDetalle1_Click()
    NumComp = Ado_datos.Recordset!venta_codigo
    gestion0 = Ado_datos.Recordset!ges_gestion
    sino = MsgBox("Elije SI, para borrar los Registros del Alcance y Volver a Cargarlos ? . Elije NO, para Cancelar... ", vbYesNo, "Confirmando")
    If sino = vbYes Then
        db.Execute "DELETE ao_ventas_alcance WHERE venta_codigo = " & NumComp & " "
    Else
        Exit Sub
    End If
    sino = MsgBox("El Contrato incluye... SERVICIO DE DESMONTAJE ? ...", vbYesNo, "Confirmando")
    Set rs_aux19 = New ADODB.Recordset
    If rs_aux19.State = 1 Then rs_aux19.Close
    If sino = vbYes Then
       rs_aux19.Open "Select * from gc_tipo_solicitud where solicitud_num = '90' order by ORDEN ", db, adOpenStatic
    Else
       rs_aux19.Open "Select * from gc_tipo_solicitud where (solicitud_num = '90' and solicitud_tipo <> '20') order by ORDEN ", db, adOpenStatic
    End If
    'Set Ado_datos1.Recordset = rs_aux19
    If rs_aux19.RecordCount > 0 Then
        'ao_ventas_alcance
        rs_aux19.MoveFirst
        While Not rs_aux19.EOF
            'solicitud_tipo, solicitud_tipo_descripcion, estado_codigo, fecha_registro, hora_registro, usr_codigo, solicitud_num, unidad_codigo, cta_dev_1, cta_dev_2, cta_fac_1, cta_fac_2, cta_cob_1, cta_cob_2, orden

            'ges_gestion, venta_codigo, solicitud_tipo, venta_codigo_new, solicitud_tipo_descripcion, unidad_codigo_tec, venta_tiempo_dias, fecha_inicio_alcance, fecha_fin_alcance, fecha_inicio_real, fecha_fin_real,
            '             doc_codigo , correl_doc, estado_codigo, usr_codigo, fecha_registro, hora_registro, estado_acta, estado_mantenimiento
            
            db.Execute "INSERT INTO ao_ventas_alcance (ges_gestion, venta_codigo, solicitud_tipo, venta_codigo_new, solicitud_tipo_descripcion, unidad_codigo_tec, venta_tiempo_dias, fecha_inicio_alcance, fecha_fin_alcance, fecha_inicio_real, fecha_fin_real, " & _
            " doc_codigo , correl_doc, estado_codigo, usr_codigo, fecha_registro, hora_registro, estado_acta, estado_mantenimiento, orden) " & _
            " VALUES ('" & gestion0 & "', " & NumComp & ", " & rs_aux19!solicitud_tipo & ", '0', '" & rs_aux19!solicitud_tipo_descripcion & "' , '" & rs_aux19!unidad_codigo & "','0', '01/01/1900' , '01/01/1900', '01/01/1900' , '01/01/1900', " & _
            " 'R-321', '0', 'REG', '" & glusuario & "', '" & Date & "', '0', 'REG', 'REG', " & rs_aux19!Orden & ") "
            
            rs_aux19.MoveNext
        Wend
    Else
        
    End If
    Call ABRIR_TABLA_DET
End Sub

Private Sub BtnAñadir_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
  If Ado_datos.Recordset.RecordCount > 0 Then
    If Ado_datos.Recordset!estado_cancelado = "N" And Ado_datos.Recordset!estado_codigo = "APR" Then
      sino = MsgBox("Esta seguro de CERRAR EL TRAMITE, ya no podrá realizar modificaciones... ", vbYesNo, "Confirmando")
      If sino = vbYes Then
          db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_cancelado = 'S' Where ao_ventas_cabecera.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  "
          db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'ANL' Where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " and estado_codigo = 'REG' "
          marca1 = Ado_datos.Recordset.Bookmark
          'Ado_datos.Recordset.Requery
          'Ado_datos.Refresh
          If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
          Else
            Call OptFilGral2_Click
          End If
          Ado_datos.Recordset.Move marca1 - 1
      End If
    Else
      MsgBox "NO se puede procesar el TRAMITE ya fue CERRADO...", , "Atencion"
    End If
  Else
    MsgBox "NO se puede procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If Ado_datos.Recordset.RecordCount > 0 Then
    'VALIDA EDIFICIO Y EQUIPOS
    Set rs_aux10 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_aux10.State = 1 Then rs_aux10.Close
    rs_aux10.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & dtc_codigo3.Text & "' and estado_codigo = 'APR' ", db, adOpenStatic
    If rs_aux10.RecordCount = 0 Then
        'Si Faltarian Aprobar
        MsgBox "No se puede APROBAR, verifique los datos del Edificio si estan correctos y si está Aprobado, luego vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    
    Set rs_aux11 = New ADODB.Recordset     'Equipos de Venta_Detalle
    If rs_aux11.State = 1 Then rs_aux11.Close
    rs_aux11.Open "Select * from mv_bienes_vs_venta_det WHERE venta_codigo = '" & Ado_datos.Recordset!venta_codigo & "'  ", db, adOpenStatic
    If rs_aux11.RecordCount > 0 Then
        'Si Faltarian REGISTRAR
        MsgBox "No se puede APROBAR, verifique los datos de los EQUIPOS y si estos están Registrados, luego vuelva a intentar ...", , "Atención"
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
    'Estado VERIFICADO
    If Ado_datos.Recordset!estado_codigo_verif = "REG" Then
        MsgBox "No se puede APROBAR, el Ejecutivo de Ventas debe VERIFICAR previamente el registro, luego vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    
   If IsNull(Ado_datos.Recordset("venta_tipo")) Or Ado_datos.Recordset("venta_tipo") = "" Or (Ado_datos.Recordset("venta_monto_total_bs") = 0) Or (Ado_datos.Recordset!unidad_codigo_ant = "") Or IsNull(Ado_datos.Recordset!unidad_codigo_ant) Then
        MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
        Exit Sub
   Else
     If Ado_datos.Recordset("estado_codigo") = "REG" Then
       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
       If sino = vbYes Then
           'ASIGNA A VARIABLES CAMPOS CLAVES
           gestion0 = Ado_datos.Recordset!ges_gestion
           correlv = Ado_datos.Recordset!venta_codigo
           NumComp = Ado_datos.Recordset!venta_codigo
           VAR_SOL = Ado_datos.Recordset!solicitud_codigo
           VAR_TIPOV = Ado_datos.Recordset!venta_tipo
           VAR_PROY2 = Ado_datos.Recordset!edif_codigo
           GlEdificio = Ado_datos.Recordset!edif_codigo
           VAR_UNIDCOD = Ado_datos.Recordset!unidad_codigo
           VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
           VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant
           VAR_GLOSA = Ado_datos.Recordset!venta_descripcion
           VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
           VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
           VAR_UNIMED = Ado_datos.Recordset!unimed_codigo
           VAR_BEND = dtc_desc2.Text
           VAR_EDIFD = dtc_desc3.Text
           VAR_UNID = dtc_desc1.Text
           VAR_DPTO = Left(VAR_PROY2, 1)    'Ado_datos.Recordset!depto_codigo
           VARG_ORGD = ""
           VAR_CTAD = ""
           VAR_MED = IIf(Ado_datos.Recordset!unimed_codigo <> "MES", "MES", Ado_datos.Recordset!unimed_codigo)
           VAR_EMPRESA = Ado_datos.Recordset!codigo_empresa
           VAR_TIPO = Ado_datos.Recordset!solicitud_tipo
           VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
           If VAR_ZONA = "" Or IsNull(VAR_ZONA) Or VAR_ZONA = 0 Then
                Set rs_datos6 = New ADODB.Recordset
                If rs_datos6.State = 1 Then rs_datos6.Close
                rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & GlEdificio & "'    ", db, adOpenStatic
                If rs_datos6.RecordCount > 0 Then
                    VAR_ZONA = rs_datos6!zpiloto_codigo
                Else
                    VAR_ZONA = 0
                    MsgBox "El Edificio de este contrato no tiene una ZONA PILOTO, se le asignará la Zona de Gratuitos que le corresponda ...", , "Atención"
                    'MsgBox "NO se puede Aprobar, debe registrar la Zona Piloto para su Mantenimiento Gratuito !!. Consulte con Area Técnica. ", vbExclamation, "Atención!"
                    'Exit Sub
                End If
           End If
           db.Execute "update ao_ventas_cabecera set zpiloto_codigo = " & VAR_ZONA & " Where venta_codigo = " & correlv & " "
           'APRUEBA ALCANCE DEL CONTRATO Y EL CONTRATO
           db.Execute "update ao_ventas_alcance set estado_codigo = 'APR' Where venta_codigo = " & correlv & " "
           db.Execute "update ao_ventas_cabecera set estado_alcance = 'S' Where venta_codigo = " & correlv & " "
           
           'INI Deptos de Bolivia
            Select Case VAR_DPTO
                 Case "1"
                     VAR_DPTOD = "CHUQUISACA"
                 Case "2"
                     VAR_DPTOD = "LA PAZ"
                 Case "3"
                     VAR_DPTOD = "COCHABAMBA"
                 Case "4"
                     VAR_DPTOD = "ORURO"
                 Case "5"
                     VAR_DPTOD = "POTOSI"
                 Case "6"
                     VAR_DPTOD = "TARIJA"
                 Case "7"
                     VAR_DPTOD = "SANTA CRUZ"
                 Case "8"
                     VAR_DPTOD = "BENI"
                 Case "9"
                     VAR_DPTOD = "PANDO"
            End Select

'           'INI CONTABILIZACION NUEVA
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW

            'Call Contabiliza_Contratos

'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            'FIN CONTABILIZACION NUEVA
'           Call Contabiliza_venta
           
           'INI GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
           If VAR_TIPOV = "V" Or VAR_TIPOV = "L" Or VAR_TIPOV = "G" Then
           'If Ado_datos.Recordset!venta_tipo = "V" Then
             Set rs_aux1 = New ADODB.Recordset
             If rs_aux1.State = 1 Then rs_aux1.Close
             rs_aux1.Open "select * from ao_ventas_alcance where venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux1.RecordCount > 0 Then
               rs_aux1.MoveFirst
               
               While Not rs_aux1.EOF
                 VAR_COD1 = rs_aux1!unidad_codigo_tec
                 VAR_CANT0 = Round((rs_aux1!venta_tiempo_dias / 30), 0)
                 'rs_aux1.MoveNext
                 If VAR_COD1 = "COMEX" Or VAR_COD1 = "DVTA" Then         'INI GRABA CRONOGRAMA COMEX
                    ' AV_COMPRA_VS_ADJUDICA
                    Set rs_aux20 = New ADODB.Recordset
                    If rs_aux20.State = 1 Then rs_aux20.Close
                    rs_aux20.Open "select * from AV_COMPRA_VS_ADJUDICA where edif_codigo = '" & VAR_PROY2 & "'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux20.RecordCount > 0 Then
                        correldetalle = rs_aux20!compra_codigo
                        db.Execute "UPDATE ao_compra_cabecera SET unidad_codigo_ant = '" & VAR_CITE & "' where compra_codigo = " & correldetalle & "  "
                        VAR_COMPRA = "SI"
                    Else
                        'EQUIPO
                        Set rs_aux2 = New ADODB.Recordset
                        If rs_aux2.State = 1 Then rs_aux2.Close
                        rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux2.RecordCount > 0 Then
                           rs_aux2!correl_negocia = rs_aux2!correl_negocia + 1
                           correldetalle = rs_aux2!correl_negocia
                           rs_aux2.Update
                        End If
                        VAR_COMPRA = "NO"
                    End If
                    'WWWWWWWWWWWWWWW
                    'correlv = Ado_datos.Recordset!venta_codigo
                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux3.RecordCount = 0 Then
                    'CAMPOS FALTANTES
                    ' beneficiario_codigo_resp, doc_numero, nro_nota_remision, estado_codigo_tra, estado_codigo_nac, estado_codigo_des,
                    ' hora_registro, usr_codigo_aprueba, fecha_registro_aprueba, archivo_respaldo, archivo_respaldo_cargado, estado_codigo_tec, adjudica_codigo
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion     'Year(Date)
                        'rs_aux3!compra_codigo = 0      'Autonumerico
                        rs_aux3!unidad_codigo_adm = VAR_COD1
                        rs_aux3!solicitud_codigo_adm = correldetalle
                        rs_aux3!unidad_codigo = VAR_UNIDCOD
                        rs_aux3!solicitud_codigo = VAR_SOL
                        rs_aux3!edif_codigo = VAR_PROY2
                        rs_aux3!beneficiario_codigo = VAR_BENEF
                        rs_aux3!beneficiario_codigo_alm = IIf(IsNull(Ado_datos.Recordset!beneficiario_codigo_resp), "0", Ado_datos.Recordset!beneficiario_codigo_resp)
                        rs_aux3!solicitud_tipo = rs_aux1!solicitud_tipo     '"15"
                        rs_aux3!venta_tipo = VAR_TIPOV
                        rs_aux3!unidad_codigo_ant = VAR_CITE
                        rs_aux3!compra_fecha = Date
                        rs_aux3!compra_DESCRIPCION = "COMPRA POR: " + VAR_GLOSA
                        rs_aux3!compra_observaciones = "PROVISION Y/O IMPORTACION DE EQUIPOS"
                        rs_aux3!compra_cantidad_total = Ado_datos.Recordset!venta_cantidad_total
                        rs_aux3!compra_monto_bs = VAR_BS2
                        rs_aux3!tipo_moneda = "USD"
                        rs_aux3!compra_monto_dol = VAR_DOL2
                        rs_aux3!proceso_codigo = "CMX"
                        rs_aux3!subproceso_codigo = "CMX-01"
                        rs_aux3!etapa_codigo = "CMX-01-01"
                        rs_aux3!clasif_codigo = "CMX"
                        rs_aux3!doc_codigo = "R-207"
                        rs_aux3!poa_codigo = "4.1.1"
                        rs_aux3!doc_codigo_alm = "R-207"
                        rs_aux3!beneficiario_codigo_resp = "4828818"
                        'doc_numero_alm
                        'GENERAR CORRELATIVO
                        rs_aux3!estado_codigo_eqp = "REG"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3!fecha_registro = Date
                        rs_aux3.Update
                        'INI ACTUALIZA CORRELATIVO POR ALMACEN
                        Set rs_aux9 = New ADODB.Recordset
                        If rs_aux9.State = 1 Then rs_aux9.Close
                        SQL_FOR = "select * from ac_almacenes where almacen_codigo = 1  "        '" & Val(dtc_codigo11.Text) & "
                        rs_aux9.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                        If rs_aux9.RecordCount > 0 Then
                            rs_aux9!correl_sal = rs_aux9!correl_sal + 1
                            VAR_NUM = rs_aux9!correl_sal
                            rs_aux9.Update
                        Else
                            VAR_NUM = 1
                        End If
                        db.Execute "UPDATE ao_compra_cabecera SET doc_numero_alm = " & VAR_NUM & " where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " "
                        'FIN ACTUALIZA CORRELATIVO POR ALMACEN
                        'db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " &
                            '"VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
                           
                        'DETALLE Carga ao_ventas_detalle
                        'Set rstdestino = New ADODB.Recordset
                        'If rstdestino.State = 1 Then rstdestino.Close
                        'rstdestino.Open "select * from ao_compra_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
                        'INI DISTRIBUYE TRAMITES EN ao_compra_detalle
                        Select Case rs_aux1!solicitud_tipo

                            Case 3
                                VAR_TRAMITE = "BANCO"
                                Set rs_aux4 = New ADODB.Recordset
                                If rs_aux4.State = 1 Then rs_aux4.Close
                                rs_aux4.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & " AND PAR_CODIGO = '43340' ", db, adOpenKeyset, adLockBatchOptimistic
                                If rs_aux4.RecordCount > 0 Then
                                   rs_aux4.MoveFirst
                                   While Not rs_aux4.EOF
                                        db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
                                        "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & rs_aux4!concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
                                        rs_aux4.MoveNext
                                   Wend
                                Else
                                    MsgBox "No existe Equipos, verifique el registro y vuelva a intentar ... ", vbInformation, "Información!"
                                End If
'                                If rstdestino.State = 1 Then rstdestino.Close
                            Case 16
                                VAR_TRAMITE = "TRANS"
                            Case 17
                                VAR_TRAMITE = "ADUAN"
                            Case 18
                                VAR_TRAMITE = "DESCA"
                            Case Else
                                VAR_TRAMITE = "BANCO"
                        End Select
                        
                        Set rs_aux4 = New ADODB.Recordset
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        rs_aux4.Open "select * from ac_bienes where bien_codigo_anterior= '" & VAR_TRAMITE & "' AND KIT = '90'  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux4.RecordCount > 0 Then
                           rs_aux4.MoveFirst
                           While Not rs_aux4.EOF
                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,            grupo_codigo,                   subgrupo_codigo,                    par_codigo,                 tipo_descuento, almacen_codigo , usr_usuario,       fecha_registro) " & _
                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1',        '0',                    '0',                    '0',                    '0',                        '0',                '0',                    '" & rs_aux4!bien_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1',           '0',            '" & glusuario & "', '" & Date & "')"
                                rs_aux4.MoveNext
                           Wend
                        End If
                        'cargar ADJUDICA_COMPRA Y CRONOGRAMA
                        '
                    Else
                        Select Case rs_aux1!solicitud_tipo
                            Case 3
                                VAR_TRAMITE = "BANCO"
                            Case 16
                                VAR_TRAMITE = "TRANS"
                            Case 17
                                VAR_TRAMITE = "ADUAN"
                            Case 18
                                VAR_TRAMITE = "DESCA"
                            Case Else
                                VAR_TRAMITE = "BANCO"
                        End Select
                        If VAR_COMPRA = "NO" Then
                            Set rs_aux4 = New ADODB.Recordset
                            If rs_aux4.State = 1 Then rs_aux4.Close
                            rs_aux4.Open "select * from ac_bienes where bien_codigo_anterior= '" & VAR_TRAMITE & "' AND KIT = '90'  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux4.RecordCount > 0 Then
                               rs_aux4.MoveFirst
                               While Not rs_aux4.EOF
                                    db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto,            grupo_codigo,                   subgrupo_codigo,                    par_codigo,                 tipo_descuento, almacen_codigo , usr_usuario,       fecha_registro) " & _
                                    "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", '" & rs_aux4!bien_codigo & "', '1',        '0',                    '0',                    '0',                    '0',                        '0',                '0',                    '" & bien_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1',           '0',            '" & glusuario & "', '" & Date & "')"
                                    rs_aux4.MoveNext
                               Wend
                            End If
                        End If
                    End If
                    'WWWWWWWWWW
                 Else
                    If rs_aux1!solicitud_tipo = 4 Then
                        'BUSCA ZONA PILOTO INSTALACION
                        VAR_ZPILOTO = Ado_datos.Recordset!depto_codigo
                        VAR_FECHAINI = rs_aux1!fecha_inicio_alcance
                        VAR_FECHAFIN = rs_aux1!fecha_fin_alcance
                        Set rs_aux2 = New ADODB.Recordset
                        If rs_aux2.State = 1 Then rs_aux2.Close
                        rs_aux2.Open "SELECT * FROM tc_zona_piloto_edif_inst WHERE EDIF_codigo = '" & GlEdificio & "' ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux2.RecordCount > 0 Then
                            VAR_ZPILOTO = rs_aux2!zpiloto_codigo
                        Else
                            Set rs_aux18 = New ADODB.Recordset
                            If rs_aux18.State = 1 Then rs_aux18.Close
                            rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif_inst where zpiloto_codigo = " & VAR_ZPILOTO & " ", db, adOpenKeyset, adLockOptimistic
                            If rs_aux18.RecordCount > 0 Then
                                VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
                                DIA_ORDEN = VAR_ORDEN
                            Else
                                VAR_ORDEN = 1
                            End If
                            'gestion0 = Ado_datos.Recordset!ges_gestion
                            'VAR_MED = IIf(Ado_datos.Recordset!unimed_codigo <> "MES", "MES", Ado_datos.Recordset!unimed_codigo)
                            'VAR_EMPRESA = Ado_datos.Recordset!codigo_empresa
                            'VAR_TIPO = Ado_datos.Recordset!solicitud_tipo
                            'CREA EDIFICIO EN ORGANIZACION DE ZONAS
                            db.Execute "INSERT INTO tc_zona_piloto_edif_inst (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
                                      " Gratuito, fecha_ini_max,           fecha_fin_max,       venta_codigo, estado_codigo, estado_activo, fecha_registro, usr_codigo,     unimed_codigo,      codigo_empresa,     solicitud_tipo) " & _
                                      " VALUES (" & VAR_ZPILOTO & ", '" & GlEdificio & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
                                      " 'SI', '" & VAR_FECHAINI & "', '" & VAR_FECHAFIN & "',  " & NumComp & ",  'REG',         'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ") "
                                      
                            ' ASIGNA ZPILOTO EN ao_ventas_cabecera
                            'db.Execute "UPDATE ao_ventas_cabecera SET zpiloto_codigo = " & VAR_ZPILOTO & " WHERE venta_codigo = " & NumComp & "    "
                            
                            'MsgBox "Se asignó la ZONA PILOTO=" + Str(VAR_ZPILOTO) + ", para generar el Cronograma de Mantenimiento Gratuito... ", , "Atencion"
                            'Exit Sub
                        End If
                        'GENERA CRONOGRAMA INSTALACION
                        Call CRONO_INSTALACION
                    End If
                    'If VAR_COD1 = "DNMAN" Then          'INI GRABA CRONOGRAMA MANTENIMIENTO
                        
                        'Call CRONO_MTTO
                    'End If
                    '
                 End If
                 rs_aux1.MoveNext
               Wend
             End If
           End If
           ' APRUEBA ao_ventas_cabecera
               Ado_datos.Recordset!estado_codigo = "APR"
               Ado_datos.Recordset.Update
           'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
           ''Actualiza Cite Trñamite (unidad_codigo_ant)
           db.Execute "update ao_solicitud set ao_solicitud.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud inner join ao_ventas_cabecera on ao_solicitud.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
           db.Execute "update ao_solicitud_calculo_trafico set ao_solicitud_calculo_trafico.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_calculo_trafico inner join ao_ventas_cabecera on ao_solicitud_calculo_trafico.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_calculo_trafico.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
           db.Execute "update ao_solicitud_cotiza_modelo set ao_solicitud_cotiza_modelo.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_cotiza_modelo inner join ao_ventas_cabecera on ao_solicitud_cotiza_modelo.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_cotiza_modelo.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
           db.Execute "update ao_solicitud_cotiza_venta set ao_solicitud_cotiza_venta.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_cotiza_venta inner join ao_ventas_cabecera on ao_solicitud_cotiza_venta.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_cotiza_venta.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
           'db.Execute "UPDATE co_diario SET co_diario.estado_codigo = co_comprobante_m.estado_codigo FROM co_diario INNER JOIN co_comprobante_m ON co_diario.Cod_Comp =co_comprobante_m.Cod_Comp where co_diario.estado_codigo Is Null "
           'FIN GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
           'Call OptFilGral1_Click
           MsgBox "La Venta fue Aprobada Exitosamente... ", vbInformation, "Información!"
       End If
     End If
        'MsgBox "Verifique si el Registro ya fue APROBADO o ANULADO previamente ...", , "Atención"
   End If
 Else
    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
 End If

End Sub

Private Sub CRONO_INSTALACION()
    Set rs_aux0 = New ADODB.Recordset
    If rs_aux0.State = 1 Then rs_aux0.Close
    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & GlEdificio & "'   ", db, adOpenStatic
    If rs_aux0.RecordCount > 0 Then
        VAR_EDIF = Ado_datos.Recordset!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
    End If
    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
    
    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'    ' jalar ORDEN de tc_zona_piloto_edif
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from tc_zona_piloto_edif_inst WHERE edif_codigo = '" & GlEdificio & "'    ", db, adOpenStatic
'    If rs_datos6.RecordCount > 0 Then
'        DIA_ORDEN = rs_datos6!zona_edif_orden
'    Else
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
'                  " VALUES (" & VAR_ZONA & ", '" & VAR_PROY2 & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
'                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
'        DIA_ORDEN = "1"
'    End If
'    'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
'    MControl = Ado_datos.Recordset!mes_inicio_crono_tec                     'mes_inicio_crono
    MControl = UCase(MonthName(Month(VAR_FECHAINI)))
    'MonthName(Month(fecha))
    VAR_FECHACTRL = VAR_FECHAINI
    VAR_FCTRLINI = VAR_FECHACTRL
    VAR_FCTRLFIN = VAR_FECHACTRL - 1
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "SELECT * FROM tc_zona_piloto_edif_inst WHERE EDIF_codigo = '" & GlEdificio & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux5.RecordCount > 0 Then
        VAR_PLANID = rs_aux5!correlativo
        VAR_BENINST = rs_aux5!beneficiario_codigo         'RESP. INSTALACION
        VAR_BENAJST = rs_aux5!beneficiario_codigo_rep    'RESP. AJUSTE
        VAR_AUX1 = rs_aux5!beneficiario_codigo_cobr   'AUX. INSTALACION
    Else
        VAR_PLANID = 1
        VAR_BENINST = "0"         'RESP. INSTALACION
        VAR_BENAJST = "0"    'RESP. AJUSTE
        VAR_AUX1 = "0"   'AUX. INSTALACIO
    End If
    Set rs_aux6 = New ADODB.Recordset
    'rs_aux6.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & "   ", db, adOpenKeyset, adLockBatchOptimistic
    rs_aux6.Open "select * from tc_tareas_crono_instalacion  ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux6.RecordCount > 0 Then
        'var_cod5 = rs_aux6.RecordCount
        rs_aux6.MoveFirst
        While Not rs_aux6.EOF
            'FECHA, MES Y DIA
            VAR_DIA = Day(VAR_FECHACTRL)
            VAR_MES = Month(VAR_FECHACTRL)
            VAR_NRODIAS = rs_aux6!NroEstimadoDias
            VAR_FCTRLINI = VAR_FCTRLFIN + 1
            VAR_FCTRLFIN = VAR_FCTRLINI + VAR_NRODIAS
            MControl = UCase(MonthName(Month(VAR_FCTRLINI)))
            VAR_IDTAREA = rs_aux6!IdTareaInst
            VAR_DESTAREA = rs_aux6!TareaDescripcion
            
            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "select * from ao_ventas_detalle where venta_codigo = " & correlv & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux7.RecordCount > 0 Then
                rs_aux7.MoveFirst
                While Not rs_aux7.EOF
                    VAR_BIEN = rs_aux7!bien_codigo
                    Select Case rs_aux7!cotiza_codigo
                        Case 1
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo1 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo1 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = rs_aux8!recorrido_codigo
                                VAR_VELOCIDAD = rs_aux8!vel_equipo_m_s
                                VAR_PASAJEROS = rs_aux8!pasajeros_descripcion
                                VAR_PARADAS = rs_aux8!trafico_num_paradas
                            End If
                        Case 2
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo2 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo2 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = rs_aux8!recorrido_codigo
                                VAR_VELOCIDAD = rs_aux8!vel_equipo_m_s
                                VAR_PASAJEROS = rs_aux8!pasajeros_descripcion
                                VAR_PARADAS = rs_aux8!trafico_num_paradas
                            End If
                        Case 3
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo3 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo3 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = rs_aux8!recorrido_codigo
                                VAR_VELOCIDAD = rs_aux8!vel_equipo_m_s
                                VAR_PASAJEROS = rs_aux8!pasajeros_descripcion
                                VAR_PARADAS = rs_aux8!trafico_num_paradas
                            End If
                        Case 4
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo4 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo4 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = rs_aux8!recorrido_codigo
                                VAR_VELOCIDAD = rs_aux8!vel_equipo_m_s
                                VAR_PASAJEROS = rs_aux8!pasajeros_descripcion
                                VAR_PARADAS = rs_aux8!trafico_num_paradas
                            End If
                        Case Else
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo1 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo1 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = rs_aux8!recorrido_codigo
                                VAR_VELOCIDAD = rs_aux8!vel_equipo_m_s
                                VAR_PASAJEROS = rs_aux8!pasajeros_descripcion
                                VAR_PARADAS = rs_aux8!trafico_num_paradas
                            End If
                    End Select
                    Select Case rs_aux6!IdTareaInst
                        Case 4
                            VAR_NRODIAS = Round((((CDbl(VAR_RECORRIDO) + 1 + 3) * 4) / 6) / 3, 0)
                        Case 10
                            VAR_NRODIAS = Round(CDbl(VAR_PARADAS) / 1.9, 0)
                        Case 15
                            VAR_NRODIAS = Round(Abs(CDbl(VAR_PASAJEROS) * CDbl(VAR_VELOCIDAD) / CDbl(VAR_PASAJEROS) * CDbl(VAR_VELOCIDAD) - 2), 0)
                            If VAR_NRODIAS = 0 Then
                                VAR_NRODIAS = 1
                            End If
                        Case Else
                            VAR_NRODIAS = VAR_NRODIAS
                    End Select
                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
                    Set rs_aux21 = New ADODB.Recordset
                    If rs_aux21.State = 1 Then rs_aux21.Close
                    rs_aux21.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_PLANID & " AND bien_codigo = '" & VAR_BIEN & "' AND horario_codigo = " & VAR_IDTAREA & " AND dia_correl = " & VAR_DIA & " ", db, adOpenKeyset, adLockBatchOptimistic
                    If rs_aux21.RecordCount > 0 Then
                        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
                        db.Execute "update to_cronograma_diario_final_INST set unidad_codigo_tec = '" & VAR_UNIDCOD & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = '" & VAR_DESTAREA & "', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & GlEdificio & "' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
                        db.Execute "update to_cronograma_diario_final_INST set bien_orden = " & VAR_IDTAREA & ", venta_codigo = " & correlv & " WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "   "
                        db.Execute "update to_cronograma_diario_final_INST set estado_activo = 'APR' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
                    Else
                        db.Execute "INSERT INTO to_cronograma_diario_final_INST (fmes_plan, dia_correl, horario_codigo, bien_orden,     bien_codigo,        unidad_codigo_tec, tec_plan_codigo,     beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha,             dia_nombre,         hora_ingreso,           hora_salida,            nro_total_horas,      observaciones,      edif_descripcion, bien_codigo1, " & _
                        " bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, nro_fojas, doc_numero, estado_activo, estado_codigo, usr_codigo,      fecha_registro, " & _
                        " hora_registro, estado_almacen, ok_almacen, doc_codigo, doc_numero_m, observaciones2, almacen_codigo, cite_certificado, estado_certificado, venta_codigo,  edif_codigo) " & _
                        " VALUES ( " & VAR_PLANID & ",      " & VAR_DIA & ",        " & VAR_IDTAREA & ",        " & VAR_IDTAREA & ", '" & VAR_BIEN & "', '" & VAR_UNIDCOD & "', " & VAR_SOL & ", '" & VAR_BENINST & "',     '" & VAR_BENAJST & "',  '" & VAR_FECHACTRL & "', '" & MControl & "', '" & VAR_FCTRLINI & "', '" & VAR_FCTRLFIN & "', " & VAR_NRODIAS & ", '" & VAR_DESTAREA & "', '" & VAR_EDIFD & "', '4211', " & _
                        " '479',        '500',          '4529',         '3113',     '0',        '0',        '0',      '0',       '0',   'NO',       '0',            '0',        '0',        'APR',      'REG',          '" & glusuario & "', '" & Date & "',  " & _
                        " '0',              'REG',      '0',          'R-115',    '0',          '',             '0',            '0',                'REG',          " & correlv & ", '" & GlEdificio & "'     )"
                        
                        ' fecha_carta, fecha_conformidad, fecha_equipo_hdm, fecha_almacen, fecha_almi, fecha_certificado, correl_prog,
                        '  '" & VAR_AUX1 & "',
                        
                    End If
                    rs_aux7.MoveNext
                Wend
            VAR_FECHACTRL = VAR_FCTRLFIN + 1
            rs_aux6.MoveNext
            End If
        Wend
    End If
End Sub

Private Sub Contabiliza_Contratos()
'    ' Contabilizacion al momento de aprobacion
'    'Base de datos
'    Dim db2 As New ADODB.Connection
'    ' Recordset
'    Dim rs_aux100 As New ADODB.Recordset
'    Dim rs_aux101 As New ADODB.Recordset
'    'Declaracion de variables
'    Dim VAR_CODTIPO As String
'    Dim VAR_EMPRESA As Integer
'    Dim VAR_TIPOCOMPID As Integer
'    Dim VAR_FECHA As Date
'    Dim VAR_MONEDAID As Integer
'    Dim VAR_TIPOCAMBIO As Double
'    Dim EntregadoA As String
'    Dim VAR_DEBEORG As Double
'    Dim VAR_HABERORG As Double
'    'Impuestos
'    Dim VAR_PorIVA As Double
'    Dim VAR_PorIT As Double
'    Dim VAR_PorITF As Double
'    'Otros valores
'    Dim VAR_ConFac As Integer
'    Dim VAR_SinFac As Integer
'    Dim VAR_Automatico As Integer
'    Dim VAR_TipoNotaId As Integer
'    Dim VAR_NotaNro As Integer
'    Dim VAR_EstadoId As Integer
'    Dim VAR_iConcurrency_id As Integer
'    Dim VAR_TipoAsientoId As Integer
'    Dim VAR_CentroCostoId As Integer
'    Dim VAR_TipoRetencionId As Integer
'    Dim VAR_TipoId As Integer
'    Dim VAR_CompDetIdOrg As Integer
'    ' Variables intermedias
'    Dim VAR_transDescripcion As String
'    ' Asignacion de valores del procedimiento Call graba_ingreso
'    VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'    VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'    ' Codigo Tipo
'    VAR_CODTIPO = "DEI"
'    ' Rubro codigo, descripcion, centro de costo id
'    Set rs_aux100 = New ADODB.Recordset
'    If rs_aux100.State = 1 Then rs_aux100.Close
'    rs_aux100.Open "SELECT trans_descripcion, rubro_codigo, CentroCostoId FROM gc_tipo_transaccion WHERE trans_codigo = '" & Ado_datos.Recordset!trans_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
'    If rs_aux100.RecordCount > 0 Then
'        VAR_transDescripcion = rs_aux100!trans_descripcion
'        VAR_PARTIDA = rs_aux100!rubro_codigo
'        VAR_CentroCostoId = rs_aux100!CentroCostoId
'        rs_aux100.Close
'    Else
'        VAR_transDescripcion = "None"
'        VAR_PARTIDA = "None"
'        VAR_CentroCostoId = "None"
'    End If
'    ' Empresa
'    If VAR_TIPOV = "G" Then
'        VAR_EMPRESA = 2
'    Else
'        VAR_EMPRESA = 1
'    End If
'    ' Fecha de venta
'    VAR_FECHA = CDate(Ado_datos.Recordset!venta_fecha)
'    ' Tipo de cambio -> BOB - USD
'    If IsNull(Ado_datos.Recordset!venta_tipo_cambio) Or (Ado_datos.Recordset!venta_tipo_cambio = 0) Or (Ado_datos.Recordset!venta_tipo_cambio = 1) Then
'        VAR_TIPOCAMBIO = GlTipoCambioOficial
'    Else
'        VAR_TIPOCAMBIO = Ado_datos.Recordset!venta_tipo_cambio
'    End If
'    'VAR_TIPOCAMBIO = Ado_datos.Recordset!venta_tipo_cambio
'    ' Tipo moneda/Debe/Haber
'    VAR_MONEDAID = 1
'    VAR_DEBEORG = VAR_BS2 'Boliviano
'    VAR_HABERORG = VAR_BS2 'Boliviano
'    ' If Ado_datos.Recordset!tipo_moneda = "USD" Then
'    '     VAR_MONEDAID = 2
'    '     VAR_DEBEORG = VAR_DOL2 'Dolar
'    '     VAR_HABERORG = VAR_DOL2 'Dolar
'    ' Else
'    '     VAR_MONEDAID = 1
'    '     VAR_DEBEORG = VAR_BS2 'Boliviano
'    '     VAR_HABERORG = VAR_BS2 'Boliviano
'    ' End If
'    ' Entregado A
'    EntregadoA = "Responsable: " & Ado_datos.Recordset!beneficiario_codigo + " - " + Ado_datos.Recordset!beneficiario_denominacion
'    ' Por Concepto
'    VAR_CONCEPTO = "Devengamiento de contrato: " & Ado_datos.Recordset!unidad_codigo_ant & " - Edificio " & Ado_datos.Recordset!edif_codigo_corto
'    Set rs_aux101 = New ADODB.Recordset
'    If rs_aux101.State = 1 Then rs_aux101.Close
'    rs_aux101.Open "select edif_descripcion from gc_edificaciones where edif_codigo = '" & VAR_PROY2 & "'  ", db, adOpenKeyset, adLockOptimistic
'    If rs_aux101.RecordCount > 0 Then
'        VAR_CONCEPTO = VAR_CONCEPTO & " " & rs_aux101!edif_descripcion
'        rs_aux101.Close
'    End If
'    If VAR_transDescripcion <> "None" Then
'        VAR_CONCEPTO = VAR_CONCEPTO & " - " & VAR_transDescripcion
'    End If
'    ' TipoCompId (Tipo comprobante id) Traspaso
'    VAR_TIPOCOMPID = 3
'    ' Impuestos
'    VAR_PorIVA = 0.13
'    VAR_PorIT = 0.03
'    VAR_PorITF = 0.0015
'    ' Otros valores
'    VAR_ConFac = 0
'    VAR_SinFac = 1
'    VAR_Automatico = 1 '0 Permite edicion, 1 no permite editar
'    VAR_TipoNotaId = Ado_datos.Recordset!solicitud_tipo
'    VAR_NotaNro = Ado_datos.Recordset!venta_codigo
'    ' Glosa general
'    VAR_GLOSA = "INGRESO POR: " & Ado_datos.Recordset!venta_descripcion & " - Nro. Venta: " & VAR_NotaNro
'    VAR_EstadoId = 11 'Libro Mayor requiere que sean de EstadoId = 10 Cerrado OR EstadoId = 11 Abierto
'    VAR_TipoAsientoId = 0 ' Operativo
'    VAR_TipoRetencionId = 0
'    VAR_TipoId = 0
'    VAR_CompDetIdOrg = 0
'    ' Creamos conexion unica para CONDOBO
'    db2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CONDOBO;Data Source=SSOFIA"
'    ' Procedimiento almacenado
'    db2.Execute ("EXEC fp_contabiliza_ingresos '" & VAR_CODTIPO & "', '" & VAR_PARTIDA & "', " & VAR_EMPRESA & ", " & VAR_DPTO & ", " & VAR_TIPOCOMPID & ", '" & VAR_FECHA & "', " & VAR_MONEDAID & ", '" & VAR_TIPOCAMBIO & "', '" & VAR_DEBEORG & "', '" & VAR_HABERORG & "', '" & EntregadoA & "', '" & VAR_CONCEPTO & "', '" & VAR_PorIVA & "', '" & VAR_PorIT & "', '" & VAR_PorITF & "', " & VAR_ConFac & ", " & VAR_SinFac & ", " & VAR_Automatico & ", '" & VAR_GLOSA & "', " & VAR_TipoNotaId & ", " & VAR_NotaNro & ", " & VAR_EstadoId & ", '" & glusuario & "', " & VAR_TipoAsientoId & ", " & VAR_CentroCostoId & ", " & VAR_TipoRetencionId & ", " & VAR_TipoId & ", " & VAR_CompDetIdOrg & ", '" & VAR_PROY2 & "'")
'    db2.Close
End Sub

Private Sub BtnAprobar1_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    If Ado_datos.Recordset.RecordCount > 0 Then
        If CDate(Format(DTPfechasol.Value, "dd/mm/yyyy")) = "01/01/1900" Or DTPfechasol.Value = "" Then
            MsgBox "Debe registrar la Fecha de Venta !! , Verifique y vuelva a Intentar ...", vbExclamation, "Atención"
            VAR_VAL = "ERR"
            Exit Sub
        End If

        correlv = Ado_datos.Recordset!venta_codigo
        GlEdificio = Ado_datos.Recordset!edif_codigo
        'VALIDA EDIFICIO Y EQUIPOS
        Set rs_aux10 = New ADODB.Recordset     'Proyecto de Edificación
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & dtc_codigo3.Text & "' and estado_codigo = 'APR' ", db, adOpenStatic
        If rs_aux10.RecordCount = 0 Then
            'Si Faltarian Aprobar
            MsgBox "No se puede APROBAR, verifique los datos del Edificio si estan correctos y si está Aprobado, luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        
        Set rs_aux11 = New ADODB.Recordset     'Equipos de Venta_Detalle
        If rs_aux11.State = 1 Then rs_aux11.Close
        rs_aux11.Open "Select * from mv_bienes_vs_venta_det WHERE venta_codigo = " & correlv & "  ", db, adOpenStatic
        If rs_aux11.RecordCount > 0 Then
            'Si Faltarian Aprobar
            MsgBox "No se puede APROBAR, verifique los datos de los EQUIPOS y si estos están Aprobados, luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        
        Set rs_aux12 = New ADODB.Recordset     'Partidas de Venta_Detalle
        If rs_aux12.State = 1 Then rs_aux12.Close
        rs_aux12.Open "Select * from ao_ventas_detalle WHERE venta_codigo = " & correlv & " and par_codigo=''  ", db, adOpenStatic
        If rs_aux12.RecordCount > 0 Then
            'Si Faltarian Partida
            MsgBox "No se puede APROBAR, verifique los datos de Detalle de Bienes , luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
        'rs_aux18
        Set rs_aux18 = New ADODB.Recordset     'Alcance del Contrato
        If rs_aux18.State = 1 Then rs_aux18.Close
        rs_aux18.Open "Select * from ao_ventas_alcance WHERE venta_codigo = " & correlv & "  ", db, adOpenStatic
        If rs_aux18.RecordCount < 6 Then
            'Si Faltarian Partida
            MsgBox "No se puede APROBAR, verifique los datos del Alcance del Contrato , luego vuelva a intentar ...", , "Atención"
            Exit Sub
        End If
       'If IsNull(Ado_datos.Recordset("venta_tipo")) Or Ado_datos.Recordset("venta_tipo") = "" Or (Ado_datos.Recordset("venta_monto_total_bs") = 0) Or (Ado_datos.Recordset!estado_alcance = "N") Or (Ado_datos.Recordset!unidad_codigo_ant = "") Or IsNull(Ado_datos.Recordset!unidad_codigo_ant) Then
       '     MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
       '     Exit Sub
       'Else
        If Ado_datos.Recordset!estado_codigo_verif = "REG" Then
           sino = MsgBox("Esta seguro de Verificar el registro?", vbYesNo, "Confirmando")
           If sino = vbYes Then
               ' APRUEBA ao_ventas_cabecera
               Ado_datos.Recordset!estado_codigo_verif = "APR"
               Ado_datos.Recordset.Update
               VAR_ZONA = Ado_datos.Recordset!zpiloto_codigo
               If VAR_ZONA = "" Or IsNull(VAR_ZONA) Then
                    Set rs_datos6 = New ADODB.Recordset
                    If rs_datos6.State = 1 Then rs_datos6.Close
                    rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & GlEdificio & "'    ", db, adOpenStatic
                    If rs_datos6.RecordCount > 0 Then
                        VAR_ZONA = rs_datos6!zpiloto_codigo
                    Else
                        MsgBox "El Edificio de este contrato no tiene una ZONA PILOTO asignada, Consulte con Area Técnica ...", , "Atención"
                    End If
               End If
               'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo_verif = 'APR' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
               ' Asigna Deudor
               db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
               
               'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux2.RecordCount > 0 Then
                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                    Ado_datos.Recordset!doc_numero = rs_aux2!correl_doc
                    'Txt_campo1.Caption = rs_aux2!correl_doc
                    rs_aux2.Update
                End If
                'Guarda Documento RESPLADO
                'If IsNull(Ado_datos.Recordset!doc_codigo) Or IsNull(Ado_datos.Recordset!doc_numero) Then
                If Ado_datos.Recordset!unidad_codigo = "DNMOD" Then
                  ' Validar consistencia de datos.
                  VAR_ARCH = "MOD_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
                Else
                  VAR_ARCH = "COM_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
                End If
                db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
                db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
                
            End If
        Else
            MsgBox "NO se puede VERIFICAR, el registro ya fue Verificado o Anulado...!! Revise y vuelva a intentar... ", vbExclamation, "Atención!"
        End If
    Else
        MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    End If
End Sub

Private Sub BtnAprobar2_Click()
 If IsNull(Ado_datos16.Recordset("cobranza_observaciones")) Or (Ado_datos16.Recordset("cobranza_programada_bs") = 0) Then
    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
    Exit Sub
 Else
    If Ado_datos.Recordset!estado_codigo_verif = "REG" Then         '("estado_codigo")
        MsgBox "No se puede APROBAR el registro (Cronograma), previamente debe APROBAR la Venta (Cabecera) y vuelva a intentar ...", , "Atención"
        Exit Sub
    End If
    NumComp = Ado_datos.Recordset!venta_codigo
    VAR_COBRANZA = Ado_datos16.Recordset!cobranza_prog_codigo
    If Ado_datos16.Recordset!estado_codigo = "REG" Then
       nroventa = Ado_datos16.Recordset!venta_codigo
       VAR_DOCFAC = Ado_datos16.Recordset!doc_codigo_fac
       db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & nroventa & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
       VAR_COBR0 = IIf(IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp), "3361040", Ado_datos16.Recordset!beneficiario_codigo_resp)
       VAR_COBR0 = IIf(Ado_datos16.Recordset!beneficiario_codigo_resp = "0", "3361040", Ado_datos16.Recordset!beneficiario_codigo_resp)

       If (VAR_DOCFAC = "R-101" Or VAR_DOCFAC = "R-100") Then
          sino = MsgBox("Realizarás la solicitud de VARIAS cuotas en UNA sola FACTURA ? ", vbYesNo, "Confirmando")
          If sino = vbYes Then             'VARIAS CUOTAS PARA UNA SOLA FACTURA
                tw_ventas_cuotas_vs_fac.Show vbModal
          Else                             'UNA CUOTA PARA UNA FACTURA
            'GRABA CABECERA DE FACTURACION NUEVA (ao_ventas_cobranza_fac)   'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            'SE GENERAN CON LA FACTURA (dosifica_autorizacion, nro_factura, fecha_fac, codigo_control, archivo_foto, depto_codigo, Gestion, mes, edif_codigo_corto)
            db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, venta_codigo, doc_codigo_fac,              beneficiario_codigo_fac,                                beneficiario_nit,           glosa_Descripcion,                                  beneficiario_RazonSocial, nro_dui,      total_bs,                                       total_dol,                                      cambio_oficial, " & _
                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal,                    factura_87_bs,                                                      factura_87_dol,                                                 debito_fiscal_13_bs,                                                debito_fiscal_13_dol,                                               literal, " & _
                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, correl_contab, estado_fac, estado_codigo_fac, estado_codigo,  " & _
                        " usr_codigo, fecha_registro, edif_codigo_corto, edif_codigo, codigo_empresa ) " & _
                " VALUES ('" & glGestion & "',  " & nroventa & ", '" & VAR_DOCFAC & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & dtc_codigo2A.Text & "', '" & Ado_datos16.Recordset!cobranza_concepto_plazo & "', '" & dtc_desc2A.Text & "',  '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ",  " & Ado_datos16.Recordset!cobranza_programada_dol & ",  " & GlTipoCambioOficial & ",  " & _
                        " '0',          '0',                    '0',            '0',            '0',    " & Ado_datos16.Recordset!cobranza_total_bs & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.87, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_bs * 0.13, 2) & ", " & Round(Ado_datos16.Recordset!cobranza_total_dol * 0.13, 2) & ", '" & Ado_datos16.Recordset!Literal & "',  " & _
                        " 'ADM',        'R-103',        '0',        'N',            'BOB',      'NN',           'NN',        '0',            'REG',      'REG',          'REG',  " & _
                        " '" & glusuario & "', '" & CDate(Date) & "', " & Ado_datos.Recordset!edif_codigo_corto & ", '" & Ado_datos.Recordset!edif_codigo & "', " & Ado_datos.Recordset!codigo_empresa & "  ) "
                        
            Set rs_aux20 = New ADODB.Recordset
            If rs_aux20.State = 1 Then rs_aux20.Close
            rs_aux20.Open "Select max(IdFactura) as Codigo3 from ao_ventas_cobranza_fac  ", db, adOpenKeyset, adLockOptimistic
            If IsNull(rs_aux20!codigo3) Then
               VAR_IDFAC = 1
            Else
               VAR_IDFAC = rs_aux20!codigo3
            End If
            'GRABA CABECERA DE LA FACTURA (QR)
            db.Execute "INSERT INTO ao_ventas_cobranza_fac_QR (IdFactura, archivo_foto_cargado, estado_codigo, usr_codigo, fecha_registro ) " & _
                " VALUES ('" & VAR_IDFAC & "',  'N',            'REG',   '" & glusuario & "', '" & CDate(Date) & "' ) "

            'Actualiza CORREO ELECTRONICO (a los que NO tienen EMail)
            db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.beneficiario_email  = gc_beneficiario.beneficiario_email FROM ao_ventas_cobranza_fac INNER JOIN gc_beneficiario ON ao_ventas_cobranza_fac.beneficiario_codigo_fac = gc_beneficiario.beneficiario_codigo where ao_ventas_cobranza_fac.beneficiario_email Is Null "
          End If
       Else
            VAR_IDFAC = 0
       End If
        'GRABA DETALLE DE FACTURACION NUEVA (ao_ventas_cobranza)
        db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo, venta_codigo,                                    beneficiario_codigo,                                    beneficiario_codigo_fac,                            beneficiario_codigo_resp,                               cobranza_programada_bs,                                 cobranza_programada_dol,                                cobranza_solicitado_bs,                                  cobranza_solicitado_dol,                 cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs,         cobranza_total_dol,                                     Literal,    cobranza_fecha_prog,                              cobranza_fecha_cobro, cobranza_observaciones, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo,  " & _
        " estado_codigo, usr_codigo, fecha_registro, cobranza_fecha_sol, estado_codigo_sol, estado_codigo_fac, venta_codigo_new) " & _
        " VALUES ('" & glGestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & nroventa & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo_resp & "', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '0', '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_fecha_cobro & "', '" & Ado_datos16.Recordset!cobranza_concepto_plazo & "', 'FIN', 'FIN-02', 'FIN-02-02', 'ADM', 'R-105', '0', '" & VAR_DOCFAC & "', '0', '0', '3.1.2',  " & _
        " 'REG', '" & glusuario & "', '" & Date & "', '" & Date & "', 'APR', 'REG', " & VAR_IDFAC & " )"

        ' APRUEBA ao_ventas_cobranza_prog
        'db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
        db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR', fecha_registro= '" & Ado_datos16.Recordset!fecha_registro & "' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
        ' Actualiza CODIGO_COBRNAZA en el cronogrma
        db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_codigo = ao_ventas_cobranza.cobranza_codigo from ao_ventas_cobranza_prog INNER JOIN ao_ventas_cobranza " & _
        " ON ao_ventas_cobranza_prog.venta_codigo = ao_ventas_cobranza.venta_codigo and ao_ventas_cobranza_prog.cobranza_prog_codigo = ao_ventas_cobranza.cobranza_prog_codigo WHERE (ao_ventas_cobranza_prog.venta_codigo = " & nroventa & " and ao_ventas_cobranza_prog.cobranza_prog_codigo=" & Ado_datos16.Recordset!cobranza_prog_codigo & " )"

        db.Execute "update ao_ventas_cobranza_prog SET Gestion = YEAR(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "

        db.Execute "update ao_ventas_cobranza_prog SET cobranza_mes = MONTH(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
        MsgBox "Se APROBOBO la Cuota y se Envió satisfactoriamente la Solicitud de FACTURA ...", , "Atención"
        Call ABRIR_TABLA_DET
        If (DtgCobro.SelBookmarks.Count <> 0) Then
            DtgCobro.SelBookmarks.Remove 0
        End If
        If Ado_datos16.Recordset.RecordCount > 0 Then
            rs_datos16.Find "cobranza_prog_codigo = " & VAR_COBRANZA & "   ", , , 1
            DtgCobro.SelBookmarks.Add (rs_datos16.Bookmark)
        Else
           rs_datos16.MoveLast
        End If
    
    
'       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'            db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & Ado_datos.Recordset!venta_codigo & " Where gc_documentos_respaldo.doc_codigo = '" & Ado_datos16.Recordset!doc_codigo & "' "
'            'If Ado_datos.Recordset!unidad_codigo = "DVTA" Then
'            '    VAR_COBR0 = "3361040"
'            'Else
'                VAR_COBR0 = IIf(IsNull(Ado_datos16.Recordset!beneficiario_codigo_resp), "3361040", Ado_datos16.Recordset!beneficiario_codigo_resp)
'                VAR_COBR0 = IIf(Ado_datos16.Recordset!beneficiario_codigo_resp = "0", "3361040", Ado_datos16.Recordset!beneficiario_codigo_resp)
'            'End If
'            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo,                                  venta_codigo,                                   beneficiario_codigo,                                beneficiario_codigo_fac,                            beneficiario_codigo_resp, " & _
'            " cobranza_programada_bs,                                   cobranza_programada_dol,                     cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs,                                      cobranza_total_dol,                                 Literal,                                        cobranza_fecha_prog,                                cobranza_fecha_cobro,                               cobranza_observaciones,         proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac,                      cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo, estado_codigo, usr_codigo,     fecha_registro, estado_codigo_sol, estado_codigo_fac, es_liquidacion,                                cobranza_fecha_sol) " & _
'            "VALUES ('" & Ado_datos16.Recordset!ges_gestion & "', " & Ado_datos16.Recordset!cobranza_prog_codigo & ", " & Ado_datos16.Recordset!venta_codigo & ", '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & Ado_datos16.Recordset!beneficiario_codigo & "', '" & VAR_COBR0 & "',  " & _
'            " " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ",  '0',            '0',                   '0',                  '0', " & Ado_datos16.Recordset!cobranza_programada_bs & ", " & Ado_datos16.Recordset!cobranza_programada_dol & ", '" & Ado_datos16.Recordset!Literal & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "', '" & Ado_datos16.Recordset!cobranza_observaciones & "', 'FIN',    'FIN-02',       'FIN-02-02',    'ADM',          'R-393', '0', '" & Ado_datos16.Recordset!doc_codigo_fac & "', '0',                  '0',                    '3.1.2',        'REG', '" & glusuario & "', '" & Date & "',     'APR',         'REG',           '" & Ado_datos16.Recordset!es_liquidacion & "', '" & Ado_datos16.Recordset!cobranza_fecha_prog & "') "
'
'            ' APRUEBA ao_ventas_cobranza_prog
'            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where  venta_codigo = " & Ado_datos.Recordset!venta_codigo & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "     'ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' And
'            ' Actualiza CODIGO_COBRNAZA en el cronogrma
'            db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_codigo = ao_ventas_cobranza.cobranza_codigo from ao_ventas_cobranza_prog INNER JOIN ao_ventas_cobranza " & _
'            " ON ao_ventas_cobranza_prog.venta_codigo = ao_ventas_cobranza.venta_codigo and ao_ventas_cobranza_prog.cobranza_prog_codigo = ao_ventas_cobranza.cobranza_prog_codigo WHERE (ao_ventas_cobranza_prog.venta_codigo = " & nroventa & " and ao_ventas_cobranza_prog.cobranza_prog_codigo=" & Ado_datos16.Recordset!cobranza_prog_codigo & " )"
'
'            Call ABRIR_TABLA_DET
''            Ado_datos16.Refresh
'       End If
    End If
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
  FrmCobranza.Visible = True
  FrmAlcance.Visible = True
  Fra_Total.Visible = True
  dg_datos.Visible = True
  FrmABMDet.Visible = True
  FrmABMDet1.Visible = True
  FrmABMDet2.Visible = True
'  TxtCobrado.Visible = False
'  Label7.Visible = False
'  Cmd_Cliente.Visible = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True
  SSTab1.TabEnabled(2) = True
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnCancelar2_Click()
    FraZona.Visible = False
End Sub

Private Sub BtnCancelarBen_Click()
    frm_benef.Visible = False
End Sub

Private Sub BtnEliminar_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
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
    FrmCobranza.Visible = True
    FrmAlcance.Visible = True
    Fra_Total.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FrmABMDet2.Visible = True
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
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

Private Sub BtnGrabar1_Click()
    NumComp = Ado_datos.Recordset!venta_codigo
    'VALIDACION FECHAS
    Set rs_aux19 = New ADODB.Recordset
    If rs_aux19.State = 1 Then rs_aux19.Close
    rs_aux19.Open "Select * from gc_tipo_solicitud where solicitud_num = '90' order by ORDEN ", db, adOpenStatic
    If rs_aux19.RecordCount > 0 Then
        rs_aux19.MoveFirst
        While Not rs_aux19.EOF
            'UPDATE  ao_ventas_alcance SET venta_tiempo_dias = (SELECT DATEDIFF(day, fecha_inicio_alcance, fecha_fin_alcance) FROM ao_ventas_alcance WHERE venta_codigo = '8204' and solicitud_tipo='3')
            ' WHERE venta_codigo = '8204' and solicitud_tipo='3'
            
            db.Execute "UPDATE  ao_ventas_alcance SET venta_tiempo_dias = (SELECT DATEDIFF(day, fecha_inicio_alcance, fecha_fin_alcance) FROM ao_ventas_alcance WHERE venta_codigo =  " & NumComp & " and solicitud_tipo= " & rs_aux19!solicitud_tipo & ") WHERE venta_codigo = " & NumComp & " and solicitud_tipo = " & rs_aux19!solicitud_tipo & " "
            
            rs_aux19.MoveNext
        Wend
    End If
    'rs_aux19.Requery
    Call ABRIR_TABLA_DET
    'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)
    'If Val(Txt_descripcion) < 0 Then
    '    MsgBox "La Fecha de Inicio NO puede ser MAYOR a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
    '    DTPfechaFin.SetFocus
    'End If

    DtgAlcance.AllowUpdate = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = True
    FrmABMDet2.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FraGrabarCancelar1.Visible = False
End Sub

Private Sub BtnGrabar2_Click()
    db.Execute "update ao_ventas_cabecera SET zpiloto_codigo = " & dtc_codigo7.Text & " WHERE venta_codigo = " & Ado_datos.Recordset!venta_codigo & "    "
    VAR_ZONA = dtc_codigo7.Text
    FraZona.Visible = False
End Sub

Private Sub BtnGrabarBen_Click()
    db.Execute "UPDATE gc_beneficiario set beneficiario_email = '" & TxtEmail.Text & "', beneficiario_telefono_Cel = '" & TxtCelular.Text & "' where beneficiario_codigo = '" & dtc_benef2A.Text & "' "
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    rs_datos9.Open "Select * from gc_beneficiario WHERE beneficiario_codigo = '" & dtc_benef2A.Text & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_codigo2A.BoundText = dtc_benef2A.BoundText
    dtc_desc2A.BoundText = dtc_benef2A.BoundText
    dtc_email2A.BoundText = dtc_benef2A.BoundText
    frm_benef.Visible = False
End Sub

Private Sub BtnImprimir_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        'fra_reportes.Visible = True
        
        Dim iResult As Variant, i%, Y%
        Dim co As New ADODB.Command

    '    Dim rs As New ADODB.Recordset
    '    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.Ado_datos.Recordset!ges_gestion & "' and " & _
    '            "correl_venta=" & Me.Ado_datos.Recordset!correl_venta & " and venta_codigo=" & Me.Ado_datos.Recordset!venta_codigo, db, adOpenStatic, adLockReadOnly
    '    i = 1
    '    y = 1
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
          Case "DVTA", "DCOMS", "DCOMB", "DCOMC"
              var_titulo = "Módulo Comercial"
        End Select

        CryV01.ReportFileName = App.Path & "\reportes\ventas\ar_lista_de_ventas.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CryV01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
        'CryV01.StoredProcParam(2) = Me.Ado_datos.Recordset!venta_codigo
        CryV01.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        
        CryV01.Formulas(1) = "titulo = '" & var_titulo & "' "
        CryV01.Formulas(2) = "subtitulo = '" & lbl_titulo.Caption & "' "
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
        FrmCobranza.Visible = False
        FrmAlcance.Visible = False
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
        If IsNull(DTPfechasol) Then
            DTPfechasol.Value = Date
        End If
        FrmABMDet.Visible = False
        FrmABMDet1.Visible = False
        FrmABMDet2.Visible = False
    
        swgrabar = 0
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
    Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
    End If
End Sub

Private Sub BtnSalir_Click()
    sino = MsgBox("Esta Seguro de Cerrar la Ventana?", vbQuestion + vbYesNo, "Confirmando...")
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

Private Sub BtnModDetalle1_Click()
    If Ado_datos6.Recordset.RecordCount > 0 Then
        'nro_licitacion = Ado_datos.Recordset!venta_codigo
        'mw_ventas_alcance.Show vbModal
        NumComp = Ado_datos.Recordset!venta_codigo
        DtgAlcance.Enabled = True
        DtgAlcance.AllowUpdate = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = False
        FraNavega.Enabled = False
        FrmABMDet2.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet1.Visible = False
        FraGrabarCancelar1.Visible = True
        
    '    NumComp = Ado_datos.Recordset!venta_codigo
    '    frm_ao_ventas_alcance.Show vbModal
    '    Call ABRIR_TABLA_DET
    Else
        MsgBox "No se puede Modificar, debe existir por lo menos un registro, vuelva a intentar...", , "Atención ..."
    End If
End Sub

Private Sub BtnVer_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        NumComp = Ado_datos.Recordset!venta_codigo
        Cod_Comp = Ado_datos.Recordset!solicitud_tipo
        tw_ventas_adenda.Show vbModal
    End If

End Sub

Private Sub BtnVer3_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
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

Private Sub Cmd_Cliente_Click()
    glPersNew = "P"
    frmBeneficiario.Show 'vbModal
End Sub

Private Sub CmdCancelaCobro_Click()
    nroventa = Ado_datos.Recordset!venta_codigo
    Ado_datos16.Recordset.Cancel
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  
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
        rs_datos.Find "venta_codigo = " & nroventa & "   ", , , 1
        dg_datos.SelBookmarks.Add (rs_datos.Bookmark)
     Else
     'VAR_SW = ""
        rs_datos.MoveLast
  End If
  
'  If Ado_datos.Recordset("estado_codigo") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    FrmAlcance.Visible = True
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FrmABMDet2.Visible = True
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroVenta.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  'Call cerea
  swnuevo = 0
  'cmdElige.Enabled = False
  'marca1 = Ado_datos.Recordset.Bookmark
'  If Ado_datos.Recordset("estado_codigo") = "REG" Then
'    Call OptFilGral1_Click
'  Else
'    Call OptFilGral2_Click
'  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    FrmDetalle.Enabled = True
    FrmAlcance.Visible = True
    FrmCobranza.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FrmABMDet2.Visible = True
  'ado_datos14.Refresh
  'Ado_datos.Recordset.Move marca1 - 1
End Sub

Private Sub BtnAnlDetalle2_Click()
 'If Ado_datos16.Recordset!estado_codigo = "REG" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
    'If Ado_datos16.Recordset.RecordCount > 0 Then
    If Ado_datos16.Recordset!estado_codigo = "REG" Then
      db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset!venta_codigo & "  And ao_ventas_cobranza_prog.cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "      'Ado_datos16.Recordset!cobranza_codigo
      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.estado_codigo = 'ANL' Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & Ado_datos16.Recordset("cobranza_codigo") & " "
      'db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_deuda_bs = '0', ao_ventas_cobranza_prog.cobranza_deuda_dol = '0'  Where ao_ventas_cobranza_prog.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cobranza_prog.venta_codigo = " & Ado_datos.Recordset("venta_codigo") & "  And ao_ventas_cobranza_prog.cobranza_codigo = " & ado_datos16.Recordset("cobranza_codigo") & " "
     'End If
     'ado_ventas_COBRANZAS.Recordset.Delete
     'ado_ventas_COBRANZAS.Recordset.Update
     'ado_ventas_COBRANZAS.Requery
     'ado_ventas_COBRANZAS.Refresh
     ''cerea
     'ado_ventas_COBRANZAS.Refresh
   Else
    MsgBox "El item " + Ado_datos16.Recordset!cobranza_prog_codigo + " del " + FrmCobranza.Caption + " Ya fue APROBADO o ANULADO !! ", vbExclamation, "Atención!"
   End If
   End If
  'Else
  '  MsgBox "El item " + Ado_datos16.Recordset!cobranza_prog_codigo + " del " + FrmCobranza.Caption + " Ya fue APROBADO o ANULADO !! ", vbExclamation, "Atención!"
  'End If
End Sub

Private Sub BtnModDetalle2_Click()
  'If Ado_datos.Recordset!venta_tipo <> "E" And Ado_datos16.Recordset!estado_codigo = "REG" Then
  If Ado_datos16.Recordset!estado_codigo = "REG" And (Ado_datos.Recordset!venta_tipo = "E" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "G" Or Ado_datos.Recordset!venta_tipo = "L") Then
    FraNavega.Enabled = False
    fraOpciones.Enabled = False
    FrmDetalle.Visible = False
    FrmCobranza.Visible = False
    FrmAlcance.Visible = False
    swnuevo = 2
    TxtCobrador.Visible = False
    'marca1 = ado_datos14.Recordset.BookMark
    SSTab1.Tab = 2
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    FrmCobros.Visible = True
    FrmCobros.Enabled = True
    FrmABMDet.Visible = False
    FrmABMDet1.Visible = False
    FrmABMDet2.Visible = False
    DTPFechaProg.Visible = True
    DTPFechaCobro.Visible = False
    If Ado_datos16.Recordset!doc_codigo_fac = "R-101" Then
        cmd_fac.Text = "FACTURA"
    Else
        cmd_fac.Text = "ORDEN DE COBRO"
    End If
'        Txt_parche.Visible = True       '&H80000005&
        'dtc_desc2A.BackColor = &H80000005
    'End If
    VAR_MBS2 = Ado_datos16.Recordset!cobranza_programada_bs
    cmd_fac.SetFocus
  Else
    MsgBox "La Venta NO tiene saldo para cobrar o el Registro ya fue Aprobado !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  'If Ado_datos.Recordset!venta_tipo = "C" And Ado_datos.Recordset!estado_codigo = "APR" Then
  If Ado_datos.Recordset!venta_tipo = "C" Or Ado_datos.Recordset!venta_tipo = "V" Or Ado_datos.Recordset!venta_tipo = "G" Or Ado_datos.Recordset!venta_tipo = "L" Then
  'If Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then
  '      MsgBox "Ya se registró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
  '  End If
    If (glusuario = "CPAREDES" Or glusuario = "ADMIN" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "MARTEAGA" Or glusuario = "RGIL" Or glusuario = "GMORA" Or glusuario = "CSALINAS") Or Ado_datos.Recordset!venta_saldo_p_cobrar_bs > 0 Then            'Or glusuario = "ADMIN"
    'If Ado_datos.Recordset!venta_monto_total_bs - Ado_datos.Recordset!venta_monto_cobrado_bs > 0 Then
        swnuevo = 1
        SSTab1.Tab = 2
        SSTab1.TabEnabled(2) = True
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        FrmCobros.Visible = True
        FrmCobros.Enabled = True
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FrmDetalle.Visible = False
        FrmCobranza.Visible = False
        FrmAlcance.Visible = False
        FrmABMDet.Visible = False
        FrmABMDet1.Visible = False
        FrmABMDet2.Visible = False
        TxtCobrador.Visible = False
        nroventa = Ado_datos.Recordset!venta_codigo
        Ado_datos16.Recordset.AddNew
        dtc_codigo2A.Text = dtc_codigo2.Text
        dtc_desc2A.Text = dtc_desc2.Text
        TxtMonto.SetFocus
        DTPFechaProg.Visible = True
        DTPFechaCobro.Visible = False
        Lbl_nombre_fac.Caption = "Cliente :"
        lbl_fechas.Caption = "Fecha Programada de la Cobranza"
        If dtc_codigo11.Text = "V" Then
            cmd_fac.Text = "FACTURA"
        Else
            cmd_fac.Text = "RECIBO"
        End If
'        Txt_parche.Visible = True
        'Ado_datos.Recordset.Move marca1 - 1
    Else
    
        MsgBox "Ya se registró el total de la deuda, Verifique por favor !! ", vbExclamation, "Atención!"
    End If
  Else
    MsgBox "La Venta (Acumulada del Cronograma) NO tiene saldo para cobrar, Verifique por favor !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub BtnDesAprobar_Click()
'  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
'  If sino = vbYes Then
'    Dim rstdestino As New ADODB.Recordset
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from ao_ventas_cabecera where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and correl_venta = " & Ado_datos.Recordset("correl_venta") & " and venta_codigo = " & Ado_datos.Recordset("venta_codigo") & " ", db, adOpenDynamic, adLockOptimistic
'    If Not rstdestino.BOF Then rstdestino.MoveFirst
'    If Not rstdestino.BOF And Not rstdestino.EOF Then
'      rstdestino("estado_codigo") = "REG"
'      rstdestino.Update
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    marca1 = Ado_datos.Recordset.Bookmark
'    Call OptFilGral1_Click
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
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
  With ALFrmMateriales
        .ALPrincipal
        If .QResp Then
            TxtCodigo.Text = .QCodigo
            txtDesc.Text = .QItem
        End If
    End With
    Txtcant_alm = 0
    Cant_Alm = 0
    DE.dbo_albSacaDetalleMaterial Mid(TxtCodigo, 3, 12), descri_bien, Cant_Alm
    Txtcant_alm = Cant_Alm
    If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
    End If
End Sub

Private Sub Contabiliza_venta()
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
'      rs_aux2("unidad_codigo") = VAR_UNIDCOD       'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux1") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux2") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux3") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux1") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux2") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux3") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux1") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux2") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("D_Cta_Aux3") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux1") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux2") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
'                rstdestino2("H_Cta_Aux3") = VAR_UNIDCOD        'Ado_datos.Recordset("unidad_codigo")
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
    Select Case VAR_UNIDCOD
        Case "DNAJS", "DNEME", "DNINS", "DNMAN", "DNMOD", "DNREP"
            VAR_PROY = 12
        Case "GCOM"
            VAR_PROY = 17
        Case "DVTA", "DCOMB", "DCOMS", "DCOMC"
            VAR_PROY = 18
            
    End Select
        
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    'SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & Ado_datos.Recordset!edif_codigo & "' "
    SQL_FOR = "select * from fo_proyectos_ejecucion where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & VAR_PROY2 & "' "
    rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    If rs_aux1.RecordCount > 0 Then
        db.Execute "update fo_proyectos_ejecucion set pro_codigo_det_descripcion = '" & RTrim(dtc_desc3.Text) & "' Where pro_codigo = " & VAR_PROY & " AND pro_codigo_det = '" & VAR_PROY2 & "' "
    Else
        db.Execute "INSERT INTO fo_proyectos_ejecucion (pro_codigo, pro_codigo_det, pro_codigo_det_descripcion, unidad_codigo, ges_gestion, estado_codigo, usr_codigo, fecha_registro) " & _
           "VALUES (" & VAR_PROY & ", '" & VAR_PROY2 & "', '" & RTrim(dtc_desc3.Text) & "', '" & VAR_UNIDCOD & "', " & Ado_datos.Recordset!ges_gestion & ", 'APR', '" & glusuario & "', '" & Date & "')"
    End If
End Sub

Private Sub graba_ingreso()
    '======= Ini grabado de datos
   'swgraba = 0
   'Call valida
   'VAR_UNIDCOD = Ado_datos.Recordset!unidad_codigo
   Select Case VAR_UNIDCOD
        Case "DVTA", "DCOMB", "DCOMS", "DCOMC"             'INI COMERCIAL
            VAR_ORG = "111"
            VAR_TIPOS = 3
            VAR_PARTIDA = "11310"
        Case "COMEX"            'INI COMEX
            VAR_ORG = "111"
            VAR_PARTIDA = "11310"
        Case "DNINS"            'INI INSTALACIONES
            VAR_ORG = "111"
            VAR_TIPOS = 4
            VAR_PARTIDA = "11350"
        Case "DNAJS"            'INI AJUSTE
            VAR_ORG = "113"
            VAR_TIPOS = 5
            VAR_PARTIDA = "11350"
        Case "DNMAN"            'INI MANTENIMIENTO
            VAR_ORG = "112"
            VAR_TIPOS = 6
            VAR_PARTIDA = "11320"
        Case "DNREP"            'INI REPARACIONES
            VAR_ORG = "113"
            VAR_TIPOS = 7
            VAR_PARTIDA = "11330"
        Case "DNMOD"            'INI MODERNIZACION
            VAR_ORG = "114"
            VAR_TIPOS = 9
            VAR_PARTIDA = "11340"
        Case "DNEME"            'INI EMERGENCIAS
            VAR_ORG = "113"
            VAR_TIPOS = 8
            VAR_PARTIDA = "11330"
        Case Else               'INI COMPRAS
            VAR_ORG = "311"
            VAR_TIPOS = 10
            VAR_PARTIDA = "11330"
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
         rstdestino("proceso_codigo") = "FIN"
         rstdestino("subproceso_codigo") = "FIN-01"
         rstdestino("etapa_codigo") = "FIN-01-01"
         rstdestino("clasif_codigo") = "ADM"
         rstdestino("doc_codigo") = "R-110"
         rstdestino("doc_numero") = correlativo1
         rstdestino("unidad_codigo") = VAR_UNIDCOD     'Ado_datos.Recordset("unidad_codigo")
         rstdestino("solicitud_codigo") = VAR_SOL   'Ado_datos.Recordset("solicitud_codigo")
         rstdestino("solicitud_tipo") = VAR_TIPOS   '"3"

         rstdestino("beneficiario_codigo") = VAR_BENEF  ' Ado_datos.Recordset("beneficiario_codigo")
'         VAR_BENEF = Ado_datos.Recordset("beneficiario_codigo")
         rstdestino("fecha_ingreso") = Date
         rstdestino("tipo_cambio") = GlTipoCambioOficial 'GlTipoCambioMercado
         rstdestino("tipo_moneda") = "BOB"
         VAR_MONEDA = "BOB"
         'VAR_GLOSA = Ado_datos.Recordset("venta_descripcion")
         rstdestino("ingreso_concepto") = "INGRESO POR: " + VAR_GLOSA
         If Ado_datos.Recordset("venta_tipo") = "E" Then
            VAR_CODTIPO = "DYR"
         Else
            If VAR_TIPOV = "V" Then
                VAR_CODTIPO = "DEI"
            Else
                VAR_CODTIPO = "DEI"
            End If
            'AQUUIIII       DEY         WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
         End If
         'CAMBIAR DEI O REC
         rstdestino("Codigo_tipo") = VAR_CODTIPO
         rstdestino("tipo_comp") = VAR_CODTIPO
         'CAMBIAR DEI O REC
         'INI FTE
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
         'FIN FTE
         'CAMBIAR RUBROS
         rstdestino("rubro_codigo") = VAR_PARTIDA       '"11200"
         'VAR_PARTIDA = "11200"
         'CAMBIAR RUBROS
         rstdestino("cheque_o_trf") = ""
         rstdestino("Bco_codigo") = "NN"
         'CAMBIAR CTA
         rstdestino("cta_codigo") = "NN"
         VAR_CTA = "NN"
         'CAMBIAR CTA
         rstdestino("numero_documento") = "0"
         rstdestino("unidad_codigo_ant") = VAR_CITE     ' Ado_datos.Recordset("unidad_codigo_ant")
         'VAR_CITE = Ado_datos.Recordset("unidad_codigo_ant")
         rstdestino("monto_dolares") = VAR_DOL2 'Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         'VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
         rstdestino("monto_bolivianos") = VAR_BS2   'Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         'VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
         rstdestino("monto_recaudado_dolares") = 0
         rstdestino("monto_recaudado_bolivianos") = 0
         rstdestino("convenio_codigo") = "NN"
         rstdestino("pro_codigo_det") = VAR_PROY2   'Ado_datos.Recordset("edif_codigo")
         'VAR_PROY2 = Ado_datos.Recordset("edif_codigo")
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

Private Sub CmdEmail_Click()
    Set rs_aux12 = New ADODB.Recordset
    If rs_aux12.State = 1 Then rs_aux12.Close
    rs_aux12.Open "Select * from gc_beneficiario where (beneficiario_codigo = '" & dtc_benef2A.Text & "') ", db, adOpenStatic   '
    If rs_aux12.RecordCount > 0 Then
        TxtEmail.Text = rs_aux12!beneficiario_email
        TxtCelular.Text = rs_aux12!beneficiario_telefono_Cel
    'Else
    End If
    frm_benef.Visible = True
End Sub

Private Sub CmdGrabaCobro_Click()
  If dtc_codigo4A = "" Then
    MsgBox "Debe Elejir " + Lbl_Cobrador.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
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
  If cmd_fac.Text = "FACTURA" Then
    Dim MyPos, MyPos2 As Integer
    Dim ARROBA, punto As String
    If IsNumeric(dtc_codigo2A.Text) Then
    Else
        MsgBox "El NIT del Cliente a Facturar es Incorrecto, debe corregir y luego vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
    End If
    MyPos = 0
    MyPos2 = 0
    ARROBA = "@"
    punto = "."
    MyPos = InStr(1, dtc_email2A.Text, ARROBA, 1)
    MyPos2 = InStr(1, dtc_email2A.Text, punto, 1)
    'MyPos = Instr(4, SearchString, SearchChar, 1)
    If (Len(dtc_email2A.Text) > 10) Then
        If (MyPos > 0) And (MyPos2 > 0) Then
        Else
            MsgBox "Debe Registrar el EMail del Cliente a Facturar, !! Vuelva a Intentar ...", vbExclamation, "Atención"
            Exit Sub
        End If
    Else
        MsgBox "Debe Registrar el EMail del Cliente a Facturar, !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
    End If
  End If
  
  correlv = Ado_datos.Recordset!venta_codigo
  'If swnuevo = 2 Then
  'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'  If DTPFechaProg.Visible = False Then
'    If TxtCmpbte = "" Or TxtCmpbte = "0" Then
'       MsgBox "Debe Registrar el " + lbl_factura.Caption + " a emitir al Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
'      Exit Sub
'    End If
'  End If
  'fin PARA COBRANZA WWWWWWWWWWWWWWWWWWW If swnuevo = 1 Then
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "select sum(cobranza_programada_bs) as totbs2, sum (cobranza_programada_dol) as totdl2 from ao_ventas_cobranza_prog where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockOptimistic
    If IsNull(rs_aux3!totbs2) Then
        If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs And swnuevo = 1 Then
            MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
            If rs_aux3.State = 1 Then rs_aux3.Close
            Exit Sub
        Else
            'db.Execute " UPDATE ao_ventas_cabecera SET correl_cobro_prog = '0' where venta_codigo = " & correlv & " "
        End If
    Else
        If swnuevo = 1 Then
            If (rs_aux3!totbs2) + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                If (glusuario = "NROMERO" Or glusuario = "GSOLIZ" Or glusuario = "CPAREDES" Or glusuario = "RGIL" Or glusuario = "GMORA" Or glusuario = "DTERCEROS" Or glusuario = "CPLATA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS") Then
                    MsgBox "ADVERTENCIA, el Monto acumulado de cobranzas <" + lbl_monto.Caption + "> sobrepasaran <" + lbl_totalBs.Caption + "> . Solo en Caso de Liquidaciones ...", vbExclamation, "Atención"
                Else
                    MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el <" + lbl_totalBs.Caption + "> . !! Vuelva a Intentar ...", vbExclamation, "Atención"
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    Exit Sub
                End If
            End If
        Else
            If (rs_aux3!totbs2) - VAR_MBS2 + CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                If (glusuario = "MARTEAGA" Or glusuario = "GSOLIZ" Or glusuario = "RGIL" Or glusuario = "GMORA" Or glusuario = "DTERCEROS" Or glusuario = "CPLATA" Or glusuario = "ADMIN" Or glusuario = "CSALINAS") Then
                    MsgBox "ADVERTENCIA, el Monto acumulado de cobranzas <" + lbl_monto.Caption + "> sobrepasaran <" + lbl_totalBs.Caption + "> . Solo en Caso de Liquidaciones ...", vbExclamation, "Atención"
                Else
                    MsgBox "No puede programar un <" + lbl_monto.Caption + "> que sobrepase el Monto <" + lbl_totalBs.Caption + "> . !! Verifique por favor ...", vbExclamation, "Atención"
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    'Exit Sub
                End If
            End If
        End If
    End If
  'valida = 1
  'If valida = 1 And dtc_codigo4A <> "" Then
'    Set rstdestino = New ADODB.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
    'db.CommitTrans
    db.BeginTrans
    If swnuevo = 1 Then
      Set rs_aux1 = New ADODB.Recordset
      If rs_aux1.State = 1 Then rs_aux1.Close
      'rs_aux1.Open "select * from ao_ventas_cabecera where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo=" & Ado_datos.Recordset!venta_codigo & "  ", db, adOpenKeyset, adLockOptimistic
      rs_aux1.Open "select * from ao_ventas_cabecera where venta_codigo = " & correlv & "  ", db, adOpenKeyset, adLockOptimistic
      If rs_aux1.RecordCount > 0 Then
         'correldet2 = rs_aux1!correl_cobro_prog + 1
         If rs_aux1!correl_cobro_prog > 0 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            'rs_aux2.Open "Select * from ao_ventas_cobranza_prog where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo=" & Ado_datos.Recordset!venta_codigo & " and cobranza_prog_codigo = " & rs_aux1!correl_cobro_prog & " ", db, adOpenStatic
            rs_aux2.Open "Select * from ao_ventas_cobranza_prog where venta_codigo = " & correlv & " and cobranza_prog_codigo = " & rs_aux1!correl_cobro_prog & " ", db, adOpenStatic       '
            If rs_aux2.RecordCount > 0 Then
                If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
                    MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
                    If rs_aux1.State = 1 Then rs_aux1.Close
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    db.CommitTrans
                    Exit Sub
                End If
            End If
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "Select max(cobranza_prog_codigo) as Codigo from ao_ventas_cobranza_prog where venta_codigo = " & correlv & "  ", db, adOpenStatic
            If rs_aux2.RecordCount > 0 Then
                correldet2 = IIf(IsNull(rs_aux2!Codigo), "0", rs_aux2!Codigo)
            Else
                correldet2 = "0"
            End If
         Else
            db.Execute " UPDATE ao_ventas_cabecera SET correl_cobro_prog = '1' where venta_codigo = " & correlv & " "
            correldet2 = "0"
         End If
         'correldet2 = rs_aux1!correl_cobro_prog + 1
         'If rs_aux2!Codigo >= correldet2 Then
         '   correldet2 = rs_aux2!Codigo + 1
         'End If
         If correldet2 = "0" Then
            correldet2 = "1"
         Else
            correldet2 = correldet2 + 1 'rs_aux2!Codigo + 1
         End If
         db.Execute " UPDATE ao_ventas_cabecera SET correl_cobro_prog = " & correldet2 & " where venta_codigo = " & correlv & " "
         'rs_aux1!correl_cobro_prog = correldet2
         'rs_aux1.Update
      End If
      'Ado_datos16.Recordset.AddNew
      Ado_datos16.Recordset!cobranza_prog_codigo = correldet2
      Ado_datos16.Recordset!venta_codigo = correlv      'Ado_datos.Recordset("venta_codigo")
      Ado_datos16.Recordset!ges_gestion = Ado_datos.Recordset!ges_gestion
    End If
    If swnuevo = 2 Then
      If Ado_datos16.Recordset!cobranza_prog_codigo > 1 Then
        correldet2 = Ado_datos16.Recordset!cobranza_prog_codigo - 1
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "Select * from ao_ventas_cobranza_prog where ges_gestion='" & Ado_datos.Recordset!ges_gestion & "' and venta_codigo=" & correlv & " and cobranza_prog_codigo = " & correldet2 & " ", db, adOpenStatic
        If rs_aux2.RecordCount > 0 Then
          If DTPFechaProg.Value <= rs_aux2!cobranza_fecha_prog Then
              MsgBox "No puede registrar una " + lbl_fechas.Caption + " menor o igual a la anterior. !! Vuelva a Intentar ...", vbExclamation, "Atención"
              If rs_aux2.State = 1 Then rs_aux2.Close
              db.CommitTrans
              Exit Sub
          End If
        End If
      End If
    End If
      Ado_datos16.Recordset!beneficiario_codigo = dtc_benef2A.Text                                 'Codigo Beneficiario/Cliente
      Ado_datos16.Recordset!beneficiario_codigo_resp = dtc_codigo4A.Text                                                     'Codigo Cobrador
      'Ado_datos16.Recordset!nombre_cobrador = dtc_desc4A.Text   '+ " " + DtcMaterno.Text + " " + DtcNombre.Text    'Nombre Cobrador
      Ado_datos16.Recordset!cobranza_programada_bs = CDbl(TxtMonto.Text)                                  'Monto Programado Bs
      Ado_datos16.Recordset!cobranza_programada_dol = CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto Programado en Dolares
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
'      Ado_datos16.Recordset!cobranza_deuda_bs = 0   'CDbl(TxtMonto.Text)                                  'Monto Cobrado
'      Ado_datos16.Recordset!cobranza_deuda_dol = 0  'CDbl(TxtMonto.Text) / GlTipoCambioMercado        'Monto en Dolares
      'If TxtDscto.Text = "" Or TxtDscto.Text = "0" Or TxtDscto.Text = "0.00" Then
        Ado_datos16.Recordset!cobranza_descuento_bs = 0                                 'Descuento Bs
        Ado_datos16.Recordset!cobranza_descuento_dol = 0                                    'Descuento Dol
      'Else
      '  Ado_datos16.Recordset!cobranza_descuento_bs = CDbl(TxtDscto.Text)                                 'Descuento Bs
      '  Ado_datos16.Recordset!cobranza_descuento_dol = CDbl(TxtDscto.Text) / GlTipoCambioMercado        'Descuento Dol
      'End If
      If cmd_fac = "FACTURA" Then
        Ado_datos16.Recordset!doc_codigo_fac = "R-101"
        VAR_EMISION = "28"
      Else
        Ado_datos16.Recordset!doc_codigo_fac = "R-393"           '"R-103"
        VAR_EMISION = "37"
      End If
      If Txt_liquida.Text = "" Then
        Txt_liquida.Text = "NO"
      End If
      Ado_datos16.Recordset!es_liquidacion = Txt_liquida.Text
      Ado_datos16.Recordset!cobranza_total_bs = 0   'Ado_datos16.Recordset!cobranza_deuda_bs - Ado_datos16.Recordset!cobranza_descuento_bs               'Monto Total Bs
      Ado_datos16.Recordset!cobranza_total_dol = 0  'Ado_datos16.Recordset!cobranza_deuda_dol - Ado_datos16.Recordset!cobranza_descuento_dol               'Monto Total Dol
      'ini PARA COBRANZA WWWWWWWWWWWWWWWWWWW
      If Ado_datos16.Recordset!cobranza_programada_bs <> 0 Then
            Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos16.Recordset!cobranza_programada_bs)) + " BOLIVIANOS"
            'Ado_datos16.Recordset!Literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
      End If
      'Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value                                'Fecha de Cobranza
'      Call acumulaMont(Ado_datos16.Recordset("ges_gestion"), Ado_datos16.Recordset("venta_codigo"))
      
      'If Chk_plazo.Value = 1 Then
        lbl_plazo.Visible = True
        txt_plazo.Visible = True
        Ado_datos16.Recordset!cobranza_requisito_plazo = "S"
        Ado_datos16.Recordset!cobranza_concepto_plazo = txt_plazo.Text
      'Else
      '  lbl_plazo.Visible = False
      '  txt_plazo.Visible = False
      '  Ado_datos16.Recordset!cobranza_requisito_plazo = "N"
      '  Ado_datos16.Recordset!cobranza_concepto_plazo = "-"
      'End If
    
      Ado_datos16.Recordset!cobranza_observaciones = TxtObs.Text
      Ado_datos16.Recordset!proceso_codigo = "COM"
      Ado_datos16.Recordset!subproceso_codigo = "COM-02"
      Ado_datos16.Recordset!etapa_codigo = "COM-02-02"
      Ado_datos16.Recordset!clasif_codigo = "ADM"
      Ado_datos16.Recordset!doc_codigo = "R-105"                    'VERIFICAR "R-110"
      Ado_datos16.Recordset!doc_numero = Ado_datos.Recordset("venta_codigo")  'IIf(Txt_cod_cobro = "", "0", Txt_cod_cobro)
'      Ado_datos16.Recordset!doc_codigo_fac = ""
'      Ado_datos16.Recordset!cobranza_nro_factura = "0"       'Trim(TxtCmpbte)
'      Ado_datos16.Recordset!cobranza_nro_autorizacion = "0"       'Trim(TxtCmpbte)
      Ado_datos16.Recordset!poa_codigo = "3.1.2"
      'If DTPFechaProg.Visible = False Then
      '  Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value         'Fecha de Cobranza
      'Else
        Ado_datos16.Recordset!cobranza_fecha_cobro = DTPFechaCobro.Value         'Fecha de Cobranza
        Ado_datos16.Recordset!cobranza_fecha_prog = DTPFechaProg.Value           'Fecha Programada de Cobranza
      'End If
      Ado_datos16.Recordset!trans_codigo = VAR_EMISION
      Ado_datos16.Recordset!estado_codigo = "REG"
      Ado_datos16.Recordset!usr_codigo = glusuario
      Ado_datos16.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      Ado_datos16.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      Ado_datos16.Recordset.Update
    db.CommitTrans
    'Ado_datos16.Recordset!doc_numero = Ado_datos16.Recordset!cobranza_codigo       'Txt_cod_cobro.Text     ' "0"
  If swnuevo = 1 Then
    'Call abre_solicitud_lista
    'rc_Cobranza.Requery
    'Ado_datos16.Refresh
    'Ado_datos16.Recordset.MoveLast
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Enabled = True
    fraOpciones.Enabled = True
    FrmDetalle.Visible = True
    FrmCobranza.Visible = True
    FrmAlcance.Visible = True
    FrmCobros.Enabled = False
    TxtCobrador.Visible = True
    FrmABMDet.Visible = True
    FrmABMDet1.Visible = True
    FrmABMDet2.Visible = True
    swnuevo = 0
    gestion0 = Ado_datos.Recordset("ges_gestion")
    'correlv = Ado_datos.Recordset("correl_venta")
    nroventa = Ado_datos.Recordset("venta_codigo")
    
'  Set rstacumdet = New ADODB.Recordset
'  If rstacumdet.State = 1 Then rstacumdet.Close
'  rstacumdet.Open "select sum(deuda_cobrada) as Cobrobs from ao_ventas_cobranza_prog where ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' and venta_codigo = " & Ado_datos.Recordset("venta_codigo"), db, adOpenKeyset, adLockOptimistic
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

  'Else
  '  MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
  'End If
End Sub

Private Sub CmdGrabaDet_Click()
    If Left(dtc_codigo15, 2) = "NA" Or (dtc_codigo15 = "") Then
       sino = MsgBox("Desea crear un nuevo código de Equipo ? ", vbYesNo + vbQuestion, "Atención ...")
       If sino = vbYes Then
         Set rs_aux6 = New ADODB.Recordset
         If rs_aux6.State = 1 Then rs_aux6.Close
         rs_aux6.Open "select * from fc_partida_gasto where par_codigo = '43340' ", db, adOpenKeyset, adLockReadOnly
         If rs_aux6.RecordCount > 0 Then
            If Val(rs_aux6!correlativo36) < 10 Then
               VAR_OA = LTrim(rs_aux6!inicial) + "000" + LTrim(Str(rs_aux6!correlativo36 + 1))
            End If
            If Val(rs_aux6!correlativo36) > 9 And Val(rs_aux6!correlativo36) < 100 Then
               VAR_OA = LTrim(rs_aux6!inicial) + "00" + LTrim(Str(rs_aux6!correlativo36 + 1))
            End If
            If Val(rs_aux6!correlativo36) > 99 And Val(rs_aux6!correlativo36) < 1000 Then
               VAR_OA = LTrim(rs_aux6!inicial) + "0" + LTrim(Str(rs_aux6!correlativo36 + 1))
            End If
            If Val(rs_aux6!correlativo36) > 999 And Val(rs_aux6!correlativo36) < 10000 Then
               VAR_OA = LTrim(rs_aux6!inicial) + LTrim(Str(rs_aux6!correlativo36 + 1))
            End If
            'If Val(rs_aux6!correlativo36) > 9999 And Val(rs_aux6!correlativo36) < 100000 Then
            If Val(rs_aux6!correlativo36) > 9999 Then
               VAR_OA = LTrim(rs_aux6!inicial) + LTrim(Str(rs_aux6!correlativo36 + 1))
            End If
            'If Val(rs_aux6!correlativo36) > 99999 Then
            '   rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
            'End If
            'VAR_OA = "OA36" + LTrim(Str(rs_aux6!correlativo36 + 1))
            'VAR_OA = "36NB" + LTrim(Str(rs_aux6!correlativo36 + 1))
            'VAR_OA = LTrim(rs_aux6!inicial) + LTrim(Str(rs_aux6!correlativo36 + 1))
            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "select * from ac_bienes where bien_codigo = '" & VAR_OA & "' ", db, adOpenKeyset, adLockReadOnly
            If rs_aux7.RecordCount > 0 Then
                MsgBox "El Código de Equipo " + VAR_OA + " YA EXISTE, Consulte con el Administrador del Sistema y vuelva a Intentar !! ", vbExclamation, "Atención!"
                'db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
                VAR_NEW = "N"
                Exit Sub
            Else
                ado_datos14.Recordset!bien_codigo = Trim(VAR_OA)
                db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
                VAR_NEW = "S"
            End If
         Else
            VAR_NEW = "N"
         End If
       Else
            VAR_OA = Trim(dtc_codigo15.Text)
            VAR_NEW = "N"
       End If
    Else
        VAR_OA = ado_datos14.Recordset!bien_codigo
    End If
     parametro = Ado_datos.Recordset!unidad_codigo
    'If dtc_desc12 = "" Then
    '    MsgBox "Debe Elejir un Descuento X Tipo de Cliente, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    '    Exit Sub
    '  End If
    If dtc_codigo15 = "" Then
         MsgBox "Debe Elejir un Equipo para Vender, !! Vuelva a Intentar ...", vbExclamation, "Atención"
         VAR_OA = Trim(dtc_codigo15.Text)
         Exit Sub
    End If
    '  If dtc_desc13 = "" Then
    '    MsgBox "Debe Elejir el Almacen de Origen, !! Vuelva a Intentar ...", vbExclamation, "Atención"
    '    Exit Sub
    '  End If
    
    'If Val(dtc_stocktotal15.Text) >= Val(TxtCantidad.Text) Then
    '    VAR_PARTIDA = "OK"
    ' Aux
    'parametro
    'If Val(Dtc_Stock13.Text) >= Val(TxtCantidad.Text) Or Dtc_partida15.Text = "43340" Then
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "select * from av_solicitud_calculo_trafico where unidad_codigo = '" & VAR_UORIGEN & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " ", db, adOpenKeyset, adLockReadOnly
    rs_datos5.Open "select * from av_ventas_cotiza_equipo where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " ", db, adOpenKeyset, adLockReadOnly
    'If rs_datos5.RecordCount > 0 Then
        
'    If Dtc_partida15.Text = "43340" Then
          'fraOpciones.Visible = True
          'FraGrabarCancelar.Visible = False
          'TxtNroVenta.Enabled = True
          marca1 = Ado_datos.Recordset.Bookmark
          FrmEdita.Enabled = False
        '  DtGListaN.Enabled = True
          'cmdElige.Enabled = False
        '  dtc_codigo15.Visible = False
        '  dtc_desc15.Visible = False
          'txt_descripcion_venta.Enabled = False
        If swnuevo = 1 Then
          'ado_datos14.Recordset!venta_codigo_det = Ado_datos.Recordset("correl_venta")
          ado_datos14.Recordset!venta_codigo = correlv      'Ado_datos.Recordset("venta_codigo")
          ado_datos14.Recordset!ges_gestion = Ado_datos.Recordset("ges_gestion")
          ado_datos14.Recordset!bien_codigo = VAR_OA        'Trim(dtc_codigo15.Text)       'Codigo Bien (Equipo, Producto, etc)
          VAR_NEW = "N"
        End If
          'ado_datos14.Recordset!nro_licitacion = dtc_partida15.Text                       'Compra ??
          'ado_datos14.Recordset!nro_adjudica = 0 'Trim(DtcNroAdjudica.Text)                 'Codigo de Adjudicacion
          'ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
          'ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
          'ado_datos14.Recordset!par_codigo = Dtc_partida15                              'Partida
          'txt_descripcion_venta.Text = rs_datos5!tipo_eqp_descripcion + "- Codigo: " + VAR_OA + "- Modelo: " + Txt_modelo1
          ado_datos14.Recordset!tipo_descuento = IIf(dtc_codigo12.Text = "", "0", dtc_codigo12.Text)                      ' Tipo de Descuento
          ado_datos14.Recordset!almacen_codigo = IIf(dtc_codigo13.Text = "", "0", dtc_codigo13.Text)
          If TxtCantidad.Text = "" Then
            TxtCantidad.Text = "1"
          End If
          ado_datos14.Recordset!venta_det_cantidad = Val(IIf(TxtCantidad = "", 1, TxtCantidad)) 'Cantidad Vendida
          'ado_datos14.Recordset!codigo_solicitud = 0                                     'Nro.Solicitud de compra
          ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text)            'Precio Unitario de Venta
          'ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text)             'Precio Unitario de Venta
          If TxtDescuento = "" Or TxtDescuento = "0" Then
            TxtDescuento.Text = "0"
'            ado_datos14.Recordset!venta_descuento_bs = 0
'            ado_datos14.Recordset!venta_descuento_dol = 0
          Else
            ado_datos14.Recordset!venta_descuento_dol = CDbl(TxtDescuento.Text)     'Dcto por producto CON DESCUENTO
            ado_datos14.Recordset!venta_descuento_bs = Val(TxtDescuento) * GlTipoCambioMercado
          End If
          ado_datos14.Recordset!venta_precio_total_dol = (CDbl(TxtPrecioU.Text) - CDbl(TxtDescuento)) * Val(TxtCantidad)   'Precio Total Producto
          'If Val(lbltipo_Cambio) = 0 Then lbltipo_Cambio = 1
          'ado_datos14.Recordset!venta_precio_unitario_dol = CDbl(TxtPrecioU.Text) / GlTipoCambioMercado                'Precio Unitario Dolares
          ado_datos14.Recordset!venta_precio_unitario_bs = CDbl(TxtPrecioU.Text) * GlTipoCambioMercado            'Precio Unitario de Venta
          ado_datos14.Recordset!venta_precio_total_bs = (ado_datos14.Recordset!venta_precio_total_dol) * GlTipoCambioMercado
          'Call acumulaMont(Ado_datos.Recordset("ges_gestion"), Ado_datos.Recordset("venta_codigo"), Ado_datos.Recordset("venta_codigo"))
          If Txt_modelo.Text = "" Then
            Txt_modelo.Text = Txt_modelo1.Text
          End If
          ado_datos14.Recordset!modelo_codigo = Txt_modelo.Text
          ado_datos14.Recordset!modelo_codigo1 = Txt_modelo1.Text
          ado_datos14.Recordset!modelo_codigo_h = Txt_modelo2.Text
          ado_datos14.Recordset!modelo_codigo_x = Txt_modelo3.Text
          'If OpMod1.Value = True Then
            ado_datos14.Recordset!modelo_elegido = "S"
'            ado_datos14.Recordset!modelo_elegido_h = "N"
'            ado_datos14.Recordset!modelo_elegido_x = "0"
          'End If
'          If OpMod2.Value = True Then
''            ado_datos14.Recordset!modelo_elegido_h = "S"
'            ado_datos14.Recordset!modelo_elegido = "N"
''            ado_datos14.Recordset!modelo_elegido_x = "N"
'          End If
'          If OpMod2.Value = True Then
''            ado_datos14.Recordset!modelo_elegido_x = "S"
'            ado_datos14.Recordset!modelo_elegido = "N"
''            ado_datos14.Recordset!modelo_elegido_h = "N"
'          End If
         'INI GUARDA BIENES
         'parametro = Ado_datos.Recordset!unidad_codigo
         If VAR_NEW = "S" Then
            Set rs_aux8 = New ADODB.Recordset
            If rs_aux8.State = 1 Then rs_aux8.Close
            rs_aux8.Open "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " ", db, adOpenKeyset, adLockReadOnly
            If rs_aux8.RecordCount > 0 Then
                Select Case ado_datos14.Recordset!cotiza_codigo
                    Case 1
                        VAR_PARADAS = Trim(Str(rs_aux8!trafico_num_paradas))
                        VAR_PASAJEROS = rs_aux8!pasajeros_codigo
                    Case 2
                        VAR_PARADAS = Trim(Str(rs_aux8!trafico_num_paradas2))
                        VAR_PASAJEROS = rs_aux8!pasajeros_codigo2
                    Case 3
                        VAR_PARADAS = Trim(Str(rs_aux8!trafico_num_paradas3))
                        VAR_PASAJEROS = rs_aux8!pasajeros_codigo3
                    Case 4
                        VAR_PARADAS = Trim(Str(rs_aux8!trafico_num_paradas4))
                        VAR_PASAJEROS = rs_aux8!pasajeros_codigo4
                    Case Else
                        VAR_PARADAS = Trim(Str(rs_aux8!trafico_num_paradas))
                        VAR_PASAJEROS = rs_aux8!pasajeros_codigo
                End Select
            End If
            If rs_datos5.RecordCount > 0 Then
                txt_descripcion_venta = rs_datos5!tipo_eqp_descripcion + " Capacidad: " + VAR_PASAJEROS + " Personas - Paradas: " + VAR_PARADAS + "- Modelo: " + Txt_modelo1 + ""
            Else
                txt_descripcion_venta = "Equipo con Capacidad: " + VAR_PASAJEROS + " Personas - Paradas: " + VAR_PARADAS + "- Modelo: " + Txt_modelo1 + ""
            End If
            '" Personas - Velocidad: " + Str(rs_datos5!vel_equipo_m_s) + " m/s - Modelo: "
            'txt_descripcion_venta = rs_datos5!tipo_eqp_descripcion + " Capacidad: " + VAR_PASAJEROS + " Personas - Paradas: " + VAR_PARADAS + "- Modelo: " + Txt_modelo1 + ""
            'txt_descripcion_venta = " Capacidad: " + VAR_PASAJEROS + " Personas - Paradas: " + VAR_PARADAS + "- Modelo: " + Txt_modelo1 + ""
            'If txt_descripcion_venta.Text = "" Then
            '    txt_descripcion_venta.Text = rs_datos5!tipo_eqp_descripcion + "- Codigo: " + VAR_OA + "- Modelo: " + Txt_modelo1
            'Else
            '    txt_descripcion_venta.Text = rs_datos5!tipo_eqp_descripcion + " - " + txt_descripcion_venta.Text
            'End If
            ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
            ado_datos14.Recordset!grupo_codigo = "40000"
            ado_datos14.Recordset!subgrupo_codigo = "43000"
            ado_datos14.Recordset!par_codigo = IIf(Dtc_partida15.Text = "", "43340", Dtc_partida15.Text)
            Set rs_datos7 = New ADODB.Recordset
            If rs_datos7.State = 1 Then rs_datos7.Close
            rs_datos7.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " AND cotiza_codigo = " & ado_datos14.Recordset!cotiza_codigo & " ", db, adOpenKeyset, adLockReadOnly
            If rs_datos7.RecordCount > 0 Then
                VAR_PAIS = rs_datos7!pais_codigo
                VAR_TIPOEQP = rs_datos7!tipo_eqp
            Else
                MsgBox "Debe registrar el Pais de Origen en el Clasificador de Equipos !..."
                VAR_PAIS = "NN"
            End If
            
            Set rs_aux17 = New ADODB.Recordset
            If rs_aux17.State = 1 Then rs_aux17.Close
            
            Select Case ado_datos14.Recordset!venta_codigo_det
              Case "1"
                  VAR_EQP = "A"
                  VAR_OA2 = VAR_OA
              Case "2"
                  VAR_EQP = "B"
                  VAR_OA2 = LTrim(Txt_campo2.Text + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "3"
'                  VAR_EQP = "C"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "4"
'                  VAR_EQP = "D"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "5"
'                  VAR_EQP = "E"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "6"
'                  VAR_EQP = "F"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "7"
'                  VAR_EQP = "G"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "8"
'                  VAR_EQP = "H"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "9"
'                  VAR_EQP = "I"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
'              Case "10"
'                  VAR_EQP = "J"
'                  VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
              Case Else
                  rs_aux17.Open "select * from gc_alfabeto where opcion = " & ado_datos14.Recordset!venta_codigo_det & " ", db, adOpenKeyset, adLockReadOnly
                  If rs_aux17.RecordCount > 0 Then
                     VAR_EQP = rs_aux17!letra
                     VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
                  Else
                     VAR_EQP = "XX"
                     VAR_OA2 = LTrim(Left(Txt_campo2.Text, 8) + "-" + LTrim(Right(VAR_OA, 2)))
                  End If
            End Select
            
            db.Execute "insert into ac_bienes(grupo_codigo, subgrupo_codigo, bien_codigo, par_codigo, bien_descripcion, observaciones, bien_precio_compra, bien_precio_venta_base, bien_precio_venta_final, bien_precio_compra_dol, bien_precio_venta_base_dol, bien_precio_venta_final_dol, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo, modelo_codigo, bien_stock_minimo, bien_stock_inicial, bien_stock_ingreso, bien_stock_salida, bien_stock_actual, " & _
            "bien_total_compra_bs, bien_total_venta_bs, bien_utilidad_Bs, bien_codigo_anterior, bien_codigo_universal, bien_descripcion_anterior, bien_rotacion, pais_codigo, edif_codigo , archivo_foto2, archivo_foto, kit, estado_vigente, estado_codigo, usr_codigo, fecha_registro,  usr_codigo_apr, fecha_registro_apr) " & _
            "VALUES ('40000', '43000', '" & VAR_OA & "', '43340', '" & txt_descripcion_venta & "', '" & txt_descripcion_venta.Text & "', " & CDbl(TxtPrecioU.Text) & ", '0', '0',                  '0',                    '0',                        '0',                         'EQP',      'EQP',                    '1', '" & IIf(IsNull(rs_datos5!marca_codigo), "OTIS", rs_datos5!marca_codigo) & "', '" & Txt_modelo.Text & "', '1', '0', '0',   '0',  '0', " & _
            "'0',                  '0',                '0',               '" & VAR_EQP & "',  '" & VAR_TIPOEQP & "',   '-',                      'PROMEDIO', '" & VAR_PAIS & "', '" & rs_datos5!edif_codigo & "', '" & VAR_OA & "' + '.JPG', '" & VAR_OA & "' + '.JPG', '0', 'APR', 'APR', '" & glusuario & "', '" & Date & "', '" & glusuario & "', '" & Date & "' ) "
             
            'db.Execute "insert into ac_bienes(grupo_codigo, subgrupo_codigo, bien_codigo, par_codigo, bien_descripcion, bien_precio_compra, bien_precio_venta_base, bien_precio_venta_final, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, marca_codigo, bien_stock_minimo, bien_stock_inicial, bien_stock_ingreso, bien_stock_salida, bien_stock_actual, bien_total_compra_bs, bien_total_venta_bs, bien_utilidad_Bs, bien_codigo_anterior, bien_codigo_universal, bien_descripcion_anterior, pais_codigo, archivo_foto2, archivo_foto, estado_codigo, fecha_registro, usr_codigo) " & _
            '"VALUES ('40000', '43000', '" & VAR_OA & "', '43340', '" & txt_descripcion_venta & "', " & CDbl(TxtPrecioU.Text) & ", '0', '0', 'EQP', 'EQP', '1', 'S/M', '1', '0', '0', '0', '0', '0', '0', '0', '" & VAR_EQP & "', '" & VAR_TIPOEQP & "', '-', '" & VAR_PAIS & "', '" & VAR_OA & "' + '.JPG', '" & VAR_OA & "' + '.JPG', 'REG', '" & Date & "', '" & glusuario & "') "
            
            Txt_campo2.Text = VAR_OA2
            db.Execute "update ao_ventas_cabecera set unidad_codigo_ant = '" & VAR_OA2 & "' where venta_codigo = " & Ado_datos.Recordset!venta_codigo & " "
            
            '"VALUES ('" & Trim(dtc_grupo15.Text) & "', '" & Trim(dtc_subgrupo15.Text) & "', '" & VAR_OA & "', '" & Dtc_partida15 & "', '" & txt_descripcion_venta & "', " & CDbl(TxtPrecioU.Text) & ", '0', '0', 'EQP', 'EQP', '1', 'S/M', '1', '0', '0', '0', '0', '0', '0', '0', '-', '-', '-', 'NN', '-' + '2.JPG', '-' + '.JPG', 'REG', '" & Date & "', '" & glusuario & "') "
         Else
            ado_datos14.Recordset!concepto_venta = txt_descripcion_venta                  'Descripcion y Caracteristicas
            ado_datos14.Recordset!grupo_codigo = Trim(dtc_grupo15.Text)
            ado_datos14.Recordset!subgrupo_codigo = Trim(dtc_subgrupo15.Text)
            ado_datos14.Recordset!par_codigo = IIf(Dtc_partida15.Text = "", "43340", Dtc_partida15.Text)                             'Partida
         End If
         'FIN GUARDA BIENES
          ado_datos14.Recordset!estado_codigo = "REG"
          ado_datos14.Recordset!usr_codigo = glusuario
          ado_datos14.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
          ado_datos14.Recordset!hora_registro = Format(Time, "hh:mm:ss")
          ado_datos14.Recordset.Update
        'db.CommitTrans
        'actualiza MODELO del equipo
        'db.Execute "update ac_bienes set modelo_codigo = '" & ado_datos14.Recordset!modelo_codigo & "' Where grupo_codigo = '" & ado_datos14.Recordset!grupo_codigo & "' And subgrupo_codigo = '" & ado_datos14.Recordset!subgrupo_codigo & "'  And bien_codigo = '" & ado_datos14.Recordset!bien_codigo & "' "
        'Acumula MONTOS
'        Call acumulaMont(Ado_datos.Recordset("solicitud_codigo"), Ado_datos.Recordset("venta_codigo"))
        
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        FraNavega.Enabled = True
        FrmDetalle.Enabled = True
        FrmAlcance.Visible = True
        FrmCobranza.Visible = True
        FrmABMDet.Visible = True
        FrmABMDet1.Visible = True
        FrmABMDet2.Visible = True
        Call ABRIR_TABLA_DET
'        If Ado_datos.Recordset("estado_codigo") = "REG" Then
'          Call OptFilGral1_Click
'        Else
'          Call OptFilGral2_Click
'        End If
'        'Call OptFilGral1_Click
'        Ado_datos.Recordset.Move marca1 - 1
        If swnuevo = 1 Then
          'Call abre_ventas_det
          'rs_datos14.Requery
          'ado_datos14.Refresh
          'ado_datos14.Recordset.MoveLast
          
        End If
        swnuevo = 0
'    Else
'        MsgBox "Saldo Insuficiente en Almacen Origen, debe realizar Transferencia de otro Almacen, Luego Intente nuevamente !..."
'    End If
'  'Else
'  '  MsgBox "Saldo Insuficiente en Stock General (Todos los Almacenes), Intente nuevamente !..."
'  'End If
'          End If
'    End If
'Else
'End If
'End If

End Sub

Private Sub BtnImprimir2_Click()
  If Ado_datos16.Recordset.RecordCount > 0 Then
    Dim iResult As Variant  ', i%, y%
    'CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_R-105_kardex.rpt"
    CryR01.ReportFileName = App.Path & "\reportes\ventas\ar_cronograma_para_cobranza.rpt"
    CryR01.WindowShowRefreshBtn = True
    CryR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CryR01.StoredProcParam(1) = Me.Ado_datos.Recordset!venta_codigo
    CryR01.StoredProcParam(2) = Me.Ado_datos16.Recordset!cobranza_prog_codigo
    'Literal por el Total de la Compra
    var_literal = Literal(CStr(Ado_datos.Recordset!venta_monto_total_bs)) + " BOLIVIANOS"
    CryR01.Formulas(1) = "literalcobro = '" & var_literal & "' "
    'CryR01.Formulas(1) = "literalcobro = '" & Ado_datos16.Recordset!Literal & "' "
    CryR01.Formulas(2) = "correlcobro = '" & Ado_datos16.Recordset!cobranza_prog_codigo & "' "
    '.StoredProcParam(3) = Me.Ado_datos16.Recordset!Literal
    iResult = CryR01.PrintReport
    If iResult <> 0 Then MsgBox CryR01.LastErrorNumber & " : " & CryR01.LastErrorString, vbCritical, "Error de impresión"
  Else
    MsgBox "No se puede IMPRIMIR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
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

Private Sub dtc_aux7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_aux7.BoundText
    dtc_desc7.BoundText = dtc_aux7.BoundText
End Sub

Private Sub dtc_benef2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_benef2A.BoundText
    dtc_desc2A.BoundText = dtc_benef2A.BoundText
    dtc_email2A.BoundText = dtc_benef2A.BoundText
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
'    dtc_aux4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    dtc_aux7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    Dtc_aux2.BoundText = dtc_desc2.BoundText
    Dtc_deudor2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc2A_LostFocus()
    'txt_plazo.Text = "SERVICIO DE PROVISION E INSTALACION DE ASCENSORES Y ESCALERAS MECANICAS, SEGUN CONTRATO " + Txt_campo2.Text
    txt_plazo.Text = "SERVICIO DE PROVISION E INSTALACION DE ASCENSORES, SEGUN CONTRATO " + Txt_campo2.Text
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
'    dtc_aux4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    dtc_codigo7.BoundText = dtc_desc7.BoundText
    dtc_aux7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub Dtc_deudor2_Click(Area As Integer)
    dtc_codigo2.BoundText = Dtc_deudor2.BoundText
    Dtc_aux2.BoundText = Dtc_deudor2.BoundText
    dtc_desc2.BoundText = Dtc_deudor2.BoundText
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
    dtc_benef2A.BoundText = dtc_codigo2A.BoundText
    dtc_email2A.BoundText = dtc_codigo2A.BoundText
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

Private Sub Command1_Click()
'    'Form1.Show
'  If Ado_datos.Recordset.RecordCount > 0 Then
'     If Ado_datos.Recordset("estado_codigo") <> "ANL" Then
'       sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'           'ASIGNA A VARIABLES CAMPOS CLAVES
'           correlv = Ado_datos.Recordset!venta_codigo
'           VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'           VAR_TIPOV = Ado_datos.Recordset!venta_tipo
'           VAR_PROY2 = Ado_datos.Recordset!edif_codigo
'           VAR_UNIDCOD = Ado_datos.Recordset!unidad_codigo
'           VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'           VAR_CITE = Ado_datos.Recordset!unidad_codigo_ant
'           VAR_GLOSA = Ado_datos.Recordset!venta_descripcion
'           VAR_DOL2 = Round(Ado_datos.Recordset("venta_monto_total_dol"), 2)
'           VAR_BS2 = Round(Ado_datos.Recordset("venta_monto_total_bs"), 2)
'           VAR_UNIMED = Ado_datos.Recordset!unimed_codigo
'           VAR_BEND = dtc_desc2.Text
'           VAR_EDIFD = dtc_desc3.Text
'           VAR_UNID = dtc_desc1.Text
'           VAR_DPTO = Left(VAR_PROY2, 1)
'           VARG_ORGD = ""
'           VAR_CTAD = ""
'
'           If Ado_datos.Recordset("venta_tipo") = "C" Or Ado_datos.Recordset("venta_tipo") = "V" Or Ado_datos.Recordset("venta_tipo") = "G" Or Ado_datos.Recordset("venta_tipo") = "L" Then
'                db.Execute "update gc_beneficiario set beneficiario_deudor = 'SI' where beneficiario_codigo = '" & dtc_codigo2 & "' "
'           End If
''           ' APRUEBA ao_ventas_cabecera
''           db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
'
'           ' Actualiza Saldos ac_bienes
'           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = av_acumula_ventas_detalle.venta_det_cantidad from ac_bienes, av_acumula_ventas_detalle Where ac_bienes.grupo_codigo = av_acumula_ventas_detalle.grupo_codigo And ac_bienes.subgrupo_codigo = av_acumula_ventas_detalle.subgrupo_codigo And ac_bienes.bien_codigo = av_acumula_ventas_detalle.bien_codigo"
'           db.Execute "update ac_bienes set bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida"
'
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
'            ' GRABA Nombre de Archivo en ao_ventas_cabecera. VERIFICAR JQA 2014-07-08
'            'rs_datos!doc_numero = Txt_campo1.Caption
'            'VAR_ARCH = RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
'
'            If IsNull(Ado_datos.Recordset!doc_codigo) Or IsNull(Ado_datos.Recordset!doc_numero) Then
'              ' Validar consistencia de datos.
'            Else
'              VAR_ARCH = "COM_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
'            End If
'
'            'VAR_ARCH = "COM_" + RTrim(RTrim(Ado_datos.Recordset!doc_codigo) + "-") + LTrim(Str(Ado_datos.Recordset!doc_numero))
'            db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo = '" & VAR_ARCH & "' + '.PDF' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
'            db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.archivo_respaldo_cargado = 'N' Where ao_ventas_cabecera.ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' And ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           ' REVISAR JQ-2014-JUL-05
'            'INI HABILITA ALMACEN PARA venta_tipo="V" (PREVENTA)
'            'correlv = 2
'           'If Ado_datos.Recordset!venta_tipo <> "V" Then
''           If VAR_TIPOV <> "V" And VAR_TIPOV <> "L" Then
''             Set rsAuxDetalle = New ADODB.Recordset
''             If rsAuxDetalle.State = 1 Then rsAuxDetalle.Close
''             rsAuxDetalle.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockBatchOptimistic
''             'Set AdoAux.Recordset = rsAuxDetalle
''             If rsAuxDetalle.RecordCount > 0 Then
''               'AdoAux.Recordset.MoveFirst
''               rsAuxDetalle.MoveFirst
''               While Not rsAuxDetalle.EOF   ' AdoAux.Recordset.EOF
''                 Set rs_almacen2 = New ADODB.Recordset
''                 If rs_almacen2.State = 1 Then rs_almacen2.Close
''                 rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = '" & rsAuxDetalle!almacen_codigo & "' and bien_codigo = '" & rsAuxDetalle!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
''                 If rs_almacen2.RecordCount > 0 Then
''                     db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida = " & rsAuxDetalle!venta_det_cantidad & "  from ao_almacen_totales, ao_ventas_detalle Where ao_almacen_totales.almacen_codigo = '" & rsAuxDetalle!almacen_codigo & "'   And ao_almacen_totales.bien_codigo = '" & rsAuxDetalle!bien_codigo & "'   "
''                     'AdoAux.Recordset.MoveNext
''                 Else
''                     'GRABA ALMACEN DETALLE
''                    Set rs_aux4 = New ADODB.Recordset
''                    If rs_aux4.State = 1 Then rs_aux4.Close
''                    rs_aux4.Open "Select * from av_acumula_compras_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = '" & Ado_datos.Recordset!solicitud_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux4.Open "Select * from ao_almacen_totales where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux4.RecordCount > 0 Then
''                        db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
''                    Else
''                        If Ado_datos.Recordset!venta_tipo = "V" Then
''                            'db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
''                            db.Execute "INSERT INTO ao_almacen_totales (almacen_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, stock_ingreso) VALUES (" & rsAuxDetalle!almacen_codigo & ", '" & rsAuxDetalle!bien_codigo & "', '" & rsAuxDetalle!grupo_codigo & "', '" & rsAuxDetalle!subgrupo_codigo & "', '" & rsAuxDetalle!par_codigo & "' , " & rsAuxDetalle!venta_det_cantidad & ")"
''                        Else
''                            'MsgBox "Error Verifique la Adjudicación de Bienes (Equipos, Repuestos u otros) ..."
''                        End If
''                    End If
''                 End If
''                 rsAuxDetalle.MoveNext
''               Wend
''               db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida"
''             Else
''                MsgBox "Error Verifique la Venta de Productos..."
''             End If
''           End If
'           'FIN HABILITA ALMACEN PARA venta_tipo="V" (PREVENTA)
'
'           'marca1 = Ado_datos.Recordset.Bookmark
'           'Ado_datos.Recordset.Requery
'    '       Ado_datos.Refresh
'           'Ado_datos.Recordset.Move marca1 - 1
'           Call Contabiliza_venta
'
'           'INI GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
'           If VAR_TIPOV = "V" Or VAR_TIPOV = "L" Or VAR_TIPOV = "G" Then
'           'If Ado_datos.Recordset!venta_tipo = "V" Then
'             Set rs_aux1 = New ADODB.Recordset
'             If rs_aux1.State = 1 Then rs_aux1.Close
'             rs_aux1.Open "select * from ao_ventas_alcance where venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockBatchOptimistic
'             If rs_aux1.RecordCount > 0 Then
'               rs_aux1.MoveFirst
'               While Not rs_aux1.EOF
'                 VAR_COD1 = rs_aux1!unidad_codigo_tec
'                 VAR_CANT0 = Round((rs_aux1!venta_tiempo_dias / 30), 0)
'                 If VAR_COD1 = "COMEX" Then         'INI GRABA CRONOGRAMA COMEX
'                    'EQUIPO
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
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
'                    rs_aux3.Open "select * from ao_compra_cabecera where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                    'beneficiario_codigo_resp,'doc_numero,estado_codigo_tra, estado_codigo_nac, estado_codigo_des, hora_registro, usr_codigo_aprueba,'                      fecha_registro_aprueba
'
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        'rs_aux3!compra_codigo = 0      'Autonumerico
'                        rs_aux3!unidad_codigo_adm = VAR_COD1
'                        rs_aux3!solicitud_codigo_adm = correldetalle
'                        rs_aux3!unidad_codigo = VAR_UNIDCOD
'                        rs_aux3!solicitud_codigo = VAR_SOL
'                        rs_aux3!edif_codigo = VAR_PROY2
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!solicitud_tipo = "15"
'                        rs_aux3!venta_tipo = VAR_TIPOV
'                        rs_aux3!unidad_codigo_ant = VAR_CITE
'                        rs_aux3!compra_fecha = Date
'                        rs_aux3!compra_DESCRIPCION = "COMPRA POR: " + VAR_GLOSA
'                        rs_aux3!compra_observaciones = "PROVISION Y/O IMPORTACION DE EQUIPOS"
'                        rs_aux3!compra_cantidad_total = Ado_datos.Recordset!venta_cantidad_total
'                        rs_aux3!compra_monto_bs = VAR_BS2
'                        rs_aux3!tipo_moneda = "USD"
'                        rs_aux3!compra_monto_dol = VAR_DOL2
'                        rs_aux3!proceso_codigo = "CMX"
'                        rs_aux3!subproceso_codigo = "CMX-01"
'                        rs_aux3!etapa_codigo = "CMX-01-01"
'                        rs_aux3!clasif_codigo = "CMX"
'                        rs_aux3!doc_codigo = "R-207"
'                        rs_aux3!poa_codigo = "4.1.1"
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
'
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        rs_aux4.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                                db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
'                                "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & rs_aux4!venta_codigo_det & ", '" & rs_aux4!bien_codigo & "', '1', " & rs_aux4!venta_precio_unitario_bs & ", '0', " & rs_aux4!venta_precio_total_bs & ", " & rs_aux4!venta_precio_unitario_dol & ", '0', " & rs_aux4!venta_precio_total_dol & ", '" & concepto_venta & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
'                                rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        'cargar ADJUDICA_COMPRA Y CRONOGRAMA
'                        '
'                    End If
'                    'WWWWWWWWWW
'                 Else
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "select * from gc_unidad_ejecutora where unidad_codigo = '" & VAR_COD1 & "'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux2.RecordCount > 0 Then
'                       rs_aux2!correl_crono = rs_aux2!correl_crono + 1
'                       correldetalle = rs_aux2!correl_crono
'                       rs_aux2.Update
'                    End If
'
'                    Set rs_aux3 = New ADODB.Recordset
'                    If rs_aux3.State = 1 Then rs_aux3.Close
'                    rs_aux3.Open "select * from to_cronograma where unidad_codigo_tec = '" & VAR_COD1 & "' AND tec_plan_codigo = " & correldetalle & " ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux3.Open "select * from to_cronograma where unidad_codigo_tec = '" & VAR_COD1 & "' AND edif_codigo = '" & VAR_PROY2 & "' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux3.RecordCount = 0 Then
'                        rs_aux3.AddNew
'                        rs_aux3!ges_gestion = glGestion     'Year(Date)
'                        rs_aux3!unidad_codigo_tec = VAR_COD1
'                        rs_aux3!tec_plan_codigo = correldetalle
'                        rs_aux3!unidad_codigo = VAR_UNIDCOD        'Ado_datos.Recordset!unidad_codigo
'                        rs_aux3!solicitud_codigo = VAR_SOL    'Ado_datos.Recordset!solicitud_codigo
'                        rs_aux3!edif_codigo = VAR_PROY2      'Ado_datos.Recordset!edif_codigo
'                        rs_aux3!venta_codigo = correlv  'Ado_datos.Recordset!venta_codigo
'                        rs_aux3!compra_codigo = 0
'                        rs_aux3!adjudica_codigo = 0
'                        rs_aux3!tec_plan_fecha = Date
'                        rs_aux3!beneficiario_codigo = VAR_BENEF
'                        rs_aux3!unidad_codigo_ant = VAR_CITE
'                        rs_aux3!unimed_codigo = VAR_UNIMED
'                        ' Fechas de ao_ventas_alcance
'                        rs_aux3!fecha_inicio_tec = rs_aux1!fecha_inicio_alcance
'                        rs_aux3!fecha_fin_tec = rs_aux1!fecha_fin_alcance
'                        rs_aux3!tec_tiempo_dias = rs_aux1!venta_tiempo_dias
'                        rs_aux3!tec_cantidad_unidades = VAR_CANT0
'                        Select Case VAR_COD1
'                            Case "DNINS"                        'INI GRABA CRONOGRAMA INSTALACIONES
'                                rs_aux3!tec_plan_concepto = "INSTALACION DE: " + VAR_GLOSA
'                                rs_aux3!proceso_codigo = "COM"
'                                rs_aux3!subproceso_codigo = "COM-03"
'                                rs_aux3!etapa_codigo = "COM-03-02"
'                                rs_aux3!clasif_codigo = "TEC"
'                                rs_aux3!doc_codigo = "R-362"
'                                rs_aux3!poa_codigo = "3.2.2"
'
'                                   'db.Execute "INSERT INTO to_cronograma (ges_gestion, unidad_codigo_tec, tec_plan_codigo, unidad_codigo, solicitud_codigo, venta_codigo, compra_codigo, adjudica_codigo, tec_plan_concepto, tec_plan_fecha, beneficiario_codigo, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, poa_codigo, estado_codigo, usr_codigo, fecha_registro)
'                                   ' values ('" & year(date) & "', '" & rs_aux1!unidad_codigo_tec & "', " & correldetalle & ", '" & Ado_datos.Recordset!unidad_codigo & "', " & Ado_datos.Recordset!solicitud_codigo & ", " & Ado_datos.Recordset!venta_codigo & ", '0', '0', '" & Ado_datos.Recordset!venta_descripcion & "', '" & DATE & "')
'                                   'SELECT " & rs_aux4!almacen_codigo & ", '" & rs_aux4!bien_codigo & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "' , '" & rs_aux4!bien_cantidad_adjudica & "' FROM av_acumula_compras_detalle WHERE almacen_codigo = '" & rs_almacen2!almacen_codigo & "'   And bien_codigo = '" & rs_almacen2!bien_codigo & "'    "
'                                   'FIN GRABA CRONOGRAMA INSTALACIONES
'                                '      rs_aux2("tipo_moneda") = VAR_MONEDA
'
'                            Case "DNAJS"
'                                rs_aux3!tec_plan_concepto = "AJUSTE DE: " + VAR_GLOSA
'                                rs_aux3!proceso_codigo = "TEC"
'                                rs_aux3!subproceso_codigo = "TEC-01"
'                                rs_aux3!etapa_codigo = "TEC-01-02"
'                                rs_aux3!clasif_codigo = "TEC"
'                                rs_aux3!doc_codigo = "R-378"
'                                rs_aux3!doc_numero = correldetalle
'                                rs_aux3!poa_codigo = "3.2.6"     'OJO
'
'                            Case "DNMAN"
'                                rs_aux3!tec_plan_concepto = "MANTENIMIENTO GRATUITO DE: " + VAR_GLOSA
'                                rs_aux3!proceso_codigo = "TEC"
'                                rs_aux3!subproceso_codigo = "TEC-02"
'                                rs_aux3!etapa_codigo = "TEC-02-02"
'                                rs_aux3!clasif_codigo = "TEC"
'                                rs_aux3!doc_codigo = "R-302"
'                                rs_aux3!doc_numero = correldetalle
'                                rs_aux3!poa_codigo = "3.2.3"     'OJO
'
'                            Case Else
'                                MsgBox "No se ha definido el tipo " & vbCrLf & " de registro que está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'                                If rstdestino.State = 1 Then rstdestino.Close
'                                Exit Sub
'                        End Select
'                        rs_aux3!estado_codigo = "REG"
'                        rs_aux3!usr_codigo = glusuario
'                        rs_aux3!fecha_registro = Date
'                        rs_aux3.Update
'                        'DETALLE
'                        Set rstdestino = New ADODB.Recordset
'                        If rstdestino.State = 1 Then rstdestino.Close
'                        rstdestino.Open "select * from to_cronograma_detalle  ", db, adOpenKeyset, adLockBatchOptimistic
'
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        rs_aux4.Open "select * from ao_ventas_detalle where venta_codigo= " & correlv & "  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                           rs_aux4.MoveFirst
'                           While Not rs_aux4.EOF
'                                VAR_CANT9 = IIf(IsNull(rs_aux4!bien_cantidad_por_empaque), 1, rs_aux4!bien_cantidad_por_empaque)
'                                db.Execute "INSERT INTO to_cronograma_detalle (ges_gestion, unidad_codigo_tec, tec_plan_codigo, bien_codigo, beneficiario_codigo, grupo_codigo, subgrupo_codigo, par_codigo, munic_codigo, fecha_inicio, fecha_fin, bien_tiempo_dias, hora_inicio, hora_fin, estado_codigo, usr_codigo, fecha_registro, bien_cantidad_por_empaque) " & _
'                                "VALUES ('" & glGestion & "', '" & VAR_COD1 & "', " & correldetalle & ", '" & rs_aux4!bien_codigo & "', '0', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '" & Left(VAR_PROY2, 5) & "', '" & Format(rs_aux1!fecha_inicio_alcance, "dd/mm/yyyy") & "', '" & Format(rs_aux1!fecha_fin_alcance, "dd/mm/yyyy") & "', " & rs_aux1!venta_tiempo_dias & ", '8:00', '18:30', 'REG', '" & glusuario & "', '" & Date & "', " & VAR_CANT9 & ")"
'                                rs_aux4.MoveNext
'                           Wend
'                        End If
'                        If rstdestino.State = 1 Then rstdestino.Close
'
'                    End If
'                 End If
'                 rs_aux1.MoveNext
'               Wend
'             End If
'           End If
'           ' APRUEBA ao_ventas_cabecera
'           db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo = 'APR' Where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           'Actualiza Cite Trñamite (unidad_codigo_ant)
'           db.Execute "update ao_solicitud set ao_solicitud.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud inner join ao_ventas_cabecera on ao_solicitud.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           db.Execute "update ao_solicitud_calculo_trafico set ao_solicitud_calculo_trafico.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_calculo_trafico inner join ao_ventas_cabecera on ao_solicitud_calculo_trafico.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_calculo_trafico.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           db.Execute "update ao_solicitud_cotiza_modelo set ao_solicitud_cotiza_modelo.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_cotiza_modelo inner join ao_ventas_cabecera on ao_solicitud_cotiza_modelo.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_cotiza_modelo.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           db.Execute "update ao_solicitud_cotiza_venta set ao_solicitud_cotiza_venta.unidad_codigo_ant = ao_ventas_cabecera.unidad_codigo_ant from ao_solicitud_cotiza_venta inner join ao_ventas_cabecera on ao_solicitud_cotiza_venta.unidad_codigo =ao_ventas_cabecera.unidad_codigo and ao_solicitud_cotiza_venta.solicitud_codigo = ao_ventas_cabecera.solicitud_codigo where ao_ventas_cabecera.venta_codigo = " & correlv & " "
'           'FIN GENERA INFORMACION COMEX, INSTALACION, AJUSTE Y/O MANTENIMIENTO
'           Call OptFilGral1_Click
'       End If
'     End If
'
' Else
'    MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
' End If

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
        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + Txt_campo2.Text
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
        TxtConcepto.Text = lbl_titulo + " - " + dtc_desc11 + " - " + Txt_campo2.Text
        'TxtPlazo.Visible = True
    End If
    If dtc_codigo11.Text = "C" Or dtc_codigo11.Text = "E" Then
            TxtConcepto.Text = "VENTA AL CONTADO - " + Txt_campo2.Text
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

Private Sub dtc_codigo15_Click(Area As Integer)
    dtc_desc15.BoundText = dtc_codigo15.BoundText
    dtc_unimed15.BoundText = dtc_codigo15.BoundText
    dtc_stocktotal15.BoundText = dtc_codigo15.BoundText
    dtc_grupo15.BoundText = dtc_codigo15.BoundText
    dtc_subgrupo15.BoundText = dtc_codigo15.BoundText
    Dtc_partida15.BoundText = dtc_codigo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_codigo15.BoundText
    dtc_precioventabase15.BoundText = dtc_codigo15.BoundText
    dtc_preciocompra15.BoundText = dtc_codigo15.BoundText
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

Private Sub dtc_email2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_email2A.BoundText
    dtc_benef2A.BoundText = dtc_email2A.BoundText
    dtc_desc2A.BoundText = dtc_email2A.BoundText
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

Private Sub dtc_subgrupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_subgrupo15.BoundText
    dtc_desc15.BoundText = dtc_subgrupo15.BoundText
    dtc_unimed15.BoundText = dtc_subgrupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_subgrupo15.BoundText
    dtc_grupo15.BoundText = dtc_subgrupo15.BoundText
    Dtc_partida15.BoundText = dtc_subgrupo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_subgrupo15.BoundText
    dtc_precioventabase15.BoundText = dtc_subgrupo15.BoundText
    dtc_preciocompra15.BoundText = dtc_subgrupo15.BoundText
End Sub

Private Sub dtc_partida15_Click(Area As Integer)
    dtc_desc15.BoundText = Dtc_partida15.BoundText
    dtc_unimed15.BoundText = Dtc_partida15.BoundText
    dtc_stocktotal15.BoundText = Dtc_partida15.BoundText
    dtc_grupo15.BoundText = Dtc_partida15.BoundText
    dtc_subgrupo15.BoundText = Dtc_partida15.BoundText
    dtc_codigo15.BoundText = Dtc_partida15.BoundText
    dtc_precioventafinal15.BoundText = Dtc_partida15.BoundText
    dtc_precioventabase15.BoundText = Dtc_partida15.BoundText
    dtc_preciocompra15.BoundText = Dtc_partida15.BoundText
End Sub

Private Sub dtc_desc15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_desc15.BoundText
    dtc_unimed15.BoundText = dtc_desc15.BoundText
    dtc_stocktotal15.BoundText = dtc_desc15.BoundText
    dtc_grupo15.BoundText = dtc_desc15.BoundText
    dtc_subgrupo15.BoundText = dtc_desc15.BoundText
    Dtc_partida15.BoundText = dtc_desc15.BoundText
    dtc_precioventafinal15.BoundText = dtc_desc15.BoundText
    dtc_precioventabase15.BoundText = dtc_desc15.BoundText
    dtc_preciocompra15.BoundText = dtc_desc15.BoundText
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
    Call AbreAlmacen
End Sub

Private Sub dtc_codigo12_Click(Area As Integer)
    Dtc_aux12.BoundText = dtc_codigo12.BoundText
    dtc_desc12.BoundText = dtc_codigo12.BoundText
End Sub

Private Sub dtc_grupo15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_grupo15.BoundText
    dtc_desc15.BoundText = dtc_grupo15.BoundText
    dtc_unimed15.BoundText = dtc_grupo15.BoundText
    dtc_stocktotal15.BoundText = dtc_grupo15.BoundText
    dtc_subgrupo15.BoundText = dtc_grupo15.BoundText
    Dtc_partida15.BoundText = dtc_grupo15.BoundText
    dtc_precioventafinal15.BoundText = dtc_grupo15.BoundText
    dtc_precioventabase15.BoundText = dtc_grupo15.BoundText
    dtc_preciocompra15.BoundText = dtc_grupo15.BoundText
End Sub

Private Sub dtc_stocktotal15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_stocktotal15.BoundText
    dtc_desc15.BoundText = dtc_stocktotal15.BoundText
    dtc_unimed15.BoundText = dtc_stocktotal15.BoundText
    dtc_grupo15.BoundText = dtc_stocktotal15.BoundText
    dtc_subgrupo15.BoundText = dtc_stocktotal15.BoundText
    Dtc_partida15.BoundText = dtc_stocktotal15.BoundText
    dtc_precioventafinal15.BoundText = dtc_stocktotal15.BoundText
    dtc_precioventabase15.BoundText = dtc_stocktotal15.BoundText
    dtc_preciocompra15.BoundText = dtc_stocktotal15.BoundText
End Sub

Private Sub Dtc_aux12_Click(Area As Integer)
    dtc_codigo12.BoundText = Dtc_aux12.BoundText
    dtc_desc12.BoundText = Dtc_aux12.BoundText
End Sub

Private Sub dtc_precioventafinal15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_precioventafinal15.BoundText
    dtc_desc15.BoundText = dtc_precioventafinal15.BoundText
    dtc_unimed15.BoundText = dtc_precioventafinal15.BoundText
    dtc_grupo15.BoundText = dtc_precioventafinal15.BoundText
    dtc_subgrupo15.BoundText = dtc_precioventafinal15.BoundText
    Dtc_partida15.BoundText = dtc_precioventafinal15.BoundText
    dtc_stocktotal15.BoundText = dtc_precioventafinal15.BoundText
    dtc_precioventabase15.BoundText = dtc_precioventafinal15.BoundText
    dtc_preciocompra15.BoundText = dtc_precioventafinal15.BoundText
End Sub

Private Sub dtc_preciocompra15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_preciocompra15.BoundText
    dtc_desc15.BoundText = dtc_preciocompra15.BoundText
    dtc_unimed15.BoundText = dtc_preciocompra15.BoundText
    dtc_stocktotal15.BoundText = dtc_preciocompra15.BoundText
    dtc_grupo15.BoundText = dtc_preciocompra15.BoundText
    dtc_subgrupo15.BoundText = dtc_preciocompra15.BoundText
    Dtc_partida15.BoundText = dtc_preciocompra15.BoundText
    dtc_precioventafinal15.BoundText = dtc_preciocompra15.BoundText
    dtc_precioventabase15.BoundText = dtc_preciocompra15.BoundText
End Sub

Private Sub dtc_stock13_Click(Area As Integer)
    dtc_codigo13.BoundText = Dtc_Stock13.BoundText
    dtc_desc13.BoundText = Dtc_Stock13.BoundText
End Sub

Private Sub dtc_desc12_Click(Area As Integer)
    Dtc_aux12.BoundText = dtc_desc12.BoundText
    dtc_codigo12.BoundText = dtc_desc12.BoundText
End Sub

Private Sub dtc_desc12_LostFocus()
'  If GlSistema = "A" Then       'Or GlSistema = "Z"
'    If dtc_codigo12.Text = "10" Then
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    End If
'  Else
'    'If lblventa_tipo.Caption = "E" Then
'    '    TxtPrecioU.Text = dtc_precioventafinal15.Text
'    'Else
'    '    TxtPrecioU.Text = dtc_precioventabase15.Text
'    'End If
'    If Val(dtc_codigo12.Text) > 19 Then
'        TxtPrecioU.Text = dtc_precioventafinal15.Text
'    Else
'        TxtPrecioU.Text = dtc_precioventabase15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 100 Then
'        TxtPrecioU.Text = dtc_preciocompra15.Text
'    End If
'    If Val(dtc_codigo12.Text) = 200 Then
'        TxtPrecioU.Text = "0"
'    End If
'  End If

End Sub

Private Sub dtc_unimed15_Click(Area As Integer)
    dtc_codigo15.BoundText = dtc_unimed15.BoundText
    dtc_desc15.BoundText = dtc_unimed15.BoundText
    dtc_stocktotal15.BoundText = dtc_unimed15.BoundText
    dtc_grupo15.BoundText = dtc_unimed15.BoundText
    dtc_subgrupo15.BoundText = dtc_unimed15.BoundText
    Dtc_partida15.BoundText = dtc_unimed15.BoundText
    dtc_precioventafinal15.BoundText = dtc_unimed15.BoundText
    dtc_precioventabase15.BoundText = dtc_unimed15.BoundText
    dtc_preciocompra15.BoundText = dtc_unimed15.BoundText
End Sub

Private Sub dtc_desc2A_Click(Area As Integer)
    dtc_codigo2A.BoundText = dtc_desc2A.BoundText
    dtc_benef2A.BoundText = dtc_desc2A.BoundText
    dtc_email2A.BoundText = dtc_desc2A.BoundText
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

Private Sub DTPfechasol_LostFocus()
    Set rs_TipoCambio = New ADODB.Recordset
    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
    rs_TipoCambio.Open "select * from gc_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
    If rs_TipoCambio.RecordCount > 0 Then
        txtTDC.Text = rs_TipoCambio!cambio_oficial_compra
    End If
'    Ado_datos4.Refresh
End Sub

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
        Case "1.9"    ' Chuquisaca
            Aux = "DCOMC"
            VAR_DPTO = "1"
        Case "1.3"    ' Modernizacion
            Aux = "DNMOD"
            VAR_DPTO = "2"
        Case "0"    ' TODO
            If glusuario = "ASANTIVAÑEZ" Then
                Aux = "DNMOD"
                VAR_DPTO = "2"
            Else
                Aux = "DVTA"
                VAR_DPTO = "2"
            End If
            'Aux = "DVTA"
            'VAR_DPTO = "2"
     End Select
    parametro = Aux
    db.Execute "UPDATE ao_ventas_cabecera SET codigo_empresa = '2' WHERE venta_tipo = 'G' AND (codigo_empresa IS NULL OR codigo_empresa ='0') "
    db.Execute "UPDATE ao_ventas_cabecera SET codigo_empresa = '1' WHERE venta_tipo <> 'G' AND (codigo_empresa IS NULL OR codigo_empresa ='0') "
    
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    If Ado_datos.Recordset.RecordCount > 0 Then
        nroventa = Ado_datos.Recordset!venta_codigo
    Else
        nroventa = 0
    End If
    Call ABRIR_TABLA_DET
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
    BtnImprimir2.Visible = False
'    BtnImprimir3.Visible = False
'    FrmEdita.Enabled = False
'    FrmCobros.Enabled = False
'    Cmd_Cliente.Visible = False
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
    VAR_NEW = "X"
    Chk_plazo.Value = 0
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset     'UNIDAD EJECUTORA
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset     'Beneficiario Personas Nat. y Juridicas
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_beneficiario where (estado_codigo ='APR' ) order by beneficiario_denominacion", db, adOpenStatic   'and tipoben_codigo <20
    'rs_datos2.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset     'Proyecto de Edificación
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
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
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
    
    'EMPRESA
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_empresas order by codigo_empresa", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText

    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'If parametro = "DNMOD" Then
    '    rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo = 'C'  ", db, adOpenStatic
    'Else
        rs_datos11.Open "select * from ac_tipo_compra_venta where venta_tipo = 'L' or venta_tipo = 'V' or venta_tipo = 'G' ", db, adOpenStatic     'or venta_tipo = 'G'
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
    'ZONAS PILOTO
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenKeyset, adLockReadOnly
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText

    Set rs_datos17 = New ADODB.Recordset
    If rs_datos17.State = 1 Then rs_datos17.Close
    rs_datos17.Open "select * from ac_bienes_grupo", db, adOpenKeyset, adLockReadOnly
    Set ado_datos17.Recordset = rs_datos17
    ado_datos17.Refresh
'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
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
            Call AbreAlmacen
            If (Ado_datos.Recordset("venta_tipo") = "C") Or (Ado_datos.Recordset("venta_tipo") = "V") Or (Ado_datos.Recordset("venta_tipo") = "G") Or (Ado_datos.Recordset("venta_tipo") = "L") Then
                FrmABMDet2.Visible = True
                FrmCobranza.Visible = True
                
            Else
                FrmABMDet2.Visible = False
                FrmCobranza.Visible = False
            End If
        Else
            deta2 = 0
            'TxtMontoBs.Text = 0
            'TxtMontoUs.Text = 0
            'Text2.Text = 0
            FrmABMDet2.Visible = False
            FrmCobranza.Visible = False
        End If
        
        Set rs_datos16 = New ADODB.Recordset
        If rs_datos16.State = 1 Then rs_datos16.Close
        rs_datos16.Open "select * from ao_ventas_cobranza_prog where venta_codigo = '" & nroventa & "'  ", db, adOpenKeyset, adLockOptimistic
        Set Ado_datos16.Recordset = rs_datos16
        Set DtgCobro.DataSource = Ado_datos16.Recordset
        'Ado_datos16.Recordset.Requery
        If Ado_datos16.Recordset.RecordCount > 0 Then
            Ado_datos16.Recordset.Requery
            FrmCobranza.Visible = True
            
            'BtnImprimir2.Visible = True
            'BtnImprimir3.Visible = True
        Else
            FrmCobranza.Visible = False
            'BtnImprimir2.Visible = False
            'BtnImprimir3.Visible = False
        End If
        
        Set rs_datos6 = New ADODB.Recordset
        If rs_datos6.State = 1 Then rs_datos6.Close
        rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & nroventa & "  order by ORDEN ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_datos6.Open "select * from ao_ventas_alcance where venta_codigo= " & nro_licitacion & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText       'order by ORDEN
        Set Ado_datos6.Recordset = rs_datos6
        Set DtgAlcance.DataSource = Ado_datos6.Recordset
        If Ado_datos6.Recordset.RecordCount > 0 Then
            DtgAlcance.Visible = True
        Else
            DtgAlcance.Visible = False
        End If
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
  If TxtConcepto = "" Then
    MsgBox "Debe Registrar ... " + lbl_concepto, vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If TxtOrigenUsd.Text = "" Then
    TxtOrigenUsd.Text = "1"
  End If
  If txtTDC.Text = "" Then
    txtTDC.Text = "6.96"
  End If
  If DTPfechasol.Value = "" Or DTPfechasol.Value = "01/01/1900" Then
    MsgBox "Debe Registrar la Fecha de Venta ... ", vbExclamation, "Atención"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
  
End Sub

Private Sub grabar()
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
  'db.BeginTrans
'       TxtOrigenBs.Text = "0"      'OK
'        TxtOrigenUsd.Text = "0"    'OK
'        TxtAdendaUsd.Text = "0"    'OK
'        TxtAdendaBs.Text = "0"     'OK
'        txtTDC.Text = "0"          'OK
'        TxtMontoUsd.Text =         'OK
'        TxtMontoBs.Text =          'OK
'        TxtCobradoUsd.Text = "0"   'OK
'       TxtCobrado.Text = "0"       'OK
'    TxtBstotalUsd.Text =           'OK
'    TxtBstotal.Text =              'OK
    
       nroventa = Ado_datos.Recordset!venta_codigo
       db.Execute " update ao_ventas_cabecera set venta_tipo = '" & dtc_codigo11.Text & "', venta_fecha= '" & DTPfechasol.Value & "' , venta_fecha_inicio= '" & DTPfechasol.Value & "' , unidad_codigo_ant = '" & Txt_campo2.Text & "' , beneficiario_codigo_resp= '" & dtc_codigo4.Text & "', beneficiario_codigo_cobr= '" & dtc_codigo4.Text & "', beneficiario_codigo= '" & dtc_codigo2.Text & "', venta_descripcion='" & TxtConcepto.Text & "' ,  estado_codigo = 'REG', usr_codigo = '" & glusuario & "', fecha_registro = '" & Format(Date, "dd/mm/yyyy") & "'  where venta_codigo =  " & nroventa & " "
       db.Execute " update ao_ventas_cabecera set venta_monto_origen_dol=" & CDbl(TxtOrigenUsd.Text) & ", venta_monto_origen_bs=" & CDbl(TxtOrigenBs.Text) & ", venta_monto_adenda_bs=" & CDbl(TxtAdendaBs.Text) & ", venta_monto_adenda_dol= " & CDbl(TxtAdendaUsd.Text) & ", venta_monto_total_dol = " & CDbl(TxtMontoUsd.Text) & " , venta_monto_total_bs= " & CDbl(TxtMontoBs.Text) & "  where venta_codigo =  " & nroventa & " "
       db.Execute " update ao_ventas_cabecera set venta_tipo_cambio=" & CDbl(txtTDC.Text) & ", venta_monto_cobrado_dol=" & CDbl(TxtCobradoUsd.Text) & ", venta_monto_cobrado_bs=" & CDbl(TxtCobrado.Text) & ", venta_saldo_p_cobrar_dol= " & CDbl(TxtBstotalUsd.Text) & ", venta_saldo_p_cobrar_bs= " & CDbl(TxtBstotal.Text) & ", codigo_empresa = " & dtc_codigo8.Text & "   where venta_codigo =  " & nroventa & " "
       If VAR_UORIGEN = "DNMOD" Then
            db.Execute " update ao_ventas_cabecera set proceso_codigo = 'TEC', subproceso_codigo= 'TEC-05' , etapa_codigo = 'TEC-05-01' , clasif_codigo= 'TEC', doc_codigo= 'R-313' , poa_codigo= '3.2.7'  where venta_codigo =  " & nroventa & " "
       Else
            db.Execute " update ao_ventas_cabecera set proceso_codigo = 'COM', subproceso_codigo= 'COM-02' , etapa_codigo = 'COM-02-01' , clasif_codigo= 'COM', doc_codigo= 'R-223' , poa_codigo= '3.1.2'  where venta_codigo =  " & nroventa & " "
       End If
       
    'db.CommitTrans
    If Ado_datos.Recordset.RecordCount > 0 Then
       marca1 = Ado_datos.Recordset.Bookmark
       If Ado_datos.Recordset("venta_tipo") = "E" Then
           db.Execute "INSERT INTO ao_ventas_cobranza_prog (venta_codigo, ges_gestion, beneficiario_codigo, beneficiario_codigo_resp, cobranza_deuda_bs, cobranza_deuda_dol, cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs, cobranza_total_dol, cobranza_fecha_prog, cobranza_fecha_cobro, cobranza_observaciones, literal, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac, cobranza_nro_factura, cobranza_nro_autorizacion, factura_impresa, poa_codigo, estado_codigo, usr_codigo, fecha_registro, hora_registro) " & _
           "VALUES ('" & Ado_datos.Recordset!venta_codigo & "', '" & Ado_datos.Recordset!ges_gestion & "', '" & Ado_datos.Recordset!beneficiario_codigo & "', '" & Ado_datos.Recordset!beneficiario_codigo_resp & "', " & Ado_datos.Recordset!venta_monto_total_bs & ", '" & Ado_datos.Recordset!venta_monto_total_dol & "', '0', '0', " & Ado_datos.Recordset!venta_monto_total_bs & ", " & Ado_datos.Recordset!venta_monto_total_dol & ", '" & Date & "', '" & Date & "', 'CANCELADO', 'CERO', 'COM', 'COM-02', 'COM-02-02', 'ADM', 'R-103', '0', 'R-101', '0', '0', 'N', '3.1.2', 'REG', '" & glusuario & "', '" & Date & "', '09:00')"
           '  cobranza_codigo       'Especif. de Identidad
       End If
'       Call OptFilGral1_Click
       'Ado_datos.Refresh
       'Ado_datos.Recordset.Move marca1 - 1
'        If swgrabar = 1 Then
'            Ado_datos.Refresh
'            Ado_datos.Recordset.MoveLast
'        End If
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

'Private Sub OpMod2_Click()
'    Txt_modelo.Text = Txt_modelo2.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and cotiza_codigo = " & ado_datos14.Recordset!cotiza_codigo & " ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtDescuento.Text = "0"
'        TxtPrecioU.Text = IIf(IsNull(rs_datos18!cotiza_fob_seg_dol), 0, rs_datos18!cotiza_fob_seg_dol)
'        'TxtPrecioU.Text = IIf(IsNull(rs_datos18!cotiza_precio_total_bs_h), 0, rs_datos18!cotiza_precio_total_bs_h)
'    End If
'End Sub

'Private Sub OpMod3_Click()
'    Txt_modelo.Text = Txt_modelo3.Text
'    Set rs_datos18 = New ADODB.Recordset
'    If rs_datos18.State = 1 Then rs_datos18.Close
'    rs_datos18.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and cotiza_codigo = " & ado_datos14.Recordset!cotiza_codigo & " ", db, adOpenKeyset, adLockReadOnly
'    If rs_datos18.RecordCount > 0 Then
'        TxtDescuento.Text = "0"
'        TxtPrecioU.Text = IIf(IsNull(rs_datos18!cotiza_fob_seg_dol), 0, rs_datos18!cotiza_fob_seg_dol)
'        'TxtPrecioU.Text = rs_datos18!cotiza_precio_total_bs
'    End If
'End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_aux13 = New ADODB.Recordset
    If rs_aux13.State = 1 Then rs_aux13.Close
    rs_aux13.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux13.RecordCount > 0 Then
        usuario2 = rs_aux13!beneficiario_codigo
        VAR_DA = rs_aux13!da_codigo
        VAR_DPTO = rs_aux13!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos = New Recordset
     If rs_datos.State = 1 Then rs_datos.Close
     Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (unidad_codigo = 'DCOMB' AND left(edif_codigo,1) = '3' ) OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9'  ) )) "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' ) )) "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "GSOLIZ" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND (unidad_codigo = 'DOCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC') )) "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "OCOLODRO" Or glusuario = "JAVIER" Or glusuario = "LNAVA" Then
                queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND unidad_codigo = 'DNMOD' )"
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "'  AND left(edif_codigo,1) = '" & VAR_DPTO & "' ))"      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From av_ventas_cabecera WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "'  AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5' or left(edif_codigo,1) = '6' ) )) "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "RVALDIVIEZOR" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "RCUELA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                Else
                    queryinicial = "select * From av_ventas_cabecera WHERE (estado_codigo = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select

'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "BINFANTE" Or glusuario = "AURBINA" Or glusuario = "GSOLIZ" Or glusuario = "VPAREDES" Or glusuario = "RCUELA" Then
'        queryinicial = "select * From av_ventas_cabecera WHERE unidad_codigo= '" & parametro & "' and estado_codigo = 'REG' "
'    Else
'        queryinicial = "select * From av_ventas_cabecera WHERE unidad_codigo= '" & VAR_UORIGEN & "' and estado_codigo = 'REG' AND usr_codigo = '" & glusuario & "'  "         ' AND beneficiario_codigo_resp = '" & usuario2 & "'"
'    End If
'    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    'rs_datos.Sort = "solicitud_codigo"
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
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
            If (glusuario = "MARTEAGA" Or glusuario = "FCABRERA") Then        'SCZ
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DCOMB' AND left(edif_codigo,1) = '3' ) OR (unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4'  )))  "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' )))  "
            End If
        Case "1.7"    'Santa Cruz
            If (glusuario = "RGIL" Or glusuario = "CPAREDES") Then        'SCZ
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DCOMS' AND left(edif_codigo,1) = '7' ) OR (unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9'  )))  "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' )))  "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "GSOLIZ" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = 'DOCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC'))  "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9' )))  "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "OCOLODRO" Or glusuario = "JAVIER" Or glusuario = "LNAVA" Then
                queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = 'DNMOD') "
            Else
                queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' ))) "      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6')))  "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "RVALDIVIEZOR" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "RCUELA" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From av_ventas_cabecera WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select

'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "BINFANTE" Or glusuario = "AURBINA" Or glusuario = "MVALDIVIA" Or glusuario = "HBUSTILLOS" Or glusuario = "SPAREDES" Or glusuario = "GSOLIZ" Or glusuario = "VPAREDES" Or glusuario = "RCUELA" Then
'        queryinicial = "select * From av_ventas_cabecera WHERE (unidad_codigo = '" & parametro & "' or unidad_codigo ='DCOMS' or unidad_codigo ='DCOMB' or unidad_codigo ='DCOMC') "
'        'queryinicial = "select * From av_ventas_cabecera WHERE unidad_codigo= '" & VAR_UORIGEN & "' "
'    Else
'        queryinicial = "select * From av_ventas_cabecera WHERE unidad_codigo= '" & VAR_UORIGEN & "' AND usr_codigo = '" & glusuario & "'   "
'    End If
'    'queryinicial = "select * From av_ventas_cabecera where unidad_codigo= '" & parametro & "' "
'    'queryinicial = "Select * from ao_solicitud where " + parametro
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
'  Set rstacumdet = New ADODB.Recordset
'  If rstacumdet.State = 1 Then rstacumdet.Close
'  Set rs_datos19 = New ADODB.Recordset
'  If rs_datos19.State = 1 Then rs_datos19.Close
'  Set rs_datos20 = New ADODB.Recordset
'  If rs_datos20.State = 1 Then rs_datos20.Close
''  LblGestion
''  lblcorrelVenta
''  lblNroVenta
'  rstacumdet.Open "select sum(venta_precio_total_bs) as totbs, sum (venta_precio_total_dol) as totdl , sum (venta_det_cantidad) as cantot from ao_ventas_detalle where venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
'  If IsNull(rstacumdet!totbs) Then
'    VAR_AUX = 0
'    VAR_AUX2 = 0
'    VAR_CANT = 1
'  Else
'    VAR_AUX = Round(rstacumdet!totbs, 2)
'    VAR_AUX2 = Round(rstacumdet!totdl, 2)
'    VAR_CANT = rstacumdet!CANTOT
'  End If
'  'ya no detalle Costos
''  rs_datos20.Open "select sum(costo_monto) as totbs4, sum(costo_monto_usd) as totdl5 from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & ges, db, adOpenKeyset, adLockOptimistic   'ges_gestion = '" & ges & "' and
''  If IsNull(rs_datos20!totbs4) Then
''    VAR_AUX4 = 0
''    VAR_AUX5 = 0
''  Else
''    VAR_AUX4 = Round(rs_datos20!totbs4, 2)
''    VAR_AUX5 = Round(rs_datos20!totdl5, 2)
''  End If
'
'  rs_datos19.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza_prog where estado_codigo = 'APR' and venta_codigo = " & Nro, db, adOpenKeyset, adLockOptimistic
'  If IsNull(rs_datos19!totbs2) Then
'    Cobrobs = 0
'    VAR_COBR = 0
'  Else
'    Cobrobs = Round(rs_datos19!totbs2, 2)
'    VAR_COBR = Round(rs_datos19!totdl2, 2)
'  End If
'
'  'VAR_Bs = VAR_AUX + VAR_AUX4 - Cobrobs
'  'VAR_Dol = VAR_AUX2 + VAR_AUX5 - VAR_COBR
'  VAR_Bs = VAR_AUX - Cobrobs
'  VAR_Dol = VAR_AUX2 - VAR_COBR
'  'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX + VAR_AUX4 & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 + VAR_AUX5 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "       'ao_ventas_cabecera.ges_gestion = '" & ges & "' And
'  db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.venta_monto_total_bs = " & VAR_AUX & " , ao_ventas_cabecera.venta_monto_total_dol = " & VAR_AUX2 & ", ao_ventas_cabecera.venta_cantidad_total = " & VAR_CANT & ", ao_ventas_cabecera.venta_monto_cobrado_bs = " & Cobrobs & ", ao_ventas_cabecera.venta_monto_cobrado_dol = " & VAR_COBR & ",  ao_ventas_cabecera.venta_saldo_p_cobrar_bs = " & VAR_Bs & ", ao_ventas_cabecera.venta_saldo_p_cobrar_dol = " & VAR_Dol & "  Where ao_ventas_cabecera.venta_codigo = " & Nro & " "
'
'  TxtMontoBs.Text = VAR_AUX '+ VAR_AUX4
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
'    If TxtMontoUsd.Text = "" Or TxtMontoUsd.Text = "0" Or TxtMontoUsd.Text = "0.00" Then
'        TxtMontoBs.Text = "0"
'        TxtMontoUsd.Text = "0"
'        TxtBstotalUsd = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd)
'    Else
'        TxtMontoBs.Text = Round(CDbl(TxtMontoUsd.Text) * GlTipoCambioMercado, 2)
'    End If
'    TxtBstotalUsd.Text = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd)
'    TxtBstotal.Text = CDbl(TxtMontoBs) - CDbl(TxtCobrado)
End Sub

Private Sub TxtOrigenUsd_LostFocus()
    If TxtOrigenUsd.Text = "" Or TxtOrigenUsd.Text = "0" Or TxtOrigenUsd.Text = "0.00" Then
        TxtOrigenBs.Text = "0"
        TxtOrigenUsd.Text = "0"
        TxtAdendaUsd.Text = "0"
        TxtAdendaBs.Text = "0"
        txtTDC.Text = "0"
        TxtMontoUsd.Text = CDbl(TxtOrigenUsd.Text) + CDbl(TxtAdendaUsd.Text)
        TxtMontoBs.Text = CDbl(TxtOrigenBs.Text) + CDbl(TxtAdendaBs.Text)
    Else
        If txtTDC.Text = "" Or txtTDC.Text = "0" Or txtTDC.Text = "0.00" Or txtTDC.Text = "1" Then
            txtTDC.Text = "6.96"
        End If
        TxtOrigenBs.Text = Round(CDbl(TxtOrigenUsd.Text) * CDbl(txtTDC.Text), 2)
        TxtAdendaUsd.Text = "0"
        TxtAdendaBs.Text = "0"
        TxtMontoUsd.Text = Round(CDbl(TxtOrigenUsd.Text) + CDbl(TxtAdendaUsd.Text), 2)
        TxtMontoBs.Text = Round(CDbl(TxtOrigenBs.Text) + CDbl(TxtAdendaBs.Text), 2)
    End If
    'If TxtMontoUsd.Text = "" Or TxtMontoUsd.Text = "0" Or TxtMontoUsd.Text = "0.00" Then
    '    TxtMontoBs.Text = "0"
    '    TxtMontoUsd.Text = "0"
    '    TxtBstotalUsd = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd)
    'Else
    '    TxtMontoBs.Text = Round(CDbl(TxtMontoUsd.Text) * GlTipoCambioMercado, 2)
    'End If
    TxtCobradoUsd.Text = "0"
    TxtCobrado.Text = "0"
    TxtBstotalUsd.Text = CDbl(TxtMontoUsd) - CDbl(TxtCobradoUsd.Text)
    TxtBstotal.Text = CDbl(TxtMontoBs) - CDbl(TxtCobrado.Text)
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

Private Sub txtTDC_LostFocus()
    If TxtOrigenUsd.Text = "" Or TxtOrigenUsd.Text = "0" Or TxtOrigenUsd.Text = "0.00" Then
        TxtOrigenBs.Text = "0"
        TxtOrigenUsd.Text = "0"
        TxtAdendaUsd.Text = "0"
        TxtAdendaBs.Text = "0"
        txtTDC.Text = "0"
        TxtMontoUsd.Text = CDbl(TxtOrigenUsd.Text) + CDbl(TxtAdendaUsd.Text)
        TxtMontoBs.Text = CDbl(TxtOrigenBs.Text) + CDbl(TxtAdendaBs.Text)
    Else
        If txtTDC.Text = "" Or txtTDC.Text = "0" Or txtTDC.Text = "0.00" Or txtTDC.Text = "1" Then
            txtTDC.Text = "6.96"
        End If
        TxtOrigenBs.Text = Round(CDbl(TxtOrigenUsd.Text) * CDbl(txtTDC.Text), 2)
        TxtAdendaUsd.Text = "0"
        TxtAdendaBs.Text = "0"
        TxtMontoUsd.Text = Round(CDbl(TxtOrigenUsd.Text) + CDbl(TxtAdendaUsd.Text), 2)
        TxtMontoBs.Text = Round(CDbl(TxtOrigenBs.Text) + CDbl(TxtAdendaBs.Text), 2)
    End If
End Sub

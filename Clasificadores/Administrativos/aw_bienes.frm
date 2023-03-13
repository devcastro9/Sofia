VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_bienes 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - Administrativos -  Bienes y/o Servicios"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   120
   ClientWidth     =   11145
   Icon            =   "aw_bienes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Fra_aux1 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00FFFFFF&
      Height          =   1320
      Left            =   15840
      ScaleHeight     =   1260
      ScaleWidth      =   9420
      TabIndex        =   115
      Top             =   6600
      Visible         =   0   'False
      Width           =   9480
      Begin VB.TextBox dtc_codigo4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "marca_codigo"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "-"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BtnGrabar4 
         BackColor       =   &H80000010&
         Height          =   675
         Left            =   8040
         Picture         =   "aw_bienes.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   315
         Width           =   1245
      End
      Begin VB.CommandButton BtnCancelar4 
         BackColor       =   &H80000010&
         Height          =   675
         Left            =   6705
         MaskColor       =   &H00000000&
         Picture         =   "aw_bienes.frx":11D8
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Cancelar"
         Top             =   315
         Width           =   1365
      End
      Begin VB.TextBox Txt_descripcion4 
         DataField       =   "calle_denominacion"
         Height          =   285
         Left            =   1680
         TabIndex        =   116
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lbl_enlace11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Marca"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   120
         Top             =   105
         Width           =   1275
      End
      Begin VB.Label lbl_descripcion4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación de la Marca"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1680
         TabIndex        =   119
         Top             =   105
         Width           =   2895
      End
   End
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   2880
      TabIndex        =   106
      Top             =   2760
      Visible         =   0   'False
      Width           =   6135
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "Fecha_Alerta"
         Height          =   310
         Left            =   3120
         TabIndex        =   111
         Top             =   1320
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   156762113
         CurrentDate     =   40245
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   107
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "aw_bienes.frx":1AC4
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   109
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "aw_bienes.frx":23B0
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   108
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
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
            TabIndex        =   110
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE CORTE:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1320
         TabIndex        =   112
         Top             =   1320
         Width           =   1800
      End
   End
   Begin VB.Frame fra_imprimir 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   3375
      Left            =   2880
      TabIndex        =   92
      Top             =   1800
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame fra_reportes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Elija una de las opciones ..."
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   120
         TabIndex        =   97
         Top             =   960
         Width           =   5895
         Begin VB.OptionButton rb_servicios 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SERVICIOS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   103
            Top             =   1560
            Width           =   2175
         End
         Begin VB.OptionButton rb_repuestos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "REPUESTOS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   101
            Top             =   840
            Width           =   1455
         End
         Begin VB.OptionButton rb_insumos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "INSUMOS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   100
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton rb_herramientas 
            BackColor       =   &H00C0C0C0&
            Caption         =   "HERRAMIENTAS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   99
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton rb_otros 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OTROS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   98
            Top             =   1920
            Width           =   2175
         End
         Begin Crystal.CrystalReport Cr_herramientas 
            Left            =   120
            Top             =   1320
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
         Begin Crystal.CrystalReport Cr_insumos 
            Left            =   120
            Top             =   360
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
         Begin Crystal.CrystalReport Cr_repuestos 
            Left            =   120
            Top             =   840
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
         Begin Crystal.CrystalReport Cr_otros 
            Left            =   120
            Top             =   1800
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
      Begin VB.PictureBox Picture35 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5880
         TabIndex        =   93
         Top             =   240
         Width           =   5880
         Begin VB.PictureBox BtnCancelar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "aw_bienes.frx":2C7D
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   95
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "aw_bienes.frx":3569
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   94
            ToolTipText     =   "Imprimir el Listado de los Registros"
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label30 
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
            TabIndex        =   96
            Top             =   195
            Width           =   1005
         End
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   81
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7080
         Picture         =   "aw_bienes.frx":3E36
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   105
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   15600
         Picture         =   "aw_bienes.frx":4703
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "aw_bienes.frx":4B45
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   89
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "aw_bienes.frx":5304
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   88
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
         Picture         =   "aw_bienes.frx":5C19
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   87
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8640
         Picture         =   "aw_bienes.frx":6365
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   86
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "aw_bienes.frx":6B98
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   85
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
         Picture         =   "aw_bienes.frx":734D
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   84
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
         Picture         =   "aw_bienes.frx":7C1A
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   83
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   14520
         Picture         =   "aw_bienes.frx":83DC
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
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
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12720
         TabIndex        =   91
         Top             =   195
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
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20280
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4680
         Picture         =   "aw_bienes.frx":85E6
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   104
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "aw_bienes.frx":8DBC
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   79
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTITULO"
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
         Left            =   13005
         TabIndex        =   80
         Top             =   195
         Width           =   1425
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00000000&
      Height          =   8040
      Left            =   120
      TabIndex        =   51
      Top             =   720
      Width           =   5895
      Begin VB.OptionButton OptFilGral2 
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
         TabIndex        =   20
         Top             =   7665
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
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
         TabIndex        =   19
         Top             =   7665
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "aw_bienes.frx":96A8
         Height          =   7290
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   12859
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ForeColor       =   -2147483630
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Código"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripción"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "grupo_codigo"
            Caption         =   "Grupo"
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
            DataField       =   "subgrupo_codigo"
            Caption         =   "SubGrupo"
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
            DataField       =   "par_codigo"
            Caption         =   "Partida"
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
            DataField       =   "marca_codigo"
            Caption         =   "Marca"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "almacen_tipo"
            Caption         =   "Tipo.Almacen"
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
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3479,811
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   585,071
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7605
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
         BackColor       =   -2147483633
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
   Begin VB.Frame FraArticulos 
      BackColor       =   &H00C0C0C0&
      Height          =   8040
      Left            =   6120
      TabIndex        =   25
      Top             =   720
      Width           =   9525
      Begin VB.TextBox TxtDescripcionSIN 
         DataField       =   "descripcion_pSIN"
         DataSource      =   "Ado_datos"
         Height          =   375
         Left            =   4920
         TabIndex        =   127
         Top             =   7080
         Width           =   4455
      End
      Begin VB.TextBox TxtCodigo_pSIN 
         DataField       =   "codigo_pSIN"
         DataSource      =   "Ado_datos"
         Height          =   375
         Left            =   1680
         TabIndex        =   124
         Top             =   7080
         Width           =   1575
      End
      Begin VB.CommandButton BtnAux1 
         Height          =   320
         Left            =   9120
         Picture         =   "aw_bienes.frx":96C0
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Nueva Marca"
         Top             =   3915
         Width           =   375
      End
      Begin VB.TextBox TxtPrecEstD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_final_dol"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   77
         Text            =   "0.00"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox TxtPrecVentaD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_base_dol"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox TxtPrecCompD 
         Alignment       =   2  'Center
         DataField       =   "bien_precio_compra_dol"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox txtStockIni 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataField       =   "bien_stock_inicial"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   5220
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cmd_rotacion 
         DataField       =   "bien_rotacion"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "aw_bienes.frx":A0C2
         Left            =   8040
         List            =   "aw_bienes.frx":A0CF
         TabIndex        =   12
         Text            =   "ALTA"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   5750
         TabIndex        =   65
         Top             =   4575
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_partida 
         Bindings        =   "aw_bienes.frx":A0E9
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   8421504
         ListField       =   "par_codigo"
         BoundColumn     =   "par_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "aw_bienes.frx":A102
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2445
         TabIndex        =   18
         Top             =   6720
         Visible         =   0   'False
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox TxtPrecVenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_base"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   5860
         Width           =   1215
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2445
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2715
         Width           =   6825
      End
      Begin VB.PictureBox Img_Foto 
         Height          =   1695
         Left            =   7180
         ScaleHeight     =   1635
         ScaleWidth      =   2115
         TabIndex        =   48
         Top             =   240
         Width           =   2175
         Begin VB.Image Image2 
            Height          =   1640
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2120
         End
      End
      Begin VB.TextBox TxtDescripcion2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_descripcion_anterior"
         DataSource      =   "Ado_datos"
         Height          =   405
         Left            =   1845
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3135
         Width           =   7425
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Fecha_Vencimiento"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   450
         _Version        =   393216
         Format          =   156696577
         CurrentDate     =   40245
      End
      Begin VB.TextBox TxtPrecEst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_final"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4560
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   5860
         Width           =   1215
      End
      Begin VB.TextBox txtStockMin 
         Alignment       =   2  'Center
         DataField       =   "bien_stock_minimo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   5220
         Width           =   1245
      End
      Begin VB.TextBox TxtPrecComp 
         Alignment       =   2  'Center
         DataField       =   "bien_precio_compra"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   5880
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo Montador 
         Bindings        =   "aw_bienes.frx":A11C
         DataField       =   "subgrupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   8421504
         ListField       =   "subgrupo_codigo"
         BoundColumn     =   "subgrupo_codigo"
         Text            =   "Elige Marca..."
      End
      Begin VB.TextBox TxtInicial 
         Alignment       =   2  'Center
         DataField       =   "bien_codigo_anterior"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   6165
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "0"
         Top             =   4560
         Width           =   1605
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Enabled         =   0   'False
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
         Left            =   160
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "12345678901234567890"
         Top             =   2715
         Width           =   2295
      End
      Begin MSDataListLib.DataCombo TDBC_marcas 
         Bindings        =   "aw_bienes.frx":A137
         DataField       =   "marca_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5880
         TabIndex        =   9
         Top             =   3915
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "marca_descripcion"
         BoundColumn     =   "marca_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo marcas 
         Bindings        =   "aw_bienes.frx":A14E
         DataField       =   "marca_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8040
         TabIndex        =   32
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "marca_codigo"
         BoundColumn     =   "marca_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo TDBC_Montador 
         Bindings        =   "aw_bienes.frx":A165
         DataField       =   "subgrupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "subgrupo_descripcion"
         BoundColumn     =   "subgrupo_codigo"
         Text            =   "Elige Marca..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo TDBC_Unidad 
         Bindings        =   "aw_bienes.frx":A17F
         DataField       =   "unimed_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   165
         TabIndex        =   7
         Top             =   3915
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "unimed_descripcion"
         BoundColumn     =   "unimed_codigo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo Unidad 
         Bindings        =   "aw_bienes.frx":A197
         DataField       =   "unimed_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   34
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "unimed_codigo"
         BoundColumn     =   "unimed_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoCod 
         Bindings        =   "aw_bienes.frx":A1AF
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   40
         Top             =   160
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   8421504
         ListField       =   "grupo_codigo"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoDes 
         Bindings        =   "aw_bienes.frx":A1C6
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "grupo_descripcion"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Fecha_Alerta"
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   6480
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         Format          =   156696577
         CurrentDate     =   40245
      End
      Begin MSDataListLib.DataCombo DtcPaisD 
         Bindings        =   "aw_bienes.frx":A1DD
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2655
         TabIndex        =   8
         Top             =   3915
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "pais_descripcion"
         BoundColumn     =   "pais_codigo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo DtcPais 
         Bindings        =   "aw_bienes.frx":A1F3
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   46
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "pais_codigo"
         BoundColumn     =   "pais_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoUni 
         Bindings        =   "aw_bienes.frx":A209
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   8421504
         ListField       =   "unidad_codigo"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "aw_bienes.frx":A220
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   960
         TabIndex        =   57
         Top             =   6720
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_partida_des 
         Bindings        =   "aw_bienes.frx":A23A
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   360
         Left            =   1635
         TabIndex        =   3
         Top             =   2025
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "par_descripcion"
         BoundColumn     =   "par_codigo"
         Text            =   "Elige Marca..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "aw_bienes.frx":A253
         DataField       =   "modelo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   60
         Top             =   4560
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "modelo_descripcion"
         BoundColumn     =   "modelo_codigo"
         Text            =   "Elige Modelo ..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "aw_bienes.frx":A26C
         DataField       =   "modelo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   165
         TabIndex        =   10
         Top             =   4560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "modelo_codigo"
         BoundColumn     =   "modelo_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "aw_bienes.frx":A285
         DataField       =   "bien_codigo_universal"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8400
         TabIndex        =   67
         Top             =   6480
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "tipo_eqp"
         BoundColumn     =   "tipo_eqp"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "aw_bienes.frx":A29E
         DataField       =   "bien_codigo_universal"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   17
         Top             =   6240
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "tipo_eqp_descripcion"
         BoundColumn     =   "tipo_eqp"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "aw_bienes.frx":A2B7
         DataField       =   "almacen_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6240
         TabIndex        =   113
         Top             =   5760
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "almacen_tipo_descripcion"
         BoundColumn     =   "almacen_tipo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "aw_bienes.frx":A2D0
         DataField       =   "almacen_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5880
         TabIndex        =   114
         Top             =   5760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   -2147483624
         ListField       =   "almacen_tipo"
         BoundColumn     =   "almacen_tipo"
         Text            =   "Elige Marca..."
      End
      Begin VB.Label LblProductoSIN 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto SIN"
         Height          =   255
         Left            =   360
         TabIndex        =   128
         Top             =   7680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LblDescripcionSIN 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion del Producto para SIN"
         Height          =   375
         Left            =   3480
         TabIndex        =   126
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label LblCodigoPSin 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo para SIN"
         Height          =   255
         Left            =   240
         TabIndex        =   125
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label txtCantVendida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_salida"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6000
         TabIndex        =   102
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4080
         TabIndex        =   75
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2160
         TabIndex        =   74
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4080
         TabIndex        =   73
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2160
         TabIndex        =   72
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   71
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   70
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Iinicial"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   69
         Top             =   4980
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rotación"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8040
         TabIndex        =   68
         Top             =   4320
         Width           =   1125
      End
      Begin VB.Label lbl_eqp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Equipo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   66
         Top             =   6000
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB GRUPO 2"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   64
         Top             =   1755
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB GRUPO 1"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   63
         Top             =   1000
         Width           =   1350
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   62
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   61
         Top             =   4320
         Width           =   690
      End
      Begin VB.Label txt_par 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "par_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   165
         TabIndex        =   59
         Top             =   2025
         Width           =   1440
      End
      Begin VB.Label lbl_edif 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   56
         Top             =   6720
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label TxtActual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_actual"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7920
         TabIndex        =   55
         Top             =   5220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label txtCantComprada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_ingreso"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3960
         TabIndex        =   54
         Top             =   5220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label TxtSub 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "subgrupo_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   165
         TabIndex        =   53
         Top             =   1320
         Width           =   1120
      End
      Begin VB.Label TxtGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "grupo_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   160
         TabIndex        =   52
         Top             =   480
         Width           =   880
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Mínimo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   50
         Top             =   4980
         Width           =   1200
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Características Complementarias"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   165
         TabIndex        =   47
         Top             =   3075
         Width           =   1740
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Industria (Pais Origen)"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2655
         TabIndex        =   45
         Top             =   3660
         Width           =   1965
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Almacen:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6000
         TabIndex        =   44
         Top             =   5520
         Width           =   1590
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Primer Vencimiento:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   6480
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Total Salida"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5895
         TabIndex        =   42
         Top             =   4980
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Total Ingreso"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3810
         TabIndex        =   41
         Top             =   4980
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Venta.Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4125
         TabIndex        =   39
         Top             =   5595
         Width           =   1845
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Referencia"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6165
         TabIndex        =   38
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label TDBFrame3D6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Actual"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7920
         TabIndex        =   37
         Top             =   4980
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label TDBFrame3D7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Compra.Base"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   135
         TabIndex        =   36
         Top             =   5595
         Width           =   1905
      End
      Begin VB.Label TDBFrame3D8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Venta.Base"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   35
         Top             =   5600
         Width           =   1725
      End
      Begin VB.Label TDBFrame3D5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad de Medida"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   31
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label TDBFrame3D9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5925
         TabIndex        =   30
         Top             =   3660
         Width           =   570
      End
      Begin VB.Label TDBFrame3D1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "  CODIGO                           NOMBRE O DESCRIPCION DEL BIEN O SERVICIO"
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
         Height          =   255
         Left            =   45
         TabIndex        =   29
         Top             =   2460
         Width           =   9450
      End
   End
   Begin MSAdodcLib.Adodc AdoMontador 
      Height          =   375
      Left            =   6960
      Top             =   8880
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   661
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
      Caption         =   "AdoSubGrp"
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
   Begin MSAdodcLib.Adodc AdoMedida 
      Height          =   375
      Left            =   9480
      Top             =   8880
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
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
      Caption         =   "medida"
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   375
      Left            =   11880
      Top             =   8880
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "marca"
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
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   22920
      TabIndex        =   26
      Top             =   12015
      Width           =   22920
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   15
         TabIndex        =   27
         Top             =   255
         Width           =   12570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   12840
         TabIndex        =   28
         Top             =   75
         Width           =   1845
      End
   End
   Begin VB.PictureBox imlMaterial 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   4200
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   123
      Top             =   4200
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   375
      Left            =   4680
      Top             =   8880
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "AdoPais"
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   375
      Left            =   2400
      Top             =   8880
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "AdoGrupo"
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
   Begin Crystal.CrystalReport CryLista 
      Left            =   120
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryBBSS 
      Left            =   600
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryFis 
      Left            =   1080
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   120
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   120
      Top             =   9360
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
      Left            =   2400
      Top             =   9360
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
      Left            =   4680
      Top             =   9360
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
   Begin Crystal.CrystalReport CryV01 
      Left            =   120
      Top             =   9600
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   6960
      Top             =   9360
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
End
Attribute VB_Name = "aw_bienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim rsMarcas As ADODB.Recordset
Dim rsUnidad As ADODB.Recordset
Dim rsMontador As ADODB.Recordset

Dim rsgrupo As ADODB.Recordset
Dim RsArt, rsPais As ADODB.Recordset
Dim rsNada As ADODB.Recordset
Dim rs_datos10 As ADODB.Recordset
Dim rs_datos6 As ADODB.Recordset
Dim rs_datos7 As ADODB.Recordset
Dim rs_datos8 As ADODB.Recordset

Dim rs_aux1, rs_aux2, rs_aux3 As ADODB.Recordset
Dim rs_aux4 As ADODB.Recordset
Dim rs_aux5 As ADODB.Recordset
Dim rs_aux6 As ADODB.Recordset
'--------
Dim estado, VAR_CONT As Integer ' 0 navegar, 1 Agregar, 2 Editar
Dim SQL_FOR As String
Dim swnuevo As Boolean

Dim sino As String
Dim NombreCarpeta, e As String
Dim CodBien, COD_EDIF, COD_MOD As String
Dim VAR_INICIAL As String
Dim VAR_DET As String

Dim marca1 As BookmarkEnum

Dim VAR_Dol As Double
Dim C_FIJO As Double
Dim C_MANOBR As Double
Dim C_GTOADM As Double
Dim C_UTILID As Double
Dim C_ROTALT As Double
Dim C_ROTBAJ As Double
Dim C_IMPSTO As Double
Dim C_IMPSTO2 As Double
Dim VAR_INSUMO As String
Dim VAR_VAL As String
'--
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
'Dim queryinicial As String

Public Sub ALPrincipal(QEstado As Integer)
    '
'    Screen.MousePointer = vbHourglass
'    estado = QEstado
'    '
'    Select Case estado
'        Case 0
'            Set RsArt = New ADODB.Recordset
'            'JQA 04/2008
'            'GlSqlAux = "SELECT * FROM ac_bienes WHERE CAST(grupo_codigo AS INT)< 50  AND bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY CAST (grupo_codigo AS INT)"
'            'GlSqlAux = "SELECT * FROM ac_bienes WHERE bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY grupo_codigo, subgrupo_codigo, bien_codigo "
'            queryinicial = "SELECT * FROM ac_bienes WHERE bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion "
'            RsArt.Open queryinicial, db, adOpenDynamic, adLockOptimistic
'            If RsArt.RecordCount > 0 Then
'               GlHayRegs = True  'Variable global
'            Else
'               GlHayRegs = False
'            End If
'            BotonesNavegar Me
'            FraArticulos.Enabled = False
'            Set Ado_datos.Recordset = RsArt
'        Case 1
'
'        Case 2
'
'    End Select4273257
'    '
'    Screen.MousePointer = vbDefault
'    Me.Show
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Dim Marca As String
'Dim a As Integer
'Dim COD_MARCAx, cod_UMedida As String
If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
        If Ado_datos.Recordset.BOF And Ado_datos.Recordset.EOF Then
            TxtGrupo.Caption = ""
            TxtDetalle.Text = ""
            txtDescripcion.Text = ""
            TxtActual.Caption = ""
'            chkEstado.Value = vbUnchecked
'            Ado_datos.Caption = "Registro: 0 de 0"
'            BuscaNodo "Grupo"          'Case "ADMIN", "SLIMACHI", "JYMAMANI"   'TODOS
            If glusuario = "CARIZACA" Or glusuario = "ADMIN" Or glusuario = "TCRUZ" Or glusuario = "AFLORES" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Then
                BtnAñadir.Visible = True
                BtnModificar.Visible = True
                btnEliminar.Visible = True
                BtnAprobar.Visible = True
            Else
                BtnAñadir.Visible = False
                BtnModificar.Visible = False
                btnEliminar.Visible = False
                BtnAprobar.Visible = False
            End If
        Else
            Exit Sub
        End If
Else
'    If swnuevo = False Then
'            If Not (IsNull(AdoMarca.Recordset("cod_marca"))) Then
'                If Not (AdoMarca.Recordset.BOF) Then AdoMarca.Recordset.MoveFirst
'                AdoMarca.Recordset.Find "cod_marca ='" & Ado_datos.Recordset!COD_MARCA & "'", , adSearchForward
'                If Not AdoMarca.Recordset.EOF Then
'                    'TDBC_marcas.Item(1) = AdoMarca.Recordset!descripcion
'                    TDBC_marcas.Text = AdoMarca.Recordset!descripcion
'                End If
'            End If
'            If Not (IsNull(AdoMontador.Recordset("subgrupo_codigo"))) Then
'                If Not (AdoMontador.Recordset.BOF) Then AdoMontador.Recordset.MoveFirst
'                AdoMontador.Recordset.Find "subgrupo_codigo ='" & Ado_datos.Recordset!subgrupo_codigo & "'", , adSearchForward
'                If Not AdoMontador.Recordset.EOF Then
'                    '
'                End If
'            End If
'            If Not (IsNull(AdoMedida.Recordset("Unidad"))) Then
'                If Not (AdoMedida.Recordset.BOF) Then AdoMedida.Recordset.MoveFirst
'                    AdoMedida.Recordset.Find "Unidad ='" & Ado_datos.Recordset!Unidad & "'", , adSearchForward
'                If Not AdoMedida.Recordset.EOF Then
'                    '
'                End If
'            End If
            'If glusuario = "CARIZACA" Or glusuario = "LCRUZ" Or glusuario = "RPALLY" Or glusuario = "ADMIN" Or glusuario = "TCRUZ" Then
            If glusuario = "CARIZACA" Or glusuario = "ADMIN" Or glusuario = "AFLORES" Or glusuario = "RCUELA" Or glusuario = "CSALINAS" Then
                BtnAñadir.Visible = True
                BtnModificar.Visible = True
                btnEliminar.Visible = True
                BtnAprobar.Visible = True
            Else
                BtnAñadir.Visible = False
                BtnModificar.Visible = False
                btnEliminar.Visible = False
                BtnAprobar.Visible = False
            End If
        If Ado_datos.Recordset!bien_stock_minimo < Ado_datos.Recordset!bien_stock_actual Then
            TxtActual.backColor = &HE0E0E0
        Else
            TxtActual.backColor = &HFF&
        End If
'    End If
        'TDBC_Montador
    Set Img_Foto = Leer_Imagen(db, "Select Foto From ac_bienes_foto Where bien_codigo = '" & Ado_datos.Recordset("bien_codigo") & "' ", "Foto")
    Image2 = Img_Foto
    If Ado_datos.Recordset!estado_codigo = "APR" Then
        'chkEstado.Value = vbChecked
'        BtnVer.Visible = True
    Else
'        BtnVer.Visible = False
        'chkEstado.Value = vbUnchecked
    End If
    If Ado_datos.Recordset!subgrupo_codigo = "43000" Then
        dtc_codigo10.Visible = True
        dtc_desc10.Visible = True
        lbl_edif.Visible = True
        'dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        lbl_eqp.Visible = True
    Else
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        lbl_edif.Visible = False
        'dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        lbl_eqp.Visible = False
    End If
        'chkEstado.Value =IIf(CBool(Ado_datos.Recordset!estado), vbChecked, vbUnchecked)
'        BuscaNodo Ado_datos.Recordset!grupo_codigo
    
End If
End Sub

Private Sub BtnAñadir_Click()
On Error GoTo Que_Error
    swnuevo = True
    Set dg_datos.DataSource = rsNada
    Ado_datos.Recordset.AddNew
    estado = 1
'    BotonesEditar Me
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    FraArticulos.Enabled = True
'    TxtGrupo.Enabled = False
    DtcGrupoDes.Enabled = True
'    TxtSub.Enabled = False
    TDBC_Montador.Enabled = False
'    trv.SetFocus
'    BuscaNodo "grupo"
    txtStockMin.Text = 0
       TDBC_Montador.Enabled = False
   dtc_partida_des.Enabled = False
   DtcPais.Text = "BOL"
   DtcPaisD.BoundText = DtcPais.BoundText
   marcas.Text = "S/M"
   TDBC_marcas.BoundText = marcas.BoundText
   dtc_codigo6.Text = "S/M"
   dtc_desc6.BoundText = dtc_codigo6.BoundText
   TxtDetalle.Enabled = True
   TxtInicial.Text = "0"
   txtStockMin = "0"
   txtStockIni = "0"
   txtCantComprada = "0"
   txtCantVendida = "0"
   TxtActual = "0"
   TxtPrecComp = "0"
   TxtPrecVenta = "0"
   TxtPrecEst = "0"
   TxtPrecCompD = "0"
   TxtPrecVentaD = "0"
   TxtPrecEstD = "0"
   Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnAprobar_Click()
   If Valida Then
       sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If Ado_datos.Recordset("estado_codigo") = "REG" Then
          If sino = vbYes Then
            CodBien = Ado_datos.Recordset!bien_codigo
            COD_EDIF = Ado_datos.Recordset!edif_codigo
            COD_MOD = Ado_datos.Recordset!modelo_codigo
'            If Ado_datos.Recordset!grupo_codigo = "40000" Then
'                Call ACTUALIZA_ID
'                Call ACTUALIZA_VENTA
'            End If
    '        Dim RUTA1, RUTA2 As String
    '        RUTA1 = "BIENES" + "\" + Trim(adoLista.Recordset("iniciales")) + "-" + Trim(adoLista.Recordset("codigo_beneficiario"))
    '        MsgBox RUTA1
    '        MkDir RUTA1
    '        MkDir RUTA1 + "\CONTRATOS"
    '        MkDir RUTA1 + "\FINIQUITO"
    '        MkDir RUTA1 + "\MEMORANDUMS"
    '        MkDir RUTA1 + "\RESPALDOS"
    '        MkDir RUTA1 + "\HOJA_VIDA"
    '        MkDir RUTA1 + "\OTROS"
    '        MkDir RUTA1 + "\EVALUACIONES"
    '        MkDir RUTA1 + "\LICENCIAS"
    '        MkDir RUTA1 + "\VACACIONES"
            'Ado_datos.Recordset("estado") = 1
            db.Execute "update ac_bienes set estado_codigo = 'APR' WHERE bien_codigo = '" & CodBien & "'  "
'            Ado_datos.Recordset("estado_codigo") = "APR"
'            Ado_datos.Recordset("fecha_registro") = Date
'            Ado_datos.Recordset("usr_codigo") = glusuario
'            Ado_datos.Recordset.Update
            If OptFilGral2.Value = True Then
                Call OptFilGral2_Click
            Else
                Call OptFilGral1_Click
            End If
          End If
       Else
            MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
       End If
   Else
        MsgBox "Existe un error en los datos registrados, Verifique y vuelva a intentar...", vbExclamation + vbOKOnly, "Atención"
   End If
End Sub

Private Sub ACTUALIZA_ID()
    'wwwwwwwwwwwwwwwwwwwwwwwwwwww
    'ACTUALIZA EQUIPOS
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ao_solicitud where edif_codigo = '" & COD_EDIF & "'   ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux1.RecordCount > 0 Then
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        'Id. CLIENTE "36NO" EXISTENTE
        rs_aux3.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND bien_codigo = '" & CodBien & "' ", db, adOpenStatic
        If rs_aux1.RecordCount > 0 Then
            db.Execute "update ao_solicitud_bienes set modelo_codigo = '" & COD_MOD & "' WHERE bien_codigo = '" & CodBien & "' AND unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & " "
        Else
            VAR_CONT = 1
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            'Id. CLIENTE "36NO" NUEVO
            rs_aux2.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND grupo_codigo = '90000' ", db, adOpenStatic
            db.Execute "INSERT INTO ao_solicitud_bienes (ges_gestion, unidad_codigo, solicitud_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, marca_codigo, modelo_codigo, bien_cantidad, bien_precio_compra, bien_total_compra, bien_precio_venta_base, bien_total_venta, tipo_moneda, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, venta_o_compra, fosa_dimension_frente, fosa_dimension_fondo, estado_codigo, usr_codigo, fecha_registro ) VALUES ('" & glGestion & "', '" & rs_aux1!unidad_codigo & "',  " & rs_aux1!solicitud_codigo & ", '" & CodBien & "', '40000', '43000', '43340', '" & Ado_datos.Recordset!marca_codigo & "', '" & COD_MOD & "', " & rs_aux2!bien_cantidad & ", 0, 0, " & rs_aux2!bien_precio_venta_base & ", " & rs_aux2!bien_total_venta & ", 'BOB', '" & rs_aux2!unimed_codigo & "', '" & rs_aux2!unimed_codigo & "', " & rs_aux2!bien_cantidad & ", 'V', 0, 0, 'APR', '" & glusuario & "', '" & Date & "')"
    
            If rs_aux2!bien_codigo = "NA1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              If rs_aux2.RecordCount > 1 Then
                  db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                  If rs_aux2.RecordCount > 2 Then
                      db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                      If rs_aux2.RecordCount > 3 Then
                          db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA4' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                          If rs_aux2.RecordCount > 4 Then
                              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA5' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                              If rs_aux2.RecordCount > 5 Then
                                  db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA6' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                                  If rs_aux2.RecordCount > 6 Then
                                      db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA7' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                                  End If
                              End If
                          End If
                      End If
                  End If
              End If
            End If
            If rs_aux2!bien_codigo = "NE1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
            End If
            If rs_aux2!bien_codigo = "NP1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
            End If
        End If
    End If
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub ACTUALIZA_VENTA()
   'VENTAS
   Set rs_aux1 = New ADODB.Recordset
   If rs_aux1.State = 1 Then rs_aux1.Close
   rs_aux1.Open "select * from ao_ventas_cabecera where edif_codigo = '" & COD_EDIF & "'   ", db, adOpenKeyset, adLockBatchOptimistic
   If rs_aux1.RecordCount > 0 Then
      Set rs_aux3 = New ADODB.Recordset
      If rs_aux3.State = 1 Then rs_aux3.Close
        'Id. detalle "36NO" EXISTENTE
      rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND bien_codigo = '" & CodBien & "' ", db, adOpenStatic
      If rs_aux1.RecordCount > 0 Then
            db.Execute "update ao_ventas_detalle set modelo_codigo = '" & COD_MOD & "' WHERE bien_codigo = '" & CodBien & "' AND venta_codigo = " & rs_aux1!venta_codigo & "   "
      Else
       'VENTAS DETALLE
       Set rs_aux2 = New ADODB.Recordset
       If rs_aux2.State = 1 Then rs_aux2.Close
       rs_aux2.Open "Select * from ao_ventas_detalle where venta_codigo = " & rs_aux1!venta_codigo & "  AND grupo_codigo = '90000' ", db, adOpenStatic
       If rs_aux2.RecordCount > 0 Then
           VAR_CONT = rs_aux2.RecordCount + 1
           If Ado_datos.Recordset!bien_precio_venta_final > 0 Then
                VAR_Dol = Round(Ado_datos.Recordset!bien_precio_venta_final / GlTipoCambioOficial, 2)
           Else
                VAR_Dol = 0
           End If
           'VENTAS
           db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido,   estado_codigo, usr_codigo, fecha_registro) " & _
           " VALUES ('" & glGestion & "', " & rs_aux1!venta_codigo & ", '" & CodBien & "', " & VAR_CONT & ", " & rs_aux2!venta_det_cantidad & ", " & rs_aux2!venta_precio_unitario_bs & ", 0, " & rs_aux2!venta_precio_total_bs & ", " & Round(rs_aux2!venta_precio_unitario_dol, 2) & ", 0, " & Round(rs_aux2!venta_precio_total_dol, 2) & ", '" & Ado_datos.Recordset!bien_descripcion & "', '40000', '43000', '43340', 0, 0, '" & COD_MOD & "', '" & COD_MOD & "', 'S/M', 'S/M', 'S', 'APR', '" & glusuario & "', '" & Date & "') "

        If rs_aux2!bien_codigo = "NA1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          If rs_aux2.RecordCount > 1 Then
              db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
              If rs_aux2.RecordCount > 2 Then
                  db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                  If rs_aux2.RecordCount > 3 Then
                      db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA4' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                      If rs_aux2.RecordCount > 4 Then
                          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA5' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                          If rs_aux2.RecordCount > 5 Then
                              db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA6' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                              If rs_aux2.RecordCount > 6 Then
                                  db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA7' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                              End If
                          End If
                      End If
                  End If
              End If
          End If
        End If
        If rs_aux2!bien_codigo = "NE1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
        End If
        If rs_aux2!bien_codigo = "NP1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
        End If
      End If
     End If
   End If
End Sub

Private Sub BtnAux1_Click()
    Fra_aux1.Visible = True
    FraArticulos.Enabled = False
End Sub

Private Sub BtnBuscar_Click()
'  Set ClBuscaGrid = New ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.QueryUtilizado = GlSqlAux
'  ClBuscaGrid.Título = "Elija un Detalle"
'  ClBuscaGrid.EsTdbGrid = True
'  Set ClBuscaGrid.GridTrabajo = dg_datos
'  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'  ClBuscaGrid.Ejecutar
''  If ClBuscaGrid.ElegidoCol1 <> "" Then
''    Ado_datos.Recordset.Filter = adFilterNone
''    Ado_datos.Recordset.MoveFirst
''    Ado_datos.Recordset.Find "grupo_codigo + '-' + bien_codigo   = " & ClBuscaGrid.ElegidoCol1 & " - " & ClBuscaGrid.ElegidoCol2 & ""
'  End If
    buscados = 1
    OptFilGral2.Visible = False
    OptFilGral1.Visible = False
    Call OptFilGral2_Click
  PosibleApliqueFiltro = False
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = dg_datos
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True

End Sub

Private Sub BtnCancelar_Click()
On Error GoTo Que_Error
    CodBien = Ado_datos.Recordset!bien_codigo
    Screen.MousePointer = vbHourglass
    If Ado_datos.Recordset.EditMode <> adEditNone Then Ado_datos.Recordset.CancelUpdate
'    Ado_datos.Caption = "Registro: " & CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & Ado_datos.Recordset.RecordCount
    'BotonesNavegar Me
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraArticulos.Enabled = False
'    TxtGrupo.Enabled = True
    DtcGrupoDes.Enabled = True
'    TxtSub.Enabled = True
    TDBC_Montador.Enabled = True
'    Set dg_datos.DataSource = Ado_datos
    Screen.MousePointer = vbDefault
    'Refrescar Grid
    Call CARGA
    If OptFilGral1.Value = True Then
        Call OptFilGral1_Click        'Pendientes
     Else
        Call OptFilGral2_Click        'TODOS
     End If
     If (dg_datos.SelBookmarks.Count <> 0) Then
        dg_datos.SelBookmarks.Remove 0
     End If
     If Ado_datos.Recordset.RecordCount > 0 And estado = 2 Then
        RsArt.Find "bien_codigo = " & CodBien & "   ", , , 1
        dg_datos.SelBookmarks.Add (RsArt.Bookmark)
     Else
        RsArt.MoveLast
     End If
    estado = 0
    swnuevo = False
    dg_datos.Enabled = True
    
Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnGrabar2_Click()

    If TxtDetalle.Text = "" Then
        'Call BtnGrabar_Click
      'Repuestos Locales
      If dtc_partida.Text = "39820" Then
          VAR_INICIAL = "REPL"
      Else
        'Repuestos Importados
        If dtc_partida.Text = "39810" Then
          VAR_INICIAL = "REPI"
        Else
          'Servicios
          If Montador.Text = "21000" Then
            VAR_INICIAL = "SERV"
          Else
            'Herramientas
            If dtc_partida.Text = "34800" Then
               VAR_INICIAL = "HERR"
            Else
               'Insumos
               If (DtcGrupoCod.Text = "30000") And (dtc_partida.Text <> "34800" And dtc_partida.Text <> "39820" And dtc_partida.Text <> "39810") Then
                  VAR_INICIAL = "INSU"
               Else
                  VAR_INICIAL = "BIEN"
               End If
            End If
          End If
        End If
      End If
'        If rb_otros2.Value = True Then
'           VAR_INICIAL = "BIEN"
'        End If
'      fragrabar.Visible = True
    Else
        'Fra_Grabar.Enabled = True
        Call BtnGrabar_Click
    End If
End Sub

Private Sub BtnCancelar3_Click()
    Fra_reporte.Visible = False
    fraOpciones.Enabled = True
    fraOpciones.Enabled = True
    FraArticulos.Enabled = True
    dg_datos.Enabled = True
End Sub

Private Sub BtnCancelar4_Click()
    Fra_aux1.Visible = False
    FraArticulos.Enabled = True
End Sub

Private Sub BtnGrabar4_Click()
  'Validacion
    VAR_VAL = "OK"
    If Txt_descripcion4.Text = "" Then
      MsgBox "Debe registrar la " + lbl_descripcion4.Caption, vbCritical + vbExclamation, "Validación de datos"
      VAR_VAL = "ERR"
      Exit Sub
    End If
    
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rsMarcas.Close
    rs_aux4.Open "SELECT * FROM ac_bienes_marcas where marca_codigo = '" & dtc_codigo4.Text & "' ", db, adOpenStatic
    'Set AdoMarca.Recordset = rs_aux4
    If rs_aux4.RecordCount Then
        MsgBox "La Marca ya existe, verifique y Vuelva a Intentar ", vbCritical + vbExclamation, "Validación de datos"
    Else
        db.Execute "INSERT INTO ac_bienes_marcas (marca_codigo, marca_descripcion, pais_codigo, estado_codigo, usr_codigo, fecha_registro)  values ('" & dtc_codigo4.Text & "', '" & Txt_descripcion4.Text & "', 'NN', 'APR', '" & glusuario & "', '" & Date & "')"
        marcas.Text = dtc_codigo4.Text
        TDBC_marcas.Text = Txt_descripcion4.Text
    End If
    
    Fra_aux1.Visible = False
    FraArticulos.Enabled = True
End Sub

Private Sub BtnImprimir1_Click()
    Fra_reporte.Visible = True
End Sub

Private Sub BtnModificar_Click()
On Error GoTo Que_Error
    If (Ado_datos.Recordset!estado_codigo = "REG") Or (glusuario = "CARIZACA") Then
        swnuevo = False
        Screen.MousePointer = vbHourglass
        'BotonesEditar Me
        estado = 2
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FraArticulos.Enabled = True
    '    TxtGrupo.Enabled = False
        TDBC_Montador.Enabled = False
        TxtDetalle.Enabled = False
        
    '    TxtSub.Enabled = False
'        If Ado_datos.Recordset!estado_codigo = "REG" Then
'            DtcGrupoDes.Enabled = True
'            TxtDetalle.Enabled = True
'        Else
'            DtcGrupoDes.Enabled = False
'            TxtDetalle.Enabled = False
'        End If
'        Ado_datos.Caption = "Editando Registro..."
        Screen.MousePointer = vbDefault
        dg_datos.Enabled = False
        If Ado_datos.Recordset!subgrupo_codigo = "43000" Then
            dtc_codigo10.Visible = True
            dtc_desc10.Visible = True
            lbl_edif.Visible = True
    '        dtc_codigo8.Visible = True
            dtc_desc8.Visible = True
            lbl_eqp.Visible = True
        Else
            dtc_codigo10.Visible = False
            dtc_desc10.Visible = False
            lbl_edif.Visible = False
    '        dtc_codigo8.Visible = False
            dtc_desc8.Visible = False
            lbl_eqp.Visible = False
        End If
    Else
        MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
    End If
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub btnEliminar_Click()
On Error GoTo Que_Error
    'ao_adjudica_detalle_D
   If Ado_datos.Recordset.RecordCount > 0 Then
      If ExisteDetalle(Ado_datos.Recordset!bien_codigo) Then MsgBox "No se puede eliminar un BIEN o SERVICIO que ya tiene Registros en COMPRAS o ALMACEN.", vbInformation + vbOKOnly, "Atención": Exit Sub
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         'Ado_datos.Recordset.Delete
         Ado_datos.Recordset!estado_codigo = "ANL"
         Ado_datos.Recordset.Update
         Ado_datos.Recordset.Requery
      End If
   Else
        MsgBox "No existen registros para Anular.", vbExclamation, "Atención"
   End If
   Exit Sub
    
'    If Not GlHayRegs Then
'        MsgBox "No existen registro para Anular", vbExclamation + vbOKOnly, "Atención"
'        Exit Sub
'    End If
'    If ExisteDetalle(Ado_datos.Recordset!grupo_codigo & "-" & Ado_datos.Recordset!bien_codigo) Then MsgBox "No se puede eliminar el Detalle seleccionado ya que se tiene registro de Movimientos en Almacen.", vbInformation + vbOKOnly, "Atención": Exit Sub
'    If MsgBox("¿ Está seguro que se va a Anular el registro visualizado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
'        Screen.MousePointer = vbHourglass
'        'Ado_datos.Recordset.Delete
'        Ado_datos.Recordset!estado = 2
'        Ado_datos.Recordset.MoveNext
'        If Ado_datos.Recordset.EOF Then
'          If Ado_datos.Recordset.RecordCount > 0 Then
'            Ado_datos.Recordset.MoveLast
'          Else
'            GlHayRegs = False
'            Ado_datos.Refresh
'          End If
'        End If
'        Screen.MousePointer = vbDefault
'    End If
'    BotonesNavegar Me
'    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnVer_Click()
  On Error GoTo QError
    If Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\BIENES\" & Trim(Ado_datos.Recordset!grupo_codigo) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOTB"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\BIENES\" & Trim(Ado_datos.Recordset!grupo_codigo) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOTB"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        ARCH_FOTO = App.Path + "\BIENES\" + Trim(Ado_datos.Recordset!grupo_codigo) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    CodBien = Ado_datos.Recordset!bien_codigo
    If Guardar_Imagen(db, "Select Foto From ac_bienes_foto Where bien_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo QError
   If Valida Then
      Screen.MousePointer = vbHourglass
      ' Empezar a grabar
      '*********************************
      db.BeginTrans
        'JQA 04/2008
          If swnuevo = True Then
            Ado_datos.Recordset!grupo_codigo = Trim(TxtGrupo.Caption)
            Ado_datos.Recordset!subgrupo_codigo = Trim(TxtSub.Caption)
            Ado_datos.Recordset!par_codigo = Trim(txt_par.Caption)
            If TxtDetalle.Text = "" Then
                Select Case dtc_codigo5.Text
                    Case "I"  'INSUMOS
                       VAR_INICIAL = "INSU"
                    Case "R"  'REPUESTOS
                       If dtc_partida.Text = "39810" Then
                            VAR_INICIAL = "REPI"
                       Else
                          If dtc_partida.Text = "39820" Then
                            VAR_INICIAL = "REPL"
                          Else
                            VAR_INICIAL = "REPO"
                          End If
                       End If
                    Case "H"    'HERRAMIENTAS
                       VAR_INICIAL = "HERR"
                    Case "S"    'SERVICIOS
                       VAR_INICIAL = "SERV"
                    Case "A"   'GENERAL
                        VAR_INICIAL = "BIEN"
                    Case Else   'OTROS NO IDENTIFICADOS
                        VAR_INICIAL = "BIEN"
                End Select

'                If dtc_partida.Text = "39820" Then
'                    VAR_INICIAL = "REPL"
'                Else
'                  'Repuestos Importados
'                  If dtc_partida.Text = "39810" Then
'                    VAR_INICIAL = "REPI"
'                  Else
'                    'Servicios
'                    If Montador.Text = "21000" Then
'                      VAR_INICIAL = "SERV"
'                    Else
'                      'Herramientas
'                      If dtc_partida.Text = "34800" Then
'                         VAR_INICIAL = "HERR"
'                      Else
'                         'Insumos
'                         If (DtcGrupoCod.Text = "30000") And (dtc_partida.Text <> "34800" And dtc_partida.Text <> "39820" And dtc_partida.Text <> "39810") Then
'                            VAR_INICIAL = "INSU"
'                         Else
'                            VAR_INICIAL = "BIEN"
'                         End If
'                      End If
'                    End If
'                  End If
'                End If
                VAR_INSUMO = "1"
                Set rs_aux6 = New ADODB.Recordset
                SQL_FOR = "select * from fc_correl where tipo_tramite = '" & VAR_INICIAL & "' "
                rs_aux6.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux6.RecordCount > 0 Then
                    rs_aux6!numero_correlativo = rs_aux6!numero_correlativo + 1
                    VAR_INSUMO = rs_aux6!numero_correlativo
                    rs_aux6.Update
                End If
                 'Ado_datos.Recordset!bien_codigo = "INS-" + Str(VAR_INSUMO)
                If VAR_INSUMO < 10 Then
                   CodBien = Trim(VAR_INICIAL + "-000" + Trim(Str(VAR_INSUMO)))
                End If
                If VAR_INSUMO > 9 And VAR_INSUMO < 100 Then
                   CodBien = Trim(VAR_INICIAL + "-00" + Trim(Str(VAR_INSUMO)))
                End If
                If VAR_INSUMO > 99 And VAR_INSUMO < 1000 Then
                   CodBien = Trim(VAR_INICIAL + "-0" + Trim(Str(VAR_INSUMO)))
                End If
                If VAR_INSUMO > 999 Then
                   CodBien = Trim(VAR_INICIAL + "-" + Trim(Str(VAR_INSUMO)))
                End If
            Else
                CodBien = Trim(TxtDetalle.Text)
            End If
            Ado_datos.Recordset!bien_codigo = CodBien
            Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo"
            Ado_datos.Recordset!bien_descripcion = txtDescripcion.Text          '+ " - " + TxtInicial
            
          End If
          If swnuevo = False Then
            Ado_datos.Recordset!bien_descripcion = txtDescripcion.Text
            CodBien = Ado_datos.Recordset!bien_codigo
          End If
            Ado_datos.Recordset!bien_descripcion_anterior = TxtDescripcion2.Text
            Ado_datos.Recordset!unimed_codigo = IIf(Unidad.Text = "", "UNI", Unidad.Text)
            Ado_datos.Recordset!marca_codigo = IIf(marcas.Text = "", "S/M", marcas.Text)
            Ado_datos.Recordset!modelo_codigo = IIf(dtc_codigo6.Text = "", "S/M", dtc_codigo6.Text)
            Ado_datos.Recordset!almacen_tipo = IIf(dtc_codigo5.Text = "", "I", dtc_codigo5.Text)
            ' Campos no ligados
            'Ado_datos.Recordset!estado = IIf(chkEstado.Value = vbChecked, 1, 0)
    '        Ado_datos.Recordset!StockInicial = IIf(TxtInicial.Text = "", 0, Val(TxtInicial.Text))      'Val(TxtInicial.Text)
            Ado_datos.Recordset!bien_codigo_anterior = TxtInicial.Text
            Ado_datos.Recordset!bien_codigo_universal = IIf(dtc_codigo8.Text = "", "X", dtc_codigo8.Text)   'TxtInicial.Text
            Ado_datos.Recordset!bien_precio_compra = IIf(TxtPrecComp.Text = "", 0, CDbl(TxtPrecComp.Text))      'CDbl(TxtPrecComp.Text)
            Ado_datos.Recordset!bien_precio_venta_base = IIf(TxtPrecVenta.Text = "", 0, CDbl(TxtPrecVenta.Text))      'CDbl(txtStockMin)
            Ado_datos.Recordset!bien_precio_venta_final = IIf(TxtPrecEst.Text = "", 0, CDbl(TxtPrecEst.Text))      'CDbl(TxtPrecEst)
            Ado_datos.Recordset!bien_precio_compra_dol = IIf(TxtPrecCompD.Text = "", 0, CDbl(TxtPrecCompD.Text))            'DOLARES
            Ado_datos.Recordset!bien_precio_venta_base_dol = IIf(TxtPrecVentaD.Text = "", 0, CDbl(TxtPrecVentaD.Text))      'DOLARES
            Ado_datos.Recordset!bien_precio_venta_final_dol = IIf(TxtPrecEstD.Text = "", 0, CDbl(TxtPrecEstD.Text))         'DOLARES
            Ado_datos.Recordset!bien_stock_inicial = IIf(txtStockIni.Text = "", 0, CDbl(txtStockIni.Text))      'CDbl(txtStockMin)
            Ado_datos.Recordset!bien_stock_minimo = IIf(txtStockMin.Text = "", 0, CDbl(txtStockMin.Text))      'CDbl(txtStockMin)
            
            Ado_datos.Recordset!bien_rotacion = IIf(cmd_rotacion.Text = "", "PROMEDIO", cmd_rotacion.Text)      'CDbl(txtStockMin)
            Ado_datos.Recordset!edif_codigo = IIf(dtc_codigo10.Text = "", "20101-0", dtc_codigo10.Text)      'CDbl(txtStockMin)
            'Ado_datos.Recordset!tipo_eqp = IIf(dtc_codigo8.Text = "", "X", dtc_codigo8.Text)
            Ado_datos.Recordset!pais_codigo = DtcPais.Text
            'Ado_datos.Recordset!ARCHIVO_F = Trim(Ado_datos.Recordset!subgrupo_codigo) + "-" + Trim(Ado_datos.Recordset!bien_codigo) + ".JPG"
            Ado_datos.Recordset!archivo_foto2 = Trim(Ado_datos.Recordset!bien_codigo) + ".JPG"
            Ado_datos.Recordset!estado_codigo = "REG"  'chkEstado
            Ado_datos.Recordset!usr_codigo = glusuario
            Ado_datos.Recordset!fecha_registro = Date
            Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
            '*********************************
            Ado_datos.Recordset!codigo_pSIN = IIf(IsNull(TxtCodigo_pSIN.Text), "", TxtCodigo_pSIN.Text)
            Ado_datos.Recordset!descripcion_pSIN = IIf(IsNull(TxtDescripcionSIN.Text), "", TxtDescripcionSIN.Text)
            '*********************************
            ' Grabar
            Ado_datos.Recordset.Update
            db.CommitTrans
        '*********************************
    '        Ado_datos.Caption = "Registro: " & CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & Ado_datos.Recordset.RecordCount
            ' Colocar los botones en modo navegar
            GlHayRegs = True
            'BotonesNavegar Me
            fraOpciones.Visible = True
            FraGrabarCancelar.Visible = False
            FraArticulos.Enabled = False
    '        TxtGrupo.Enabled = True
            DtcGrupoDes.Enabled = True
    '        TxtSub.Enabled = True
            TDBC_Montador.Enabled = True
            Screen.MousePointer = vbDefault
            Call CARGA
'            CodBien = Ado_datos.Recordset!bien_codigo
            'Refrescar Grid
            If OptFilGral1.Value = True Then
                Call OptFilGral1_Click        'Pendientes
             Else
                Call OptFilGral2_Click        'TODOS
             End If
             If (dg_datos.SelBookmarks.Count <> 0) Then
                dg_datos.SelBookmarks.Remove 0
             End If
             If Ado_datos.Recordset.RecordCount > 0 And estado = 2 Then
                RsArt.Find "bien_codigo = '" & CodBien & "'   ", , , 1
                dg_datos.SelBookmarks.Add (RsArt.Bookmark)
             Else
                RsArt.MoveLast
             End If
            TxtDetalle.Enabled = True
            estado = 0
            'CARGA
            swnuevo = False
            dg_datos.Enabled = True
    'End If
   Else
        MsgBox "Existe un error en los datos registrados, Verifique y vuelva a intentar...", vbExclamation + vbOKOnly, "Atención"
   End If
   swnuevo = False
   Exit Sub
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRefrescar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    Ado_datos.Recordset.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnImprimirA_Click()
  Dim iResult As Integer
'     LiteralCry = Str(Round(AdoRegularizacion.Recordset!monto_Bolivianos, 2))
'  Literal2 = Literal(LiteralCry) + "  Bolivianos"
'  org2 = AdoRegularizacion.Recordset!org_codigo
'  cocmCod_Comp = AdoRegularizacion.Recordset!codigo_pago
  With CryBBSS
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
'    .StoredProcParam(0) = org2
'    .StoredProcParam(1) = cocmCod_Comp
'    .StoredProcParam(2) = Literal2
        .ReportFileName = App.Path & "\Reportes\Almacen\productos.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub
Private Sub BtnImprimir_Click()
    fra_imprimir.Visible = True
    fraOpciones.Enabled = False
    fraOpciones.Enabled = False
    FraArticulos.Enabled = False
    dg_datos.Enabled = False
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub dg_datos_Click()
    If buscados = 0 Then
        OptFilGral2.Visible = True
        OptFilGral1.Visible = True
    End If
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_partida_Click(Area As Integer)
    dtc_partida_des.BoundText = dtc_partida.BoundText
End Sub

Private Sub dtc_partida_des_Click(Area As Integer)
    dtc_partida.BoundText = dtc_partida_des.BoundText
End Sub

Private Sub dtc_partida_des_LostFocus()
    txt_par.Caption = dtc_partida.Text
End Sub

Private Sub DtcGrupoCod_Click(Area As Integer)
    DtcGrupoDes.BoundText = DtcGrupoCod.BoundText
    DtcGrupoUni.BoundText = DtcGrupoCod.BoundText
End Sub

Private Sub DtcGrupoDes_Click(Area As Integer)
   DtcGrupoCod.BoundText = DtcGrupoDes.BoundText
   DtcGrupoUni.BoundText = DtcGrupoDes.BoundText
   Call pOrganismo(DtcGrupoCod.BoundText)
       TDBC_Montador.Enabled = True
   dtc_partida_des.Enabled = False
    
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from ac_bienes_subgrupo where grupo_codigo='" & CodFuente & "'"
   
   Set Montador.RowSource = Nothing
   Set Montador.RowSource = db.Execute(strConsultaF, , adCmdText)
   Montador.ReFill
   Montador.BoundText = Empty
   
   Set TDBC_Montador.RowSource = Nothing
   Set TDBC_Montador.RowSource = db.Execute(strConsultaF, , adCmdText)
   TDBC_Montador.ReFill
   TDBC_Montador.BoundText = Empty

End Sub

Private Sub DtcGrupoDes_LostFocus()
    TxtGrupo.Caption = DtcGrupoCod.Text
End Sub

Private Sub DtcGrupoUni_Click(Area As Integer)
    DtcGrupoDes.BoundText = DtcGrupoUni.BoundText
    DtcGrupoCod.BoundText = DtcGrupoUni.BoundText
End Sub

Private Sub DtcPais_Click(Area As Integer)
    DtcPaisD.BoundText = DtcPais.BoundText
End Sub

Private Sub DtcPaisD_Click(Area As Integer)
    DtcPais.BoundText = DtcPaisD.BoundText
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys vbTab
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Dim Nodo As Node
    Me.Top = 0
    Me.Left = 0
    Screen.MousePointer = vbHourglass
    estado = 0
    db.Execute "update ac_bienes SET bien_stock_inicial = '0' where  bien_stock_inicial IS NULL"
    db.Execute "update ac_bienes SET bien_stock_ingreso = '0' where  bien_stock_ingreso IS NULL"
    db.Execute "update ac_bienes SET bien_stock_salida_mant = '0' where  bien_stock_salida_mant IS NULL"
    db.Execute "update ac_bienes SET bien_stock_salida = '0' where  bien_stock_salida IS NULL"
    
    ' db.Execute "UPDATE ac_bienes SET bien_stock_actual = bien_stock_inicial + bien_stock_ingreso - bien_stock_salida_mant - bien_stock_salida"
    db.Execute "UPDATE ac_bienes SET bien_stock_actual = bien_stock_ingreso - bien_stock_salida"
    ' Cargamos el Arbol
    ' Nodo Principal
'    Set Nodo = trv.Nodes.Add(, , "Grupo", "Grupos", "Grupos")
'    Nodo.Expanded = True
'    Nodo.Bold = True
    dtc_codigo10.Visible = False
    dtc_desc10.Visible = False
    lbl_edif.Visible = False
'    dtc_codigo8.Visible = False
    dtc_desc8.Visible = False
    lbl_eqp.Visible = False
    If glusuario = "CARIZACA" Or glusuario = "RCUELA" Or glusuario = "AFLORES" Or glusuario = "ADMIN" Or glusuario = "JYMAMANI" Or glusuario = "AFLORES" Or glusuario = "CSALINAS" Then
        BtnAñadir.Visible = True
        BtnModificar.Visible = True
        btnEliminar.Visible = True
        BtnAprobar.Visible = True
    Else
        BtnAñadir.Visible = False
        BtnModificar.Visible = False
        btnEliminar.Visible = False
        BtnAprobar.Visible = False
    End If
            
    OptFilGral1 = True
    Call OptFilGral1_Click
    Call CARGA
'    Set rsgrupo = New ADODB.Recordset
'    rsgrupo.Open "SELECT * FROM ALClGrupo ORDER BY CAST (grupo_codigo AS INT) ", db, adOpenStatic
'    Set AdoGrupo.Recordset = rsgrupo
'    If rsgrupo.RecordCount > 0 Then
'      rsgrupo.MoveFirst
'      While Not rsgrupo.EOF
'        Set Nodo = trv.Nodes.Add("Grupo", tvwChild, "D" & Trim(rsgrupo!grupo_codigo), rsgrupo!descgrupo, "NoElegido", "Elegido")
'        rsgrupo.MoveNext
'      Wend
'    Else
'        trv.Nodes(1).Text = "No Existen Grupos Creados..."
'    End If
    '
    'BotonesNavegar Me
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraArticulos.Enabled = False
    Screen.MousePointer = vbDefault
    C_FIJO = 0      '1.92
    C_MANOBR = 0.01
    C_GTOADM = 0.6094
    C_UTILID = 0.2
    C_ROTALT = 0.01
    C_ROTBAJ = 0.02
    C_FIJO = 0.0636
    C_IMPSTO2 = 0.1494
        Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral2_Click()
    Set RsArt = New ADODB.Recordset
    If RsArt.State = 1 Then RsArt.Close
    'rs_aux2.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Select Case glusuario
        Case "AFLORES"   'INSUMOS
            queryinicial = "select * from ac_bienes where (almacen_tipo = 'I' OR almacen_tipo = 'H') "
            'VAR_DET = "30000"
            VAR_DET = "TODOS"
        Case "CARIZACA", "JMAMANI", "SRAMIREZ" 'REPUESTOS
            queryinicial = "select * from ac_bienes where (almacen_tipo = 'R') "
            VAR_DET = "39800"
'        Case "AFLORES"    'HERRAMIENTAS
'            queryinicial = "select * from ac_bienes where (almacen_tipo = 'H') "
'            VAR_DET = "34800"
'        Case "RTORREZ"   'TODOS
'            queryinicial = "select * from ac_bienes where (almacen_tipo <> 'Q')"            '(almacen_tipo <> 'Q')"
'            VAR_DET = "TODOS"
        Case "TCRUZ", "RRONDAL", "RPRIETO"   'INSUMOS, REPUESTOS Y HERRAMIENTAS
            queryinicial = "select * from ac_bienes where (almacen_tipo = 'I' OR almacen_tipo= 'H' OR almacen_tipo= 'R') "
            VAR_DET = "TODOS"
        Case "ADMIN", "RCUELA", "LVASQUEZ", "VPAREDES", "CSALINAS"     'TODOS
            queryinicial = "select * from ac_bienes where (almacen_tipo <> 'Q')"
            VAR_DET = "TODOS"
        Case Else   'ALGUNOS
            queryinicial = "select * from ac_bienes where (almacen_tipo = 'A') "
            VAR_DET = "NINGUNO"
    End Select
        
    RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    RsArt.Sort = "bien_descripcion"
    'RsArt.Sort = "grupo_codigo, subgrupo_codigo, bien_descripcion"
    If RsArt.RecordCount > 0 Then
       GlHayRegs = True  'Variable global
    Else
       GlHayRegs = False
    End If
    Set Ado_datos.Recordset = RsArt
    'Set dg_datos.DataSource = Ado_datos.Recordset
'    Ado_datos.Recordset.Requery
'    Ado_datos.Refresh
    Set dg_datos.DataSource = Ado_datos
End Sub

Private Sub OptFilGral1_Click()
    Set RsArt = New ADODB.Recordset
    If RsArt.State = 1 Then RsArt.Close
    'rs_aux2.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
     Select Case glusuario
        Case "AFLORES"   'INSUMOS Y REPUESTOS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39800' and par_codigo <> '34800'))  ", db, adOpenKeyset, adLockOptimistic, adCmdText
           queryinicial = "select * from ac_bienes where (estado_codigo = 'REG' and (almacen_tipo = 'I' OR almacen_tipo = 'H')) "
            'VAR_DET = "30000"
            VAR_DET = "TODOS"
        Case "CARIZACA" 'REPUESTOS
            queryinicial = "select * from ac_bienes where (estado_codigo = 'REG' and almacen_tipo = 'R') "
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '39800' )  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            VAR_DET = "39800"
'        Case "AFLORES"    'HERRAMIENTAS
'            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            queryinicial = "select * from ac_bienes where (estado_codigo = 'REG' and almacen_tipo = 'H') "
'            VAR_DET = "34800"
        Case "TCRUZ", "RRONDAL"   'INSUMOS, REPUESTOS Y HERRAMIENTAS
            'queryinicial = "select * from ac_bienes where estado_codigo = 'REG' ORDER BY bien_descripcion "
            queryinicial = "select * from ac_bienes where (estado_codigo = 'REG' and almacen_tipo = 'I' OR almacen_tipo= 'H' OR almacen_tipo= 'R') "
            VAR_DET = "TODOS"
        Case "ADMIN", "RCUELA", "LVASQUEZ", "VPAREDES", "CSALINAS"    'TODOS
            'queryinicial = "select * from ac_bienes where estado_codigo = 'REG' ORDER BY bien_descripcion "
            queryinicial = "select * from ac_bienes where (estado_codigo = 'REG'  and almacen_tipo <> 'Q')"
            VAR_DET = "TODOS"
        Case Else   'ALMACEN ADMINISTRATIVO
            queryinicial = "select * from ac_bienes where (estado_codigo = 'REG' and almacen_tipo = 'A') "
            VAR_DET = "NINGUNO"
     End Select
     RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
     RsArt.Sort = "bien_descripcion"
     'RsArt.Sort = "grupo_codigo, subgrupo_codigo, bien_descripcion"
     If RsArt.RecordCount > 0 Then
        GlHayRegs = True  'Variable global
     Else
        GlHayRegs = False
     End If
     Set Ado_datos.Recordset = RsArt
    'Set dg_datos.DataSource = Ado_datos.Recordset
'    Ado_datos.Recordset.Requery
'    Ado_datos.Refresh
    Set dg_datos.DataSource = Ado_datos
End Sub

Private Function Valida() As Boolean
    Valida = False
    If Trim(TxtGrupo.Caption) = "" Then
        MsgBox "Elija el Grupo al Cual pertenece el Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            DtcGrupoDes.SetFocus
        End If
        Exit Function
    End If
    
     If Trim(txt_par.Caption) = "" Then
        MsgBox "Elija el SubGrupo al Cual pertenece el Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            dtc_partida_des.SetFocus
        End If
        Exit Function
    End If
    
    If Trim(TxtSub.Caption) = "" Then
        MsgBox "Elija el Sub-Grupo al Cual pertenece el Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            DtcGrupoDes.SetFocus
        End If
        Exit Function
    End If
'    If Trim(TxtDetalle.Text) = "" Then
'        sino = MsgBox("Desea generar el codigo automaticamente ? ", vbYesNo + vbQuestion, "Atención")
'          If sino = vbYes Then
'          Else
'            If estado <> 0 Then
'                TxtDetalle.SetFocus
'            End If
'            Exit Function
'          End If
'    End If
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            txtDescripcion.SetFocus
        End If
        Exit Function
    End If
    If Trim(Unidad.Text) = "" Then
        MsgBox "Ingrese la Unidad de Medida.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            Unidad.SetFocus
'        End If
        Exit Function
    End If
    If Trim(TxtPrecComp.Text) = "" Then
        MsgBox "Ingrese EL Precio de Compra del Bien.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            TxtPrecComp.SetFocus
'        End If
        Exit Function
    End If
    If Trim(txtStockMin.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Base del Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            txtStockMin.SetFocus
        End If
        Exit Function
    End If
    If Trim(TxtPrecEst.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Cliente del Bien.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            TxtPrecEst.SetFocus
        End If
        Exit Function
    End If
    If dtc_codigo6.Text = "" Or dtc_desc6.Text = "" Then
        MsgBox "El MODELO Registrado es incorrecto, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            dtc_codigo6.SetFocus
        End If
        Exit Function
    End If
    'If TxtGrupo.Caption = "40000" Then
    If txt_par.Caption = "43340" Then
        If dtc_codigo8.Text = "" Then
            MsgBox "El TIPO de EQUIPO Registrado es incorrecto, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
            If estado <> 0 Then
                dtc_codigo8.SetFocus
            End If
            Exit Function
        End If
    End If
    If txtStockIni.Text = "" Then
            MsgBox "Ingrese el Stock inicial, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
            If estado <> 0 Then
                txtStockIni.SetFocus
            End If
            Exit Function
    End If
    If dtc_codigo5.Text = "" Then
            MsgBox "Ingrese <Almacen de: >, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
            If estado <> 0 Then
                dtc_codigo5.SetFocus
            End If
            Exit Function
    End If
    '
    Valida = True
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub

Private Sub Imprimir_Click()
  Dim iResult As Integer
'     LiteralCry = Str(Round(AdoRegularizacion.Recordset!monto_Bolivianos, 2))
'  Literal2 = Literal(LiteralCry) + "  Bolivianos"
'  org2 = AdoRegularizacion.Recordset!org_codigo
'  cocmCod_Comp = AdoRegularizacion.Recordset!codigo_pago
  With CryFis
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
'    .StoredProcParam(0) = org2
'    .StoredProcParam(1) = cocmCod_Comp
'    .StoredProcParam(2) = Literal2
        .ReportFileName = App.Path & "\Reportes\Almacen\productos_inventario.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub

Private Sub marcas_Click(Area As Integer)
    TDBC_marcas.BoundText = marcas.BoundText
End Sub

Private Sub Picture1_Click()
fra_imprimir.Visible = False
fraOpciones.Enabled = True
fraOpciones.Enabled = True
FraArticulos.Enabled = True
dg_datos.Enabled = True

End Sub


Private Sub rb_prevision_Click()
  Dim iResult As Integer
  With CryLista
  
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
        '.ReportFileName = App.Path & "\Reportes\Almacen\Alm_Listado_Gral_Productos.rpt"
        .ReportFileName = App.Path & "\Reportes\Clasificadores\ar_bienes.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With
End Sub

Private Sub BtnImprimir2_Click()

 If rb_insumos.Value = True Then
    Dim iResult As Integer
     Cr_insumos.WindowShowPrintSetupBtn = True
     Cr_insumos.WindowShowRefreshBtn = True
     Cr_insumos.ReportFileName = App.Path & "\REPORTES\clasificadores\ar_bienes_insumos.rpt"
     iResult = Cr_insumos.PrintReport
     If iResult <> 0 Then
         MsgBox Cr_insumos.LastErrorNumber & " : " & Cr_insumos.LastErrorString, vbExclamation + vbOKOnly, "Error"
     End If
   End If
 If rb_repuestos.Value = True Then
    Cr_repuestos.WindowShowPrintSetupBtn = True
    Cr_repuestos.WindowShowRefreshBtn = True
    Cr_repuestos.ReportFileName = App.Path & "\REPORTES\clasificadores\ar_bienes_repuestos.rpt"
    iResult = Cr_repuestos.PrintReport
    If iResult <> 0 Then
        MsgBox Cr_repuestos.LastErrorNumber & " : " & Cr_repuestos.LastErrorString, vbExclamation + vbOKOnly, "Error"
    End If
  End If
  
  If rb_herramientas.Value = True Then
     Cr_herramientas.WindowShowPrintSetupBtn = True
     Cr_herramientas.WindowShowRefreshBtn = True
     Cr_herramientas.ReportFileName = App.Path & "\REPORTES\clasificadores\ar_bienes_herramientas.rpt"
     iResult = Cr_herramientas.PrintReport
     If iResult <> 0 Then
         MsgBox Cr_herramientas.LastErrorNumber & " : " & Cr_herramientas.LastErrorString, vbExclamation + vbOKOnly, "Error"
     End If
   End If
   
  If rb_servicios.Value = True Then
    Cr_otros.WindowShowPrintSetupBtn = True
    Cr_otros.WindowShowRefreshBtn = True
    Cr_otros.ReportFileName = App.Path & "\REPORTES\clasificadores\ar_bienes_otros.rpt"
    iResult = Cr_otros.PrintReport
    If iResult <> 0 Then
        MsgBox Cr_otros.LastErrorNumber & " : " & Cr_otros.LastErrorString, vbExclamation + vbOKOnly, "Error"
    End If
  End If
  
  If rb_otros.Value = True Then
    Cr_otros.WindowShowPrintSetupBtn = True
    Cr_otros.WindowShowRefreshBtn = True
    Cr_otros.ReportFileName = App.Path & "\REPORTES\clasificadores\ar_bienes_otros.rpt"
    iResult = Cr_otros.PrintReport
    If iResult <> 0 Then
        MsgBox Cr_otros.LastErrorNumber & " : " & Cr_otros.LastErrorString, vbExclamation + vbOKOnly, "Error"
    End If
  End If
End Sub

Private Sub BtnCancelar2_Click()
    fra_imprimir.Visible = False
    fraOpciones.Enabled = True
    fraOpciones.Enabled = True
    FraArticulos.Enabled = True
    dg_datos.Enabled = True

End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Picture3_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
       ' CryV01.StoredProcParam(1) = DTPicker3.Value
        CryV01.StoredProcParam(1) = Format(DTPicker3.Value, "dd/mm/yyyy")
        DTPicker3.Value = Date
'        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
'        VAR_TITULO = "MODULO ALMACENES"
'        CryV01.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
        CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
        CryV01.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "
      
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
    Fra_reporte.Visible = True
End Sub

Private Sub TDBC_marcas_Click(Area As Integer)
    marcas.BoundText = TDBC_marcas.BoundText
End Sub

Private Sub Montador_Click(Area As Integer)
    TDBC_Montador.BoundText = Montador.BoundText
End Sub

Private Sub TDBC_Montador_Click(Area As Integer)
   dtc_partida_des.Enabled = True
   TDBC_Montador.Enabled = True
    Montador.BoundText = TDBC_Montador.BoundText
    Call pPartida(Montador.BoundText)
    dtc_partida_des.Enabled = True
End Sub

Private Sub pPartida(CodPar As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from fc_partida_gasto where subgrupo_codigo='" & CodPar & "' AND estado_codigo = 'APR' "
   
   Set dtc_partida.RowSource = Nothing
   Set dtc_partida.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_partida.ReFill
   dtc_partida.BoundText = Empty
   
   Set dtc_partida_des.RowSource = Nothing
   Set dtc_partida_des.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_partida_des.ReFill
   dtc_partida_des.BoundText = Empty

End Sub

Private Sub TDBC_Montador_LostFocus()
    If TxtSub.Caption = "43000" Then
        dtc_codigo10.Visible = True
        dtc_desc10.Visible = True
        lbl_edif.Visible = True
'        dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        lbl_eqp.Visible = True
    Else
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        lbl_edif.Visible = False
'        dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        lbl_eqp.Visible = False
    End If
    TxtSub.Caption = Montador.Text
End Sub
    
Private Sub TDBC_Unidad_Click(Area As Integer)
    Unidad.BoundText = TDBC_Unidad.BoundText
End Sub

Private Sub Txt_descripcion4_Change()
    dtc_codigo4.Text = Left(Txt_descripcion4, 5)
End Sub

Private Sub TxtPrecComp_LostFocus()
    If TxtPrecComp = "" Then
        TxtPrecCompD = 0
        TxtPrecComp = 0
        TxtPrecVenta = 0
        TxtPrecVentaD = 0
        TxtPrecEst = 0
        TxtPrecEstD = 0
    Else
        TxtPrecCompD = Round(CDbl(TxtPrecComp) / GlTipoCambioOficial, 2)
        'TxtPrecVenta = Round(CDbl(TxtPrecComp) * C_FIJO, 2)
        TxtPrecVenta = Round(CDbl(TxtPrecComp) + (CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTALT) + (CDbl(TxtPrecComp) * C_IMPSTO), 2)
        TxtPrecVentaD = Round(CDbl(TxtPrecVenta) / GlTipoCambioOficial, 2)
        If cmd_rotacion.Text = "ALTA" Then
            C_FIJO = Round((CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTALT), 2)
            TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTALT) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
            'TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(TxtPrecVenta) * C_MANOBR) + (CDbl(TxtPrecVenta) * C_GTOADM) + (CDbl(TxtPrecVenta) * C_UTILID) + (CDbl(TxtPrecVenta) * C_ROTALT) + (CDbl(TxtPrecVenta) * C_IMPSTO), 2)
        Else
            C_FIJO = Round((CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTBAJ), 2)
            TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTBAJ) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
            'TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(TxtPrecVenta) * C_MANOBR) + (CDbl(TxtPrecVenta) * C_GTOADM) + (CDbl(TxtPrecVenta) * C_UTILID) + (CDbl(TxtPrecVenta) * C_ROTBAJ) + (CDbl(TxtPrecVenta) * C_IMPSTO), 2)
        End If
        TxtPrecEstD = Round(CDbl(TxtPrecEst) / GlTipoCambioOficial, 2)
    End If
End Sub

Private Sub TxtPrecCompD_LostFocus()
    If TxtPrecCompD = "" Then
        TxtPrecCompD = 0
        TxtPrecComp = 0
        TxtPrecVenta = 0
        TxtPrecVentaD = 0
        TxtPrecEst = 0
        TxtPrecEstD = 0
    Else
        TxtPrecComp = Round(CDbl(TxtPrecCompD) * GlTipoCambioOficial, 2)
        'TxtPrecVentaD = Round(CDbl(TxtPrecCompD) * C_FIJO, 2)
        TxtPrecVentaD = Round(CDbl(TxtPrecCompD) + (CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_IMPSTO), 2)
        TxtPrecVenta = Round(CDbl(TxtPrecVentaD) * GlTipoCambioOficial, 2)
        If cmd_rotacion.Text = "ALTA" Then
            C_FIJO = Round((CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_UTILID) + (CDbl(TxtPrecCompD) * C_ROTALT), 2)
            'TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(TxtPrecVentaD) * C_MANOBR) + (CDbl(TxtPrecVentaD) * C_GTOADM) + (CDbl(TxtPrecVentaD) * C_UTILID) + (CDbl(TxtPrecVentaD) * C_ROTALT) + (CDbl(TxtPrecVentaD) * C_IMPSTO), 2)
            TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTALT) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
        Else
            C_FIJO = Round((CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_UTILID) + (CDbl(TxtPrecCompD) * C_ROTBAJ), 2)
            'TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(TxtPrecVentaD) * C_MANOBR) + (CDbl(TxtPrecVentaD) * C_GTOADM) + (CDbl(TxtPrecVentaD) * C_UTILID) + (CDbl(TxtPrecVentaD) * C_ROTBAJ) + (CDbl(TxtPrecVentaD) * C_IMPSTO), 2)
            TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTBAJ) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
        End If
        TxtPrecEst = Round(CDbl(TxtPrecEstD) * GlTipoCambioOficial, 2)
    End If
End Sub

Private Sub Unidad_Click(Area As Integer)
    TDBC_Unidad.BoundText = Unidad.BoundText
End Sub

'Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
'    If InStr(Node.Key, "G") = 0 Then
'        TxtGrupo.caption = Mid(Node.Key, 2)
'    Else
'        TxtGrupo.caption = ""
'    End If
'End Sub

'Private Sub BuscaNodo(QNodo As String)
'Dim Nodo As Node
'Dim Indice As Integer
'    For Indice = 1 To trv.Nodes.Count
'        If Mid(trv.Nodes(Indice).Key, 2) = QNodo Then
'            trv.Nodes(Indice).Selected = True
'            Exit For
'        End If
'    Next
'End Sub

'Private Sub txtStockMin_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
'End Sub
'Private Sub txtUnidadCaja_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
'End Sub

Private Function ExisteDetalle(bien_codigo As String) As Boolean
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_no_objecion_detalle_D WHERE bien_codigo = '" & bien_codigo & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM Ao_solicitud_bienes WHERE bien_codigo = '" & bien_codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteDetalle = rs!Cuantos > 0
End Function

Private Sub CARGA()
    Set rsMarcas = New ADODB.Recordset
    If rsMarcas.State = 1 Then rsMarcas.Close
    rsMarcas.Open "SELECT * FROM ac_bienes_marcas ORDER BY marca_descripcion", db, adOpenStatic
    Set AdoMarca.Recordset = rsMarcas
    
    Set rsUnidad = New ADODB.Recordset
    If rsUnidad.State = 1 Then rsUnidad.Close
    rsUnidad.Open "Select * from ac_bienes_unidad_medida order by unimed_descripcion", db, adOpenStatic
    Set AdoMedida.Recordset = rsUnidad
    
    Set rsMontador = New ADODB.Recordset
    If rsMontador.State = 1 Then rsMontador.Close
    rsMontador.Open "select * from ac_bienes_subgrupo order by subgrupo_descripcion", db, adOpenStatic
    Set AdoMontador.Recordset = rsMontador
    
    Set rsgrupo = New ADODB.Recordset
    If rsgrupo.State = 1 Then rsgrupo.Close
    rsgrupo.Open "SELECT * FROM ac_bienes_grupo WHERE estado_codigo='APR' ", db, adOpenStatic
    Set AdoGrupo.Recordset = rsgrupo
    
    Set rsPais = New ADODB.Recordset
    If rsPais.State = 1 Then rsPais.Close
    rsPais.Open "SELECT * FROM gc_pais WHERE estado_codigo='APR' order by pais_descripcion", db, adOpenStatic
    Set AdoPais.Recordset = rsPais
    'gc_edificaciones
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    'gc_edificaciones
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from ac_bienes_modelos ", db, adOpenStatic     'order by modelo_descripcion
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    'fc_partidas_gasto
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_partida_gasto order by par_descripcion", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_partida_des.BoundText = dtc_partida.BoundText
    'ac_bienes_equipo_tipos
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from ac_bienes_equipo_tipos order by tipo_eqp_descripcion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    'ac_almacen_tipo
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "Select * from ac_almacen_tipo where estado_codigo = 'APR' order by almacen_tipo_descripcion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_aux5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub


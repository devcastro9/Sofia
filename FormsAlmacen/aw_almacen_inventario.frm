VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form aw_almacen_inventario 
   Caption         =   "Inventario de Almacenes"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "aw_almacen_inventario.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame fra_reportes 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Elija una de las opciones ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   1920
      TabIndex        =   32
      Top             =   2640
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton btnSalirPanel 
         Caption         =   "Salir"
         Height          =   495
         Left            =   7320
         TabIndex        =   43
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton btnPrintOption 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   5760
         TabIndex        =   42
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de BIENES"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   39
         Top             =   3120
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Certificado de Cumplimiento de Contrato por Mantenimiento Integral"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   38
         Top             =   3480
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contratos VIGENTES con detalle de EQUIPOS (Zona Piloto) MIGRAR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Salidas de Almacenes con PPP, para Exportar a Excel"
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
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   2400
         Width           =   6735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inventario DETALLADO del Almacen - Ingresos y Salidas (VALORADO)"
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
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   1440
         Width           =   6855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inventario GENERAL del Almacen - Totalizado por Bien (Cantidades)"
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
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   720
         Width           =   6495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inventario GENERAL del Almacen - Totalizado por Bien (VALORADO)"
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
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   1080
         Width           =   6735
      End
      Begin Crystal.CrystalReport Cr_otros 
         Left            =   120
         Top             =   600
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIO DE TODOS LOS ALMACENES"
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
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   2040
         Width           =   5415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTARIO POR ALMACEN ELEGIDO"
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
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Fra_reporte 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFF00&
      Height          =   1935
      Left            =   2280
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   8055
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         FillColor       =   &H00404040&
         FillStyle       =   2  'Horizontal Line
         ForeColor       =   &H80000008&
         Height          =   676
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   7800
         TabIndex        =   14
         Top             =   240
         Width           =   7800
         Begin VB.PictureBox CmdFiltrar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3240
            Picture         =   "aw_almacen_inventario.frx":6852
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   29
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6240
            Picture         =   "aw_almacen_inventario.frx":7040
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   27
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1800
            Picture         =   "aw_almacen_inventario.frx":7802
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   26
            ToolTipText     =   "Kardex Valorado"
            Top             =   0
            Width           =   1455
         End
         Begin VB.PictureBox BtnImprimir1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "aw_almacen_inventario.frx":80CF
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   15
            ToolTipText     =   "Kardex Físico"
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
            TabIndex        =   16
            Top             =   195
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker DTP_Finicio 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44197
      End
      Begin MSComCtl2.DTPicker DTP_Ffin 
         DataField       =   "Fecha_Alerta"
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   110624769
         CurrentDate     =   44561
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE INICIO"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1680
         TabIndex        =   19
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE FIN"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4440
         TabIndex        =   17
         Top             =   1080
         Width           =   1485
      End
   End
   Begin VB.PictureBox fraOpciones 
      Align           =   1  'Align Top
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11400
      TabIndex        =   5
      Top             =   0
      Width           =   11400
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8760
         Picture         =   "aw_almacen_inventario.frx":899C
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   44
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6480
         Picture         =   "aw_almacen_inventario.frx":9151
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   30
         ToolTipText     =   "Salidas de Almacenes p/Exportar a Excel..."
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnImprimir3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   4920
         Picture         =   "aw_almacen_inventario.frx":9A1E
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   28
         ToolTipText     =   "Inventario Detallado de un Almacen Elegido..."
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17400
         Picture         =   "aw_almacen_inventario.frx":A2EB
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   8
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
         Left            =   240
         Picture         =   "aw_almacen_inventario.frx":AAAD
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   7
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3360
         Picture         =   "aw_almacen_inventario.frx":B262
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   6
         ToolTipText     =   "Inventario General de un Almacen Elegido..."
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
         Left            =   13305
         TabIndex        =   9
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.PictureBox Fra_Elegir 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   11400
      TabIndex        =   4
      Top             =   660
      Width           =   11400
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "aw_almacen_inventario.frx":BB2F
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   120
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "almacen_descripcion"
         BoundColumn     =   "almacen_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "aw_almacen_inventario.frx":BB48
         DataField       =   "almacen_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   -120
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "almacen_codigo"
         BoundColumn     =   "almacen_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "aw_almacen_inventario.frx":BB61
         DataField       =   "bien_codigo"
         Height          =   315
         Left            =   9000
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_cod2 
         Bindings        =   "aw_almacen_inventario.frx":BB7E
         DataField       =   "bien_codigo"
         Height          =   315
         Left            =   16080
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   ""
      End
      Begin VB.Label lbl_bien2 
         Caption         =   "Buscar por Código -->"
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
         Height          =   255
         Left            =   14040
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl_bien 
         Caption         =   "Buscar por Nombre -->"
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
         Height          =   255
         Left            =   6960
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label cmdItem 
         Caption         =   "Elija Almacen -->"
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   7920
      Width           =   11400
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   1215
         TabIndex        =   1
         Top             =   255
         Width           =   6945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Inventario"
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
         Left            =   8340
         TabIndex        =   2
         Top             =   90
         Width           =   3360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Inventario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   8355
         TabIndex        =   3
         Top             =   105
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos 
      Height          =   330
      Left            =   14040
      Top             =   6480
      Visible         =   0   'False
      Width           =   2145
      _ExtentX        =   3784
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
      Caption         =   "Ado_datos"
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
      Left            =   14040
      Top             =   6960
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
   Begin Crystal.CrystalReport Cry 
      Left            =   14040
      Top             =   7800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   14640
      Top             =   7800
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
   Begin MSAdodcLib.Adodc ado_datos_busq 
      Height          =   330
      Left            =   14040
      Top             =   7320
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
      Caption         =   "ado_datos_busq"
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
   Begin MSDataGridLib.DataGrid tdbgInventario 
      Align           =   3  'Align Left
      Bindings        =   "aw_almacen_inventario.frx":BB9B
      Height          =   6315
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   25
      Top             =   1215
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11139
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12572159
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
      Caption         =   $"aw_almacen_inventario.frx":BBB3
      ColumnCount     =   9
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "bien_codigo"
         Caption         =   "Código"
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
         DataField       =   "bien_descripcion"
         Caption         =   "Descripcion"
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
         DataField       =   "total_compra_bs"
         Caption         =   "Valor en Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "stock_ingreso"
         Caption         =   "Cantidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "total_venta_bs"
         Caption         =   "Valor en Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "stock_salida"
         Caption         =   "Cantidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "utilidad_Bs"
         Caption         =   "Valor en Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "stock_actual"
         Caption         =   "Cantidad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4529.764
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CryV02 
      Left            =   15240
      Top             =   7800
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
   Begin Crystal.CrystalReport Cry01 
      Left            =   14040
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport Cry02 
      Left            =   14640
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   45
      Top             =   7530
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   688
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Digite ""DOBLE CLICK"", para ver KARDEX de cada Item"
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
      Height          =   1095
      Left            =   14160
      TabIndex        =   18
      Top             =   5160
      Width           =   1935
   End
End
Attribute VB_Name = "aw_almacen_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cnn As ADODB.Connection
Dim RsInventario As ADODB.Recordset
Dim rs_datos1 As ADODB.Recordset
Dim rs_datos2 As ADODB.Recordset
Dim rs_datos3 As ADODB.Recordset
Dim RsGrupos As ADODB.Recordset
Dim rs_aux1 As ADODB.Recordset
Dim rs_aux2 As ADODB.Recordset

'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
'Dim queryinicial As String
Dim iResult As Integer

Dim VAR_SW As String
Dim VAR_BIEN, VAR_BIEN2 As String
Dim CodGrupo As String

Dim CANT_ING, COMPRA_UNIT_BS, COMPRA_TOT_BS As Double
Dim CANT_SAL, VENTA_UNIT_BS, VENTA_TOT_BS As Double
Dim CANT_SALDO, SALDO_UNIT_BS, SALDO_TOT_BS As Double
Dim UNIT87_BS, TOT87_BS As Double

'Dim cmm As ADODB.Command
Dim VAR_ALM, VAR_CONTAR As Integer

Dim FInicio, FFin As Date

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
        'MsgBox "OK ..."
    Else
        'MsgBox "ERROR ..."
    End If
End Sub

Private Sub Buscar()
    If (tdbgInventario.SelBookmarks.Count <> 0) Then
       tdbgInventario.SelBookmarks.Remove 0
    End If
   
    If RsInventario.RecordCount > 0 Then
        RsInventario.Find "bien_codigo = '" & dtc_cod2.Text & "'", , , 1
        tdbgInventario.SelBookmarks.Add (RsInventario.Bookmark)
    Else
        'sino = MsgBox("No se encontro ningun bien con ese nombre", vbInformation, "Aviso")
        'Call Carga_Beneficiario(1)
        'dtc_buscar_desc.Text = ""
        MsgBox "No Existe el bien ..."
    End If
End Sub

Private Sub BtnCancelar3_Click()
    Fra_reporte.Visible = False
    tdbgInventario.Enabled = True
    Fra_Elegir.Enabled = True
End Sub

Private Sub BtnImprimir_Click()
  If dtc_codigo1.Text <> "" Then
    'opt_salir.Value = True
    fra_reportes.Visible = True
'    'If Ado_datos.Recordset.RecordCount > 0 Then
'      Dim iResult As Integer
'      Screen.MousePointer = vbHourglass
'      Cry.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex_tot_alm.rpt"
'      Cry.StoredProcParam(0) = dtc_codigo1.Text         'Ado_datos.Recordset!almacen_codigo
'
'      iResult = Cry.PrintReport
'      Screen.MousePointer = vbDefault
'      If iResult <> 0 Then
'          MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Atención"
'      End If
'
''      Dim IResult As Integer
''        'Dim co As New ADODB.Command
''        CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
''        CryV01.WindowShowPrintSetupBtn = True
''        CryV01.WindowShowRefreshBtn = True
''        'CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
''        CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
''        CryV01.StoredProcParam(1) = Format(DTPicker3.Value, "dd/mm/yyyy")
''        CryV01.StoredProcParam(2) = Ado_datos.Recordset!almacen_codigo            'dtc_codigo1.Text
''        DTPicker3.Value = Date
'''        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
'''        VAR_TITULO = "MODULO ALMACENES"
'''        CryV01.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
''        CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
''        CryV01.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "
''
''        IResult = CryV01.PrintReport
''        If IResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
''        CryV01.WindowState = crptMaximized
'    'Else
'    '      MsgBox "No se puede Imprimir. Debe elegir el Almacen y vuelva a intentar ...", , "Atención"
'    'End If
  Else
        MsgBox "No se puede Imprimir. Debe elegir el Almacen y vuelva a intentar ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
'        Dim iResult As Integer
'        'Dim co As New ADODB.Command
'        'CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex.rpt"
        CryV01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado.rpt" '
        CryV01.WindowShowPrintSetupBtn = True
        CryV01.WindowShowRefreshBtn = True
        'CryV01.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryV01.StoredProcParam(0) = VAR_BIEN
        CryV01.StoredProcParam(1) = Trim(Str(VAR_ALM))            'dtc_codigo1.Text
        CryV01.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryV01.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        
'        DTPicker3.Value = Date
''        CryV01.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
''        VAR_TITULO = "MODULO ALMACENES"
''        CryV01.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
'        CryV01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'        CryV01.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "

        CryV01.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
        'CryV01.Formulas(2) = "DEL_AL = '' "
        'CryV01.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
        
        iResult = CryV01.PrintReport
        If iResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"
        CryV01.WindowState = crptMaximized
        Fra_reporte.Visible = False
        tdbgInventario.Enabled = True
        Fra_Elegir.Enabled = True
    Else
        MsgBox "No se puede Imprimir. Verifique si existen datos y vuelva a intentar ...", , "Atención"
    End If
    'Fra_reporte.Visible = True
End Sub

Private Sub BtnImprimir2_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        VAR_ALM = dtc_codigo1.Text
        Call ACTUALIZA_PPP_BIEN
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        FInicio = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        FFin = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        'CryV02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado_valorado.rpt" '
        CryV02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado_valorado_full.rpt" '
        CryV02.WindowShowPrintSetupBtn = True
        CryV02.WindowShowRefreshBtn = True
        'CryV02.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryV02.StoredProcParam(0) = VAR_BIEN
        CryV02.StoredProcParam(1) = Trim(Str(VAR_ALM))            'dtc_codigo1.Text
        CryV02.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryV02.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        
'        DTPicker3.Value = Date
''        CryV02.StoredProcParam(1) = Ado_datos.Recordset!ges_gestion
''        VAR_TITULO = "MODULO ALMACENES"
''        CryV02.Formulas(0) = "titulo = '" & VAR_TITULO & "' "
'        CryV02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'        CryV02.Formulas(2) = "FechaAl = '" & DTPicker3.Value & "' "
'
        CryV02.Formulas(0) = "almace = '" & dtc_desc1.Text & "' "
        CryV02.Formulas(1) = "FInicio = '" & DTP_Finicio.Value & "' "
        CryV02.Formulas(2) = "FFin = '" & DTP_Ffin.Value & "' "
        CryV02.Formulas(3) = "CodAlm = '" & Trim(Str(VAR_ALM)) & "' "
        'CryV02.Formulas(2) = "DEL_AL = '' "
        'CryV02.Formulas(3) = "fechafin = '" & DTP_Ffin.Value & "' "
        
        iResult = CryV02.PrintReport
        If iResult <> 0 Then MsgBox CryV02.LastErrorNumber & " : " & CryV02.LastErrorString, vbCritical, "Error de impresión"
        CryV02.WindowState = crptMaximized
        Fra_reporte.Visible = False
        tdbgInventario.Enabled = True
        Fra_Elegir.Enabled = True
    Else
        MsgBox "No se puede Imprimir. Verifique si existen datos y vuelva a intentar ...", , "Atención"
    End If

End Sub

Private Sub BtnImprimir3_Click()
'    Fra_reporte.Visible = True
'    tdbgInventario.Enabled = False
'    Fra_Elegir.Enabled = False
'    CmdFiltrar.Visible = True
'    BtnImprimir1.Visible = False
'    BtnImprimir2.Visible = False
End Sub

Private Sub BtnImprimir4_Click()
''Cry01
'    Dim iResult As Integer
'      Screen.MousePointer = vbHourglass
'      Cry01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_todos_TXT.rpt"
'      'Cry01.StoredProcParam(0) = dtc_codigo1.Text         'Ado_datos.Recordset!almacen_codigo
'
'      iResult = Cry01.PrintReport
'      Screen.MousePointer = vbDefault
'      If iResult <> 0 Then
'          MsgBox Cry01.LastErrorNumber & " : " & Cry01.LastErrorString, vbExclamation + vbOKOnly, "Atención"
'      End If
      
End Sub

Private Sub btnPrintOption_Click()
    Dim iResult As Integer
    If Option1.Value = True Then
        Screen.MousePointer = vbHourglass
        Cry.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex_tot_alm.rpt"
        Cry.StoredProcParam(0) = dtc_codigo1.Text         'Ado_datos.Recordset!almacen_codigo
      
        iResult = Cry.PrintReport
        Screen.MousePointer = vbDefault
        If iResult <> 0 Then
            MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Atención"
        End If
    ElseIf Option2.Value = True Then
        DTP_Finicio.Value = "01/01/" & glGestion
        DTP_Ffin.Value = ObtenerFechaServidor()
        Fra_reporte.Visible = True
        tdbgInventario.Enabled = False
        Fra_Elegir.Enabled = False
        CmdFiltrar.Visible = True
        BtnImprimir1.Visible = False
        BtnImprimir2.Visible = False
        fra_reportes.Visible = False
    ElseIf Option3.Value = True Then
        DTP_Finicio.Value = "01/01/" & glGestion
        DTP_Ffin.Value = ObtenerFechaServidor()
        Fra_reporte.Visible = True
        tdbgInventario.Enabled = False
        Fra_Elegir.Enabled = False
        CmdFiltrar.Visible = True
        BtnImprimir1.Visible = False
        BtnImprimir2.Visible = False
        fra_reportes.Visible = False
    ElseIf Option4.Value = True Then
        Screen.MousePointer = vbHourglass
        Cry01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacenes_todos_TXT.rpt"

        iResult = Cry01.PrintReport
        Screen.MousePointer = vbDefault
        If iResult <> 0 Then
            MsgBox Cry01.LastErrorNumber & " : " & Cry01.LastErrorString, vbExclamation + vbOKOnly, "Atención"
        End If
    ElseIf Option5.Value = True Then
        MsgBox "Reporte en desarrollo...", vbExclamation + vbOKOnly, "Atención"
    ElseIf Option6.Value = True Then
        MsgBox "Reporte en desarrollo...", vbExclamation + vbOKOnly, "Atención"
    ElseIf Option7.Value = True Then
        MsgBox "Reporte en desarrollo...", vbExclamation + vbOKOnly, "Atención"
    End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub btnSalirPanel_Click()
    fra_reportes.Visible = False
End Sub

Private Sub cmdFiltrar_Click()
   If Ado_datos.Recordset.RecordCount > 0 Then
        'Dim iResult As Integer
        'Dim co As New ADODB.Command
        If Option2.Value = True Then
            CryV02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_almacen_kardex_tot_alm_valorado.rpt"
        Else
            If Option3.Value = True Then
                CryV02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado_valorado_full.rpt" '
            Else
                CryV02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_kardex_almacen_acumulado_valorado_full.rpt" '
            End If
        End If
        CryV02.WindowShowPrintSetupBtn = True
        CryV02.WindowShowRefreshBtn = True
        'CryV02.StoredProcParam(0) = Ado_datos.Recordset!bien_codigo
        CryV02.StoredProcParam(0) = "%"                                    'VAR_BIEN
        CryV02.StoredProcParam(1) = Trim(Str(dtc_codigo1.Text))            'Trim(Str(VAR_ALM))
        CryV02.StoredProcParam(2) = Format(DTP_Finicio.Value, "dd/mm/yyyy")
        CryV02.StoredProcParam(3) = Format(DTP_Ffin.Value, "dd/mm/yyyy")
        
        CryV02.Formulas(1) = "almace = '" & dtc_desc1.Text & "' "
        
        iResult = CryV02.PrintReport
        If iResult <> 0 Then MsgBox CryV02.LastErrorNumber & " : " & CryV02.LastErrorString, vbCritical, "Error de impresión"
        CryV02.WindowState = crptMaximized
        Fra_reporte.Visible = False
        tdbgInventario.Enabled = True
        Fra_Elegir.Enabled = True
        
        CmdFiltrar.Visible = False
        BtnImprimir1.Visible = True
        BtnImprimir2.Visible = True
    Else
        MsgBox "No se puede Imprimir. Verifique si existen datos y vuelva a intentar ...", , "Atención"
    End If
End Sub

Private Sub cmdItem_Click()
'JQA
'  Set ClBuscaGrid = New ClBuscaEnGridExterno    ' ClBuscaEnGridPropio
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.FiltrosMultiples = True
'  'ClBuscaGrid.QueryUtilizado = "SELECT CodGrupo +'-'+ CodDetalle As CodGrupo, DescDetalle FROM ALCLdetalle"
'  ClBuscaGrid.QueryUtilizado = "SELECT almacen_codigo +'-'+ almacen_descripcion As CodGrupo, almacen_descripcion as DescDetalle FROM ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1'  "
'  ClBuscaGrid.Título = "Elija un Almacen"
'  ClBuscaGrid.OcultarPrimero = True
'  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    CodGrupo = ClBuscaGrid.ElegidoCol1
'    tdbcGrupos.Text = ClBuscaGrid.ElegidoCol2
'  End If
'  Set ClBuscaGrid = Nothing
'JQA
End Sub

Private Sub dtc_cod2_Change()
'    dtc_desc2.BoundText = dtc_cod2.BoundText
'    If dtc_cod2.SelectedItem <> "" Then
'         Call Buscar
'     End If
End Sub

Private Sub dtc_cod2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_cod2.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_desc2_Change()
'    RsInventario.Sort = "bien_descripcion"
'    Set ado_datos_busq.Recordset = RsInventario.DataSource
    dtc_cod2.BoundText = dtc_desc2.BoundText
    If dtc_cod2.SelectedItem <> "" Then
         Call Buscar
     End If
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_cod2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc1_Change()
    dtc_desc2.Visible = False
    lbl_bien.Visible = False
    'lbl_bien2.Visible = True
    'dtc_cod2.Visible = True
    If VAR_SW = "NO" Then
        dtc_codigo1.BoundText = dtc_desc1.BoundText
        dtc_cod2.BoundText = dtc_desc2.BoundText
    End If
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select count(almacen_codigo) as cont1 from av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & "   ", db, adOpenStatic
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "select count(almacen_codigo) as cont2 from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & "   ", db, adOpenStatic
    
    If rs_aux1!CONT1 <> rs_aux2!CONT2 Then
        Call ACTUALIZA_PPP
    End If
    'f dtc_codigo1.Text = "" Then
    '    MsgBox "El Almacen No existe o no tiene Movimiento... , vuelva a intentar ...", vbInformation + vbOKOnly, "Atención"
    '    VAR_SW = "SI"
    'Else
        'dtc_codigo1.BoundText = dtc_desc1.BoundText
        'dtc_cod2.BoundText = dtc_desc2.BoundText
        Set RsInventario = New ADODB.Recordset
        If RsInventario.State = 1 Then RsInventario.Close
        'queryinicial = "select * from av_almacen_inventario where almacen_codigo = " & dtc_codigo1.Text & "  "
        'RsInventario.Open queryinicial, db, adOpenKeyset, adLockReadOnly            'adLockOptimistic
        RsInventario.Open "select * from av_almacen_inventario_total where almacen_codigo = " & dtc_codigo1.Text & " order by bien_descripcion ", db, adOpenKeyset, adLockReadOnly
        'RsInventario.Sort = "bien_descripcion"
        VAR_SW = "SI"
        Set Ado_datos.Recordset = RsInventario.DataSource
        Set tdbgInventario.DataSource = RsInventario.DataSource
        If RsInventario.RecordCount > 0 Then
            'RsInventario.Sort = "bien_descripcion"
            Set ado_datos_busq.Recordset = RsInventario.DataSource
            dtc_cod2.BoundText = dtc_desc2.BoundText
            dtc_desc2.Visible = True
            lbl_bien.Visible = True
            'lbl_bien2.Visible = True
            'dtc_cod2.Visible = True
            VAR_SW = "NO"
        Else
            MsgBox "No existe Movimiento en el Almacen, vuelva a intentar ...", vbInformation + vbOKOnly, "Atención"
            VAR_SW = "NO"
        End If
'    Totales
    'End If
    'VAR_SW = "NO"
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub ACTUALIZA_PPP()
'Dim UNIT87_BS, TOT87_BS As Double
    ' CLONA INGRESOS Y SALIDAS DE ALMACEN
    'db.Execute "DROP TABLE ao_saldos3 "
    'db.Execute "SELECT * INTO ao_saldos3 FROM av_almacenes_saldos "
    
'    db.Execute "DELETE ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " "
'    db.Execute "INSERT INTO ao_saldos3 SELECT * FROM av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & ""
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "select bien_codigo from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " GROUP BY bien_codigo  ", db, adOpenStatic
    rs_datos2.Open "select bien_codigo from av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & " GROUP BY bien_codigo  ", db, adOpenStatic
    'Set Ado_datos2.Recordset = rs_datos2
    If rs_datos2.RecordCount > 0 Then
        ProgressBar1.Visible = True
        With ProgressBar1
            .Max = rs_datos2.RecordCount
            .Min = 0
            .Value = 0
        End With
        rs_datos2.MoveFirst
        While Not rs_datos2.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            VAR_BIEN2 = rs_datos2!bien_codigo
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux1 = New ADODB.Recordset
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "select count(almacen_codigo) as cont1 from av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "'  ", db, adOpenStatic
            
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "select count(almacen_codigo) as cont2 from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "'  ", db, adOpenStatic
            
            If rs_aux1!CONT1 <> rs_aux2!CONT2 Then
                'ACTUALIZA PPP DEL ALMACEN Y BIEN
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
                db.Execute "DELETE ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "'  "
                db.Execute "INSERT INTO ao_saldos3 SELECT * FROM av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & "  AND bien_codigo = '" & VAR_BIEN2 & "' "
    
                Set rs_datos3 = New ADODB.Recordset
                If rs_datos3.State = 1 Then rs_datos3.Close
                rs_datos3.Open "select * from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' ORDER BY fecha_ingreso, doc_codigo ", db, adOpenStatic
                If rs_datos3.RecordCount > 0 Then
                    rs_datos3.MoveFirst
                    CANT_ING = rs_datos3!cantidad_ingreso
                    COMPRA_UNIT_BS = rs_datos3!compra_bs_unit
                    COMPRA_TOT_BS = rs_datos3!importe_compra_bs
                    CANT_SAL = rs_datos3!cantidad_salida
                    'VENTA_UNIT_BS = rs_datos3!CostoUnitario
                    '=SI(cantidad_salida=0;CostoUnitarioCPP;0)
                    If CANT_SAL = 0 Then
                        VENTA_UNIT_BS = 0
                    Else
                        VENTA_UNIT_BS = COMPRA_UNIT_BS                  'SALDO_UNIT_BS
                    End If
                    'VENTA_TOT_BS = rs_datos3!importe_venta_bs
                    VENTA_TOT_BS = Round(CANT_SAL * VENTA_UNIT_BS, 2)
                    'CANT_SALDO = rs_datos3!cantidad_saldo
                    CANT_SALDO = CANT_ING - CANT_SAL
                    'SALDO_UNIT_BS = rs_datos3!CostoUnitarioCPP
                    ''SI(CostoUnitario>0;CostoUnitario;(importe_compra_bs+importe_venta_bs)/)cantidad_saldo
                    If VENTA_UNIT_BS > 0 Then
                        SALDO_UNIT_BS = VENTA_UNIT_BS
                    Else
                        SALDO_UNIT_BS = Round((COMPRA_TOT_BS + VENTA_TOT_BS) / CANT_SALDO, 2)
                    End If
                    'SALDO_TOT_BS = rs_datos3!CostoTotalCPP
                    SALDO_TOT_BS = Round(CANT_SALDO * SALDO_UNIT_BS, 2)
                    'UNIT87_BS = rs_datos3!CostoUnitario87
                    UNIT87_BS = Round(SALDO_UNIT_BS * 0.87, 2)
                    'TOT87_BS = rs_datos3!valortotal87
                    TOT87_BS = Round(SALDO_TOT_BS * 0.87, 2)
                    While Not rs_datos3.EOF
                        db.Execute "UPDATE ao_saldos3 SET CostoUnitario = " & VENTA_UNIT_BS & ", importe_venta_bs = " & VENTA_TOT_BS & ", cantidad_saldo = " & CANT_SALDO & ", CostoUnitarioCPP = " & SALDO_UNIT_BS & ", CostoTotalCPP = " & SALDO_TOT_BS & " where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' AND doc_codigo = '" & rs_datos3!doc_codigo & "' AND doc_numero = " & rs_datos3!doc_numero & "  "            'AND fecha_ingreso = '" & Format(rs_datos3!fecha_ingreso, "dd/mm/yyyy") & "'
                        db.Execute "UPDATE ao_saldos3 SET correlativo = " & rs_datos3.RecordCount & " where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' AND doc_codigo = '" & rs_datos3!doc_codigo & "' AND doc_numero = " & rs_datos3!doc_numero & "  "
                        rs_datos3.MoveNext
                        If Not rs_datos3.EOF Then
                            'VARIABLES
                            If rs_datos3!compra_bs_unit = 0 Then
                                VENTA_UNIT_BS = SALDO_UNIT_BS
                            Else
                                VENTA_UNIT_BS = 0
                            End If
                            VENTA_TOT_BS = Round(rs_datos3!cantidad_salida * VENTA_UNIT_BS, 2)
                            CANT_SALDO = (rs_datos3!cantidad_ingreso - rs_datos3!cantidad_salida) + CANT_SALDO
                            If rs_datos3!compra_bs_unit = 0 Then
                                SALDO_UNIT_BS = SALDO_UNIT_BS
                            Else
                                If CANT_SALDO = 0 Then
                                    SALDO_UNIT_BS = Round((rs_datos3!importe_compra_bs + SALDO_TOT_BS), 2)
                                Else
                                    SALDO_UNIT_BS = Round((rs_datos3!importe_compra_bs + SALDO_TOT_BS) / CANT_SALDO, 2)
                                End If
                            End If
                            SALDO_TOT_BS = Round(CANT_SALDO * SALDO_UNIT_BS, 2)
                        End If
                    Wend
                End If
            End If
            rs_datos2.MoveNext
        Wend
        
        db.Execute "UPDATE ao_saldos3 SET CostoUnitario87 = CostoUnitarioCPP * 0.87 , valortotal87 = CostoTotalCPP * 0.87  where almacen_codigo = " & dtc_codigo1.Text & "  "
        'ACTUALIZA PRECIO SALIDA ALMACEN
        db.Execute "update ao_almacen_salidas set precio_unitario_bs = ao_saldos3.CostoUnitario, importe_venta_bs = ao_saldos3.importe_venta_bs FROM ao_almacen_salidas INNER JOIN ao_saldos3 ON ao_almacen_salidas.almacen_codigo = ao_saldos3.almacen_codigo AND ao_almacen_salidas.doc_codigo  = ao_saldos3.doc_codigo AND ao_almacen_salidas.doc_numero = ao_saldos3.doc_numero AND ao_almacen_salidas.bien_codigo = ao_saldos3.bien_codigo WHERE ao_almacen_salidas.almacen_codigo = " & dtc_codigo1.Text & "  "
        'ACTUALIZA ao_almacen_totales
        ProgressBar1.Visible = False
    Else
        rs_datos2.Close
    End If
End Sub

Private Sub ACTUALIZA_PPP_BIEN()
'Dim UNIT87_BS, TOT87_BS As Double
    ' CLONA INGRESOS Y SALIDAS DE ALMACEN
    'db.Execute "DROP TABLE ao_saldos3 "
    'db.Execute "SELECT * INTO ao_saldos3 FROM av_almacenes_saldos "
    VAR_CONTAR = 0
    VAR_BIEN2 = Ado_datos.Recordset!bien_codigo
    'VAR_BIEN2 = Ado_datos.Recordset!bien_codigo
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select count(almacen_codigo) as cont1 from av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "'  ", db, adOpenStatic
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "select count(almacen_codigo) as cont2 from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "'  ", db, adOpenStatic
    
    If rs_aux1!CONT1 <> rs_aux2!CONT2 Then
        db.Execute "DELETE ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & Trim(VAR_BIEN2) & "' "
        db.Execute "INSERT INTO ao_saldos3 SELECT * FROM av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' "
    End If
    'db.Execute "UPDATE ao_saldos3  SET bien_codigo = LTRIM(RTRIM(bien_codigo)) WHERE bien_codigo LIKE '% %'"
'    Set rs_datos2 = New ADODB.Recordset
'    If rs_datos2.State = 1 Then rs_datos2.Close
'    rs_datos2.Open "select bien_codigo from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " GROUP BY bien_codigo  ", db, adOpenStatic
'    'Set Ado_datos2.Recordset = rs_datos2
'    If rs_datos2.RecordCount > 0 Then
'        ProgressBar1.Visible = True
'        With ProgressBar1
'            .Max = rs_datos2.RecordCount
'            .Min = 0
'            .Value = 0
'        End With
'      'ProgressBar1.Max =
'        rs_datos2.MoveFirst
'        While Not rs_datos2.EOF
'            ProgressBar1.Value = ProgressBar1.Value + 1
'            VAR_BIEN2 = rs_datos2!bien_codigo
            Set rs_datos3 = New ADODB.Recordset
            If rs_datos3.State = 1 Then rs_datos3.Close
            rs_datos3.Open "select * from ao_saldos3 where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' ORDER BY fecha_ingreso, doc_codigo ", db, adOpenStatic
            'rs_datos3.Open "select * from av_almacenes_saldos where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' ORDER BY fecha_ingreso, doc_codigo ", db, adOpenStatic
            If rs_datos3.RecordCount > 0 Then
                ProgressBar1.Visible = True
                With ProgressBar1
                    .Max = rs_datos3.RecordCount
                    .Min = 0
                    .Value = 0
                End With
                rs_datos3.MoveFirst
                VAR_CONTAR = 1
                CANT_ING = rs_datos3!cantidad_ingreso
                COMPRA_UNIT_BS = rs_datos3!compra_bs_unit
                COMPRA_TOT_BS = rs_datos3!importe_compra_bs
                CANT_SAL = rs_datos3!cantidad_salida
                'VENTA_UNIT_BS = rs_datos3!CostoUnitario
                '=SI(cantidad_salida=0;CostoUnitarioCPP;0)
                If CANT_SAL = 0 Then
                    VENTA_UNIT_BS = 0
                Else
                    VENTA_UNIT_BS = COMPRA_UNIT_BS                  'SALDO_UNIT_BS
                End If
                'VENTA_TOT_BS = rs_datos3!importe_venta_bs
                VENTA_TOT_BS = Round(CANT_SAL * VENTA_UNIT_BS, 2)
                'CANT_SALDO = rs_datos3!cantidad_saldo
                CANT_SALDO = CANT_ING - CANT_SAL
                'SALDO_UNIT_BS = rs_datos3!CostoUnitarioCPP
                ''SI(CostoUnitario>0;CostoUnitario;(importe_compra_bs+importe_venta_bs)/)cantidad_saldo
                If VENTA_UNIT_BS > 0 Then
                    SALDO_UNIT_BS = VENTA_UNIT_BS
                Else
                    SALDO_UNIT_BS = Round((COMPRA_TOT_BS + VENTA_TOT_BS) / CANT_SALDO, 2)
                End If
                'SALDO_TOT_BS = rs_datos3!CostoTotalCPP
                SALDO_TOT_BS = Round(CANT_SALDO * SALDO_UNIT_BS, 2)
                'UNIT87_BS = rs_datos3!CostoUnitario87
                UNIT87_BS = Round(SALDO_UNIT_BS * 0.87, 2)
                'TOT87_BS = rs_datos3!valortotal87
                TOT87_BS = Round(SALDO_TOT_BS * 0.87, 2)
                While Not rs_datos3.EOF
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    db.Execute "UPDATE ao_saldos3 SET CostoUnitario = " & VENTA_UNIT_BS & ", importe_venta_bs = " & VENTA_TOT_BS & ", cantidad_saldo = " & CANT_SALDO & ", CostoUnitarioCPP = " & SALDO_UNIT_BS & ", CostoTotalCPP = " & SALDO_TOT_BS & " where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' AND doc_codigo = '" & rs_datos3!doc_codigo & "' AND doc_numero = " & rs_datos3!doc_numero & "  "            'AND fecha_ingreso = '" & Format(rs_datos3!fecha_ingreso, "dd/mm/yyyy") & "'
                    db.Execute "UPDATE ao_saldos3 SET correlativo = " & VAR_CONTAR & " where almacen_codigo = " & dtc_codigo1.Text & " AND bien_codigo = '" & VAR_BIEN2 & "' AND doc_codigo = '" & rs_datos3!doc_codigo & "' AND doc_numero = " & rs_datos3!doc_numero & "  "
                    rs_datos3.MoveNext
                    VAR_CONTAR = VAR_CONTAR + 1
                    If Not rs_datos3.EOF Then
                        'VARIABLES
                            If rs_datos3!compra_bs_unit = 0 Then
                                VENTA_UNIT_BS = SALDO_UNIT_BS
                            Else
                                VENTA_UNIT_BS = 0
                            End If
                            VENTA_TOT_BS = Round(rs_datos3!cantidad_salida * VENTA_UNIT_BS, 2)
                            CANT_SALDO = (rs_datos3!cantidad_ingreso - rs_datos3!cantidad_salida) + CANT_SALDO
                            If rs_datos3!compra_bs_unit = 0 Then
                                SALDO_UNIT_BS = SALDO_UNIT_BS
                            Else
                                If CANT_SALDO = 0 Then
                                    SALDO_UNIT_BS = Round((rs_datos3!importe_compra_bs + SALDO_TOT_BS), 2)
                                    'SALDO_UNIT_BS = Round((rs_datos3!compra_bs_unit + SALDO_TOT_BS), 2)
                                Else
                                    SALDO_UNIT_BS = Round((rs_datos3!importe_compra_bs + SALDO_TOT_BS) / CANT_SALDO, 2)
                                End If
                            End If
                            SALDO_TOT_BS = Round(CANT_SALDO * SALDO_UNIT_BS, 2)
                    End If
                Wend
            End If
'            rs_datos2.MoveNext
'        Wend
        db.Execute "UPDATE ao_saldos3 SET CostoUnitarioCPP =0  where almacen_codigo = " & dtc_codigo1.Text & " and CostoUnitarioCPP is null  "
        db.Execute "UPDATE ao_saldos3 SET CostoTotalCPP =0  where almacen_codigo = " & dtc_codigo1.Text & " and CostoTotalCPP is null  "
        db.Execute "UPDATE ao_saldos3 SET CostoUnitario87 = CostoUnitarioCPP * 0.87 , valortotal87 = CostoTotalCPP * 0.87  where almacen_codigo = " & dtc_codigo1.Text & "  "
        'ACTUALIZA PRECIO SALIDA ALMACEN
        db.Execute "update ao_almacen_salidas set precio_unitario_bs = ao_saldos3.CostoUnitario, importe_venta_bs = ao_saldos3.importe_venta_bs FROM ao_almacen_salidas INNER JOIN ao_saldos3 ON ao_almacen_salidas.almacen_codigo = ao_saldos3.almacen_codigo AND ao_almacen_salidas.doc_codigo  = ao_saldos3.doc_codigo AND ao_almacen_salidas.doc_numero = ao_saldos3.doc_numero AND ao_almacen_salidas.bien_codigo = ao_saldos3.bien_codigo WHERE ao_almacen_salidas.almacen_codigo = " & dtc_codigo1.Text & "  "
        'ACTUALIZA ao_almacen_totales
        ProgressBar1.Visible = False
'    Else
'        rs_datos2.Close
'    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Top = 0
    Me.Left = 0
    VAR_SW = "NO"

    'ac_almacenes ' Origen
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select * from ac_almacenes where almacen_codigo <> '0' AND almacen_codigo <> '1' AND almacen_tipo = '" & Aux & "' ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
   
    db.Execute "UPDATE ao_almacen_totales SET stock_ingreso ='0' WHERE (stock_ingreso IS NULL) "
    db.Execute "UPDATE ao_almacen_totales SET stock_salida = '0' where (stock_salida IS NULL)"
    db.Execute "UPDATE ao_almacen_totales SET stock_actual = stock_ingreso - stock_salida"
'    Totales

    Screen.MousePointer = vbDefault
        Call SeguridadSet(Me)
End Sub

'Private Sub Form_Resize()
'On Error Resume Next
''    tdbgInventario.Width = Me.ScaleWidth - picBoton.Width
'End Sub

Public Sub Totales()
'Dim rs As ADODB.Recordset
'Dim ValorSus As Currency
'Dim PrecIng As Currency
'Dim EjmIng As Long
'Dim PrecSal As Long
'Dim EjmEnt As Long
'Dim valor As Long
'Dim Ejmtot As Long
'    Set rs = New ADODB.Recordset
'    Set rs = RsInventario
'    PrecIng = 0
'    EjmIng = 0
'    PrecSal = 0
'    EjmEnt = 0
'    valor = 0
'    Ejmtot = 0
'    'ValorSus = 0
'    While Not rs.EOF
'        'JQA 04/2008
'        PrecIng = PrecIng + IIf(IsNull(rs!PrecIng), 0, rs!PrecIng)
'        'CajaIng = CajaIng + 1
'        EjmIng = EjmIng + IIf(IsNull(rs!EjmIng), 0, rs!EjmIng)
'        PrecSal = PrecSal + IIf(IsNull(rs!PrecSal), 0, rs!PrecSal)
'        'CajaEnt = CajaEnt + 1
'        EjmEnt = EjmEnt + IIf(IsNull(rs!EjmEnt), 0, rs!EjmEnt)
'        valor = valor + IIf(IsNull(rs!valor), 0, rs!valor)
'        'CajaSal = CajaSal + 1
'        'EjmSal = EjmSal + IIf(IsNull(rs!EjmSal), 0, rs!EjmSal)
'        Ejmtot = EjmIng - EjmEnt
'        valor = PrecIng - PrecSal
'        'ValorSus = ValorSus + IIf(IsNull(rs!valor), 0, rs!valor)
'        rs.MoveNext
'    Wend
''    tdbgInventario.Columns("bien_descripcion").FooterText = "TOTALES"
''    tdbgInventario.Columns("PrecIng").FooterText = Format(PrecIng, "###,###,##0") & ""
''    tdbgInventario.Columns("EjmIng").FooterText = Format(EjmIng, "###,###,##0") & ""
''    tdbgInventario.Columns("PrecSal").FooterText = Format(PrecSal, "###,###,##0") & ""
''    tdbgInventario.Columns("EjmEnt").FooterText = Format(EjmEnt, "###,###,##0") & ""
''    tdbgInventario.Columns("valor").FooterText = Format(valor, "###,###,##0") & ""
''    tdbgInventario.Columns("Ejmtot").FooterText = Format(Ejmtot, "###,###,##0") & ""
''    'tdbgInventario.Columns("Valor").FooterText = Format(ValorSus, "###,###,##0.00") & " $us"
End Sub

Private Sub tdbgInventario_DblClick()
    VAR_BIEN = Ado_datos.Recordset!bien_codigo
    VAR_ALM = Ado_datos.Recordset!almacen_codigo
    'db.Execute "UPDATE ao_almacen_totales SET stock_salida = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & VAR_BIEN & "' AND almacen_codigo = " & VAR_ALM & ") where almacen_codigo = " & VAR_ALM & " and bien_codigo = '" & VAR_BIEN & "' "
    'db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida"
    DTP_Finicio.Value = "01/01/" & glGestion
    DTP_Ffin.Value = ObtenerFechaServidor()
    Fra_reporte.Visible = True
    tdbgInventario.Enabled = False
    Fra_Elegir.Enabled = False
    CmdFiltrar.Visible = False
    BtnImprimir1.Visible = True
    BtnImprimir2.Visible = True
End Sub

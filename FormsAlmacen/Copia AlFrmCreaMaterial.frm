VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlFrmCreaMaterial 
   Caption         =   "Clasificadores - Almacenes -  Materiales(Productos)"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   120
   ClientWidth     =   14415
   Icon            =   "AlFrmCreaMaterial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   14415
   WindowState     =   2  'Maximized
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00C0C0C0&
      Height          =   950
      Left            =   4080
      TabIndex        =   26
      Top             =   960
      Width           =   8580
      Begin VB.CommandButton Imprimir 
         Caption         =   "Inventario Fisico"
         Height          =   735
         Left            =   6000
         Picture         =   "AlFrmCreaMaterial.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdImpCabeza 
         Caption         =   "Inventario Valorado"
         Height          =   735
         Left            =   5160
         Picture         =   "AlFrmCreaMaterial.frx":7FD4
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdAnadir 
         Caption         =   "Adicionar"
         Height          =   735
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "AlFrmCreaMaterial.frx":9756
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Modificar"
         Height          =   735
         Left            =   960
         Picture         =   "AlFrmCreaMaterial.frx":9A60
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   735
         Left            =   2640
         Picture         =   "AlFrmCreaMaterial.frx":9C6A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   4320
         Picture         =   "AlFrmCreaMaterial.frx":9E74
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Anular"
         Height          =   735
         Left            =   1800
         Picture         =   "AlFrmCreaMaterial.frx":A07E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3480
         Picture         =   "AlFrmCreaMaterial.frx":A768
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   735
         Left            =   7560
         Picture         =   "AlFrmCreaMaterial.frx":A972
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc AdoMontador 
      Height          =   375
      Left            =   6360
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "montador"
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
      Left            =   5160
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   3960
      Top             =   7680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   375
      Left            =   0
      Top             =   7320
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14415
      TabIndex        =   20
      Top             =   7515
      Width           =   14415
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   15
         TabIndex        =   22
         Top             =   255
         Width           =   8370
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   10200
         TabIndex        =   23
         Top             =   75
         Width           =   1845
      End
   End
   Begin TrueOleDBGrid60.TDBGrid tdbgArt 
      Bindings        =   "AlFrmCreaMaterial.frx":AB7C
      Height          =   6255
      Left            =   0
      OleObjectBlob   =   "AlFrmCreaMaterial.frx":AB91
      TabIndex        =   7
      Top             =   1035
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   990
      Left            =   0
      Picture         =   "AlFrmCreaMaterial.frx":EBD5
      ScaleHeight     =   930
      ScaleWidth      =   14355
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   14415
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1035
         TabIndex        =   18
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   60
         TabIndex        =   17
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Administrativa"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1035
         TabIndex        =   16
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   60
         TabIndex        =   15
         Top             =   420
         Width           =   750
      End
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE MATERIALES (Productos)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   5490
         TabIndex        =   24
         Top             =   240
         Width           =   6555
      End
      Begin VB.Image Image1 
         Height          =   1440
         Left            =   0
         Picture         =   "AlFrmCreaMaterial.frx":11E6F
         Top             =   0
         Width           =   15360
      End
   End
   Begin MSComctlLib.ImageList imlMaterial 
      Left            =   4200
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrmCreaMaterial.frx":1695F
            Key             =   "Grupos"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrmCreaMaterial.frx":17239
            Key             =   "NoElegido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrmCreaMaterial.frx":17B13
            Key             =   "Elegido"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraArticulos 
      Height          =   5850
      Left            =   4050
      TabIndex        =   19
      Top             =   1840
      Width           =   8565
      Begin VB.TextBox TxtPrecEst 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "Precio_estimado"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   53
         Text            =   "0"
         Top             =   5400
         Width           =   1365
      End
      Begin VB.TextBox TxtInicial 
         Alignment       =   2  'Center
         DataField       =   "StockInicial"
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   46
         Text            =   "0"
         Top             =   4680
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         DataField       =   "Cod_Montador"
         DataSource      =   "AdoArt"
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "11"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtStockMin 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "Precio_salon"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "0"
         Top             =   5400
         Width           =   1365
      End
      Begin VB.TextBox txtUnidadCaja 
         Alignment       =   2  'Center
         DataField       =   "Precio_compra"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "0"
         Top             =   5400
         Width           =   1365
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         DataField       =   "CodDetalle"
         DataSource      =   "AdoArt"
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
         Height          =   315
         Left            =   120
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "111"
         Top             =   3975
         Width           =   1335
      End
      Begin VB.TextBox TxtGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         DataField       =   "CodGrupo"
         DataSource      =   "AdoArt"
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "11"
         Top             =   660
         Width           =   615
      End
      Begin VB.CheckBox chkEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activo"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7200
         TabIndex        =   14
         Top             =   5280
         Width           =   1185
      End
      Begin MSComctlLib.TreeView trv 
         Height          =   1665
         Left            =   810
         TabIndex        =   8
         Top             =   690
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   2937
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlMaterial"
         Appearance      =   1
         Enabled         =   0   'False
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H80000018&
         DataField       =   "DescDetalle"
         DataSource      =   "AdoArt"
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3975
         Width           =   7095
      End
      Begin VB.TextBox TxtActual 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "StockActual"
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "0"
         Top             =   4680
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo TDBC_marcas 
         Bindings        =   "AlFrmCreaMaterial.frx":183ED
         DataField       =   "COD_MARCA"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   4695
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion"
         BoundColumn     =   "COD_MARCA"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo marcas 
         Bindings        =   "AlFrmCreaMaterial.frx":18404
         DataField       =   "COD_MARCA"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ListField       =   "COD_MARCA"
         BoundColumn     =   "COD_MARCA"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo TDBC_Montador 
         Bindings        =   "AlFrmCreaMaterial.frx":1841B
         DataField       =   "COD_MONTADOR"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   1080
         TabIndex        =   36
         Top             =   3030
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777152
         ListField       =   "descripcion"
         BoundColumn     =   "COD_MONTADOR"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo Montador 
         Bindings        =   "AlFrmCreaMaterial.frx":18435
         DataField       =   "COD_MONTADOR"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   1080
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777152
         ListField       =   "COD_MONTADOR"
         BoundColumn     =   "COD_MONTADOR"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo TDBC_Unidad 
         Bindings        =   "AlFrmCreaMaterial.frx":1844F
         DataField       =   "Unidad"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   2880
         TabIndex        =   38
         Top             =   4695
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "descripcion"
         BoundColumn     =   "Unidad"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo Unidad 
         Bindings        =   "AlFrmCreaMaterial.frx":18467
         DataField       =   "Unidad"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   4440
         TabIndex        =   39
         Top             =   4320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Unidad"
         BoundColumn     =   "Unidad"
         Text            =   "Elige Marca..."
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Cliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   52
         Top             =   5160
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción del Grupo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   390
         Width           =   7785
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUB-GRUPO"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2460
         Width           =   8415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Inicial"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   45
         Top             =   4425
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Sub"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label TDBFrame3D6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Actual"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7080
         TabIndex        =   42
         Top             =   4425
         Width           =   1365
      End
      Begin VB.Label TDBFrame3D7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Compra"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5145
         Width           =   1365
      End
      Begin VB.Label TDBFrame3D8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio Salon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   5145
         Width           =   1365
      End
      Begin VB.Label TDBFrame3D5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad de Medida"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   4350
         Width           =   1455
      End
      Begin VB.Label TDBFrame3D10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción del Sub-Grupo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   32
         Top             =   2760
         Width           =   7425
      End
      Begin VB.Label TDBFrame3D9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marca"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4350
         Width           =   1095
      End
      Begin VB.Label TDBFrame3D3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código Producto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label TDBFrame3D4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción del Producto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   3720
         Width           =   7080
      End
      Begin VB.Label TDBFrame3D1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3420
         Width           =   8415
      End
      Begin VB.Label TDBFrame3D2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GRUPO"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   8415
      End
   End
   Begin Crystal.CrystalReport CryBBSS 
      Left            =   12120
      Top             =   1200
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
      Left            =   12120
      Top             =   1680
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
End
Attribute VB_Name = "AlFrmCreaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim rsMarcas As ADODB.Recordset
Dim rsunidad As ADODB.Recordset
Dim rsMontador As ADODB.Recordset

Dim rsgrupo As ADODB.Recordset
Dim RsArt As ADODB.Recordset
Dim rsNada As ADODB.Recordset
'--------
Dim estado As Integer ' 0 navegar, 1 Agregar, 2 Editar
'--
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim swnuevo As Boolean

Public Sub ALPrincipal(QEstado As Integer)
    '
    Screen.MousePointer = vbHourglass
    estado = QEstado
    '
    Select Case estado
        Case 0
            Set RsArt = New ADODB.Recordset
            'JQA 04/2008
            'GlSqlAux = "SELECT * FROM ALCLDetalle WHERE CAST(CODGRUPO AS INT)< 50  AND coddetalle = ISNULL(coddetalle, NULL) ORDER BY CAST (CODGRUPO AS INT)"
            GlSqlAux = "SELECT * FROM ALCLDetalle WHERE coddetalle = ISNULL(coddetalle, NULL) ORDER BY CODGRUPO, cod_montador, codDetalle "
            RsArt.Open GlSqlAux, db, adOpenKeyset, adLockOptimistic
            If RsArt.RecordCount > 0 Then
               GlHayRegs = True  'Variable global
            Else
               GlHayRegs = False
            End If
            BotonesNavegar Me
            FraArticulos.Enabled = False
            Set AdoArt.Recordset = RsArt
        Case 1
                    
        Case 2
        
    End Select
    '
    Screen.MousePointer = vbDefault
    Me.Show
End Sub

Private Sub AdoArt_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim Marca As String
Dim a As Integer
Dim COD_MARCAx, cod_MONTADOR, cod_UMedida, cod_grupo As String
If AdoArt.Recordset.BOF Or AdoArt.Recordset.EOF Then
        If AdoArt.Recordset.BOF And AdoArt.Recordset.EOF Then
            TxtGrupo.Text = ""
            TxtDetalle.Text = ""
            txtDescripcion.Text = ""
            TxtActual.Text = ""
            chkEstado.Value = vbUnchecked
            AdoArt.Caption = "Registro: 0 de 0"
            BuscaNodo "rupo"
        Else
            Exit Sub
        End If
Else
    If swnuevo = False Then
            If Not (IsNull(AdoMarca.Recordset("cod_marca"))) Then
                If Not (AdoMarca.Recordset.BOF) Then AdoMarca.Recordset.MoveFirst
                AdoMarca.Recordset.Find "cod_marca ='" & AdoArt.Recordset!COD_MARCA & "'", , adSearchForward
                If Not AdoMarca.Recordset.EOF Then
                    'TDBC_marcas.Item(1) = AdoMarca.Recordset!descripcion
                    TDBC_marcas.Text = AdoMarca.Recordset!descripcion
                End If
            End If
            If Not (IsNull(AdoMontador.Recordset("cod_montador"))) Then
                If Not (AdoMontador.Recordset.BOF) Then AdoMontador.Recordset.MoveFirst
                AdoMontador.Recordset.Find "Cod_montador ='" & AdoArt.Recordset!cod_MONTADOR & "'", , adSearchForward
                If Not AdoMontador.Recordset.EOF Then
                    '
                End If
            End If
            If Not (IsNull(AdoMedida.Recordset("Unidad"))) Then
                If Not (AdoMedida.Recordset.BOF) Then AdoMedida.Recordset.MoveFirst
                    AdoMedida.Recordset.Find "Unidad ='" & AdoArt.Recordset!Unidad & "'", , adSearchForward
                If Not AdoMedida.Recordset.EOF Then
                    '
                End If
            End If
            
    End If
    
    
    If AdoArt.Recordset!estado = 1 Then
        chkEstado.Value = vbChecked
    Else
        chkEstado.Value = vbUnchecked
        'chkEstado.Value =IIf(CBool(AdoArt.Recordset!estado), vbChecked, vbUnchecked)
    End If
        'TDBC_Montador
        
        
        BuscaNodo AdoArt.Recordset!CodGrupo
End If
End Sub




Private Sub CmdAnadir_Click()
    swnuevo = True
    Set tdbgArt.DataSource = rsNada
    AdoArt.Recordset.AddNew
    estado = 1
    BotonesEditar Me
    FraArticulos.Enabled = True
    trv.SetFocus
    BuscaNodo "grupo"
    txtStockMin.Text = 0
  '  txtUnidadCaja.Text = 0
End Sub

Private Sub CmdBuscar_Click()
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.QueryUtilizado = GlSqlAux
  ClBuscaGrid.Título = "Elija un Detalle"
  ClBuscaGrid.EsTdbGrid = True
  Set ClBuscaGrid.GridTrabajo = tdbgArt
  Set ClBuscaGrid.RecordsetTrabajo = AdoArt.Recordset
  ClBuscaGrid.Ejecutar
'  If ClBuscaGrid.ElegidoCol1 <> "" Then
'    AdoArt.Recordset.Filter = adFilterNone
'    AdoArt.Recordset.MoveFirst
'    AdoArt.Recordset.Find "CodGrupo + '-' + CodDetalle   = " & ClBuscaGrid.ElegidoCol1 & " - " & ClBuscaGrid.ElegidoCol2 & ""
'  End If

End Sub
Private Sub CmdCancelar_Click()

On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If AdoArt.Recordset.EditMode <> adEditNone Then AdoArt.Recordset.CancelUpdate
    AdoArt.Recordset.Requery
    AdoArt.Caption = "Registro: " & CStr(AdoArt.Recordset.AbsolutePosition) & " de " & AdoArt.Recordset.RecordCount
    BotonesNavegar Me
    FraArticulos.Enabled = False
    Set tdbgArt.DataSource = AdoArt
    Screen.MousePointer = vbDefault
    estado = 0
    CARGA
    swnuevo = False
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub Cmdeditar_Click()
On Error GoTo Que_Error
    '
    Screen.MousePointer = vbHourglass
    BotonesEditar Me
    estado = 2
    FraArticulos.Enabled = True
    AdoArt.Caption = "Editando Registro..."
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo Que_Error
    If Not GlHayRegs Then
        MsgBox "No existen registro para Anular", vbExclamation + vbOKOnly, "Atención"
        Exit Sub
    End If
    If ExisteDetalle(AdoArt.Recordset!CodGrupo & "-" & AdoArt.Recordset!codDetalle) Then MsgBox "No se puede eliminar el Detalle seleccionado ya que se tiene registro de Movimientos en Almacen.", vbInformation + vbOKOnly, "Atención": Exit Sub
    If MsgBox("¿ Está seguro que se va a Anular el registro visualizado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
        Screen.MousePointer = vbHourglass
        'AdoArt.Recordset.Delete
        AdoArt.Recordset!estado = 2
        AdoArt.Recordset.MoveNext
        If AdoArt.Recordset.EOF Then
          If AdoArt.Recordset.RecordCount > 0 Then
            AdoArt.Recordset.MoveLast
          Else
            GlHayRegs = False
            AdoArt.Refresh
          End If
        End If
        Screen.MousePointer = vbDefault
    End If
    BotonesNavegar Me
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub
Private Sub CmdGrabar_Click()
On Error GoTo QError
    If valida Then
        Screen.MousePointer = vbHourglass
        ' Empezar a grabar
        '*********************************
        db.BeginTrans
        'JQA 04/2008
        'AdoArt.Recordset!COD_MARCA = Trim(TDBC_marcas.Columns(1))
        'AdoArt.Recordset!CD_REFMATE = TxtDetalle
        'AdoArt.Recordset!cod_MONTADOR = Trim(TDBC_Montador.Columns(1))
        'AdoArt.Recordset!unidad = Trim(TDBC_Unidad.Columns(1))
        AdoArt.Recordset!CodGrupo = Trim(TxtGrupo.Text)
        AdoArt.Recordset!cod_MONTADOR = Trim(Montador.Text)
        AdoArt.Recordset!codDetalle = Trim(TxtDetalle.Text)
        AdoArt.Recordset!descdetalle = txtDescripcion.Text
        AdoArt.Recordset!Unidad = Unidad.Text
        AdoArt.Recordset!COD_MARCA = marcas.Text
        ' Campos no ligados
        'AdoArt.Recordset!estado = IIf(chkEstado.Value = vbChecked, 1, 0)
        AdoArt.Recordset!StockInicial = Val(TxtInicial.Text)
        'AdoArt.Recordset!StockActual = Val(TxtActual.Text)
        AdoArt.Recordset!Precio_compra = CDbl(txtUnidadCaja)
        AdoArt.Recordset!Precio_salon = CDbl(txtStockMin)
        AdoArt.Recordset!Precio_estimado = CDbl(TxtPrecEst)
        AdoArt.Recordset!estado = chkEstado
        AdoArt.Recordset!usr_usuario = GlUsuario
        AdoArt.Recordset!fecha_registro = Date
        AdoArt.Recordset!hora_registro = Format(Time, "hh:mm:ss")
        '*********************************
        ' Grabar
        AdoArt.Recordset.Update
        db.CommitTrans
    '*********************************
        AdoArt.Caption = "Registro: " & CStr(AdoArt.Recordset.AbsolutePosition) & " de " & AdoArt.Recordset.RecordCount
        ' Colocar los botones en modo navegar
        GlHayRegs = True
        BotonesNavegar Me
        FraArticulos.Enabled = False
        Screen.MousePointer = vbDefault
        CARGA
        AdoArt.Refresh
        Set tdbgArt.DataSource = AdoArt
        estado = 0
        'CARGA
        swnuevo = False
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
    AdoArt.Recordset.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub CmdImpCabeza_Click()
  Dim IResult As Integer
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
    IResult = .PrintReport
    If IResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim Nodo As Node
    Me.Top = 0
    Me.Left = 0
    Screen.MousePointer = vbHourglass
    ' Cargamos el Arbol
    ' Nodo Principal
    Set Nodo = trv.Nodes.Add(, , "Grupo", "Grupos", "Grupos")
    Nodo.Expanded = True
    Nodo.Bold = True
    
    Set rsMarcas = New ADODB.Recordset
    rsMarcas.Open "SELECT * FROM Al_Marcas ORDER BY descripcion", db, adOpenStatic
    Set AdoMarca.Recordset = rsMarcas
    
    Set rsunidad = New ADODB.Recordset
    rsunidad.Open "Select * from Al_UnidadMedida order by descripcion", db, adOpenStatic
    Set AdoMedida.Recordset = rsunidad
    
    Set rsMontador = New ADODB.Recordset
    rsMontador.Open "select * from al_montador order by descripcion", db, adOpenStatic
    Set AdoMontador.Recordset = rsMontador
    
    Set rsgrupo = New ADODB.Recordset
    rsgrupo.Open "SELECT * FROM ALClGrupo ORDER BY CAST (CodGrupo AS INT) ", db, adOpenStatic
    If rsgrupo.RecordCount > 0 Then
      rsgrupo.MoveFirst
      While Not rsgrupo.EOF
        Set Nodo = trv.Nodes.Add("Grupo", tvwChild, "D" & Trim(rsgrupo!CodGrupo), rsgrupo!descgrupo, "NoElegido", "Elegido")
        rsgrupo.MoveNext
      Wend
    Else
        trv.Nodes(1).Text = "No Existen Grupos Creados..."
    End If
    '
    Screen.MousePointer = vbDefault
	Call SeguridadSet(Me)
End Sub

Private Function valida() As Boolean
    valida = False
    If Trim(TxtGrupo.Text) = "" Then
        MsgBox "Elija el Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
        trv.SetFocus
        Exit Function
    End If
    If Trim(TxtDetalle.Text) = "" Then
        MsgBox "Ingrese el Codigo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtDetalle.SetFocus
        Exit Function
    End If
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Detalle.", vbExclamation + vbOKOnly, "Atención"
        txtDescripcion.SetFocus
        Exit Function
    End If
    If Trim(Unidad.Text) = "" Then
        MsgBox "Ingrese la Unidad de Medida.", vbExclamation + vbOKOnly, "Atención"
        Unidad.SetFocus
        Exit Function
    End If
    'alb
'    If Trim(txtUnidadCaja.Text) = "" Then
'        MsgBox "Ingrese la Unidad por Caja del Detalle.", vbExclamation + vbOKOnly, "Atención"
'        TxtActual.SetFocus
'        Exit Function
'    End If
    If Trim(txtStockMin.Text) = "" Then
        MsgBox "Ingrese el Stock Mínimo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtActual.SetFocus
        Exit Function
    End If
    valida = True
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub

Private Sub Imprimir_Click()
  Dim IResult As Integer
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
    IResult = .PrintReport
    If IResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub

Private Sub marcas_Click(Area As Integer)
    TDBC_marcas.BoundText = marcas.BoundText
End Sub

Private Sub TDBC_marcas_Click(Area As Integer)
    marcas.BoundText = TDBC_marcas.BoundText
End Sub

Private Sub Montador_Click(Area As Integer)
    TDBC_Montador.BoundText = Montador.BoundText
End Sub

Private Sub TDBC_Montador_Click(Area As Integer)
    Montador.BoundText = TDBC_Montador.BoundText
End Sub

Private Sub TDBC_Unidad_Click(Area As Integer)
    Unidad.BoundText = TDBC_Unidad.BoundText
End Sub

Private Sub Unidad_Click(Area As Integer)
    TDBC_Unidad.BoundText = Unidad.BoundText
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
    If InStr(Node.Key, "G") = 0 Then
        TxtGrupo.Text = Mid(Node.Key, 2)
    Else
        TxtGrupo.Text = ""
    End If
End Sub

Private Sub BuscaNodo(QNodo As String)
Dim Nodo As Node
Dim Indice As Integer
    For Indice = 1 To trv.Nodes.Count
        If Mid(trv.Nodes(Indice).Key, 2) = QNodo Then
            trv.Nodes(Indice).Selected = True
            Exit For
        End If
    Next
End Sub

Private Sub txtStockMin_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
End Sub
Private Sub txtUnidadCaja_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
End Sub
Private Function ExisteDetalle(codDetalle As String) As Boolean
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ALIngresoAlmDet WHERE CodArt = '" & codDetalle & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteDetalle = rs!Cuantos > 0
End Function

Private Sub CARGA()
            Set rsMarcas = New ADODB.Recordset
            If AdoMarca.Recordset.State = 1 Then AdoMarca.Recordset.Close
            rsMarcas.Open "SELECT * FROM Al_Marcas ORDER BY descripcion", db, adOpenStatic
            Set AdoMarca.Recordset = rsMarcas
            
            Set rsunidad = New ADODB.Recordset
            If AdoMedida.Recordset.State = 1 Then AdoMedida.Recordset.Close
            rsunidad.Open "Select * from Al_UnidadMedida order by descripcion", db, adOpenStatic
            Set AdoMedida.Recordset = rsunidad
    
'            Set rsMontador = New ADODB.Recordset
'            If AdoMontador.Recordset.State = 1 Then AdoMontador.Recordset.Close
'            rsMontador.Open "select * from al_montador order by descripcion", db, adOpenStatic
'            Set AdoMontador.Recordset = rsMontador
End Sub


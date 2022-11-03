VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlFrmCreaMaterial 
   Caption         =   "Clasificadores - Almacenes -  Bienes(Productos)"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   120
   ClientWidth     =   11145
   Icon            =   "AlFrmCreaMaterial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   880
      Left            =   0
      Picture         =   "AlFrmCreaMaterial.frx":6852
      ScaleHeight     =   825
      ScaleWidth      =   15180
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   0
      Width           =   15240
      Begin VB.Label LblCabecera 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE BIENES Y SERVICIOS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   405
         Index           =   0
         Left            =   9240
         TabIndex        =   68
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   1680
         TabIndex        =   67
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   950
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   15360
      Begin VB.OptionButton ABREDETALLE2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1200
         TabIndex        =   64
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton ABREDETALLE 
         BackColor       =   &H00C0E0FF&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2880
         TabIndex        =   63
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton CmdLista 
         Caption         =   "Listado de Productos"
         Height          =   720
         Left            =   12360
         Picture         =   "AlFrmCreaMaterial.frx":83F8
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   180
         Width           =   885
      End
      Begin VB.CommandButton cmdAprueba 
         BackColor       =   &H0080C0FF&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   10080
         Picture         =   "AlFrmCreaMaterial.frx":9B7A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Aprueba Registro"
         Top             =   180
         Width           =   770
      End
      Begin VB.CommandButton CmdFoto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Foto"
         Height          =   720
         Left            =   9240
         Picture         =   "AlFrmCreaMaterial.frx":9D84
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton CmdAnadir 
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   5880
         Picture         =   "AlFrmCreaMaterial.frx":A786
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Nuevo Registro"
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton Imprimir 
         Caption         =   "Inventario Fisico"
         Height          =   720
         Left            =   13320
         Picture         =   "AlFrmCreaMaterial.frx":11274
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton CmdImpCabeza 
         Caption         =   "Inventario Valorado"
         Height          =   720
         Left            =   11400
         Picture         =   "AlFrmCreaMaterial.frx":129F6
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   6720
         Picture         =   "AlFrmCreaMaterial.frx":14178
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Anular"
         Height          =   720
         Left            =   7560
         Picture         =   "AlFrmCreaMaterial.frx":14382
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   8400
         Picture         =   "AlFrmCreaMaterial.frx":14A6C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   14280
         Picture         =   "AlFrmCreaMaterial.frx":14C76
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO DE REGISTROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1080
         TabIndex        =   65
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame FraGraba 
      BackColor       =   &H00C0E0FF&
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
      Height          =   975
      Left            =   10
      TabIndex        =   58
      Top             =   840
      Visible         =   0   'False
      Width           =   15195
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   7080
         Picture         =   "AlFrmCreaMaterial.frx":14E80
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   5640
         Picture         =   "AlFrmCreaMaterial.frx":1508A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc AdoArt 
      Height          =   375
      Left            =   10
      Top             =   7200
      Width           =   5595
      _ExtentX        =   9869
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
      BackColor       =   12640511
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
   Begin VB.Frame FraArticulos 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5730
      Left            =   5730
      TabIndex        =   25
      Top             =   1840
      Width           =   9645
      Begin VB.TextBox TxtPrecVenta 
         Alignment       =   2  'Center
         DataField       =   "Precio_salon"
         DataSource      =   "AdoArt"
         Height          =   285
         Left            =   5400
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   4440
         Width           =   1365
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "DescDetalle"
         DataSource      =   "AdoArt"
         Height          =   405
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2480
         Width           =   7575
      End
      Begin VB.PictureBox Img_Foto 
         Height          =   1455
         Left            =   7680
         ScaleHeight     =   1395
         ScaleWidth      =   1755
         TabIndex        =   55
         Top             =   240
         Width           =   1815
         Begin VB.Image Image2 
            Height          =   1404
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1765
         End
      End
      Begin VB.TextBox TxtDescripcion2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Nombre_Anterior"
         DataSource      =   "AdoArt"
         Height          =   405
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2880
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Fecha_Vencimiento"
         DataSource      =   "AdoArt"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   5760
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   40245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "StockSalida"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   5190
         Width           =   1365
      End
      Begin VB.TextBox TxtPrecEst 
         Alignment       =   2  'Center
         DataField       =   "Precio_estimado"
         DataSource      =   "AdoArt"
         Height          =   285
         Left            =   8085
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   4440
         Width           =   1365
      End
      Begin VB.TextBox txtStockMin 
         Alignment       =   2  'Center
         DataField       =   "StockMin"
         DataSource      =   "AdoArt"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   5160
         Width           =   1365
      End
      Begin VB.TextBox TxtPrecComp 
         Alignment       =   2  'Center
         DataField       =   "precio_compra"
         DataSource      =   "AdoArt"
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   4440
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo Montador 
         Bindings        =   "AlFrmCreaMaterial.frx":15294
         DataField       =   "COD_MONTADOR"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   2400
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ListField       =   "COD_MONTADOR"
         BoundColumn     =   "COD_MONTADOR"
         Text            =   "Elige Marca..."
      End
      Begin VB.TextBox TxtInicial 
         Alignment       =   2  'Center
         DataField       =   "Cod_Ant"
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "0"
         Top             =   4440
         Width           =   1365
      End
      Begin VB.TextBox TxtSub 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "11"
         Top             =   1365
         Width           =   1215
      End
      Begin VB.TextBox txtCantVendida 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "StockIngreso"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoArt"
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   5190
         Width           =   1365
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "111"
         Top             =   2480
         Width           =   1815
      End
      Begin VB.TextBox TxtGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "11"
         Top             =   435
         Width           =   855
      End
      Begin VB.CheckBox chkEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   8085
         TabIndex        =   15
         Top             =   5040
         Width           =   1305
      End
      Begin VB.TextBox TxtActual 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         DataField       =   "StockActual"
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
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "0"
         Top             =   5190
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo TDBC_marcas 
         Bindings        =   "AlFrmCreaMaterial.frx":152AE
         DataField       =   "COD_MARCA"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "COD_MARCA"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo marcas 
         Bindings        =   "AlFrmCreaMaterial.frx":152C5
         DataField       =   "COD_MARCA"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   1200
         TabIndex        =   37
         Top             =   3390
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ListField       =   "COD_MARCA"
         BoundColumn     =   "COD_MARCA"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo TDBC_Montador 
         Bindings        =   "AlFrmCreaMaterial.frx":152DC
         DataField       =   "COD_MONTADOR"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   1365
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "COD_MONTADOR"
         Text            =   "Elige Marca..."
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
      Begin MSDataListLib.DataCombo TDBC_Unidad 
         Bindings        =   "AlFrmCreaMaterial.frx":152F6
         DataField       =   "Unidad"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   3720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "Unidad"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo Unidad 
         Bindings        =   "AlFrmCreaMaterial.frx":1530E
         DataField       =   "Unidad"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   4965
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "Unidad"
         BoundColumn     =   "Unidad"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoCod 
         Bindings        =   "AlFrmCreaMaterial.frx":15326
         DataField       =   "CodGrupo"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   2400
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ListField       =   "CODGrupo"
         BoundColumn     =   "CodGrupo"
         Text            =   "Elige Grupo ..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoDes 
         Bindings        =   "AlFrmCreaMaterial.frx":1533D
         DataField       =   "CodGrupo"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   435
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "DescGrupo"
         BoundColumn     =   "CodGrupo"
         Text            =   "Elige Grupo ..."
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
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Fecha_Alerta"
         DataSource      =   "AdoArt"
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   40245
      End
      Begin MSDataListLib.DataCombo DtcPaisD 
         Bindings        =   "AlFrmCreaMaterial.frx":15354
         DataField       =   "pais_codigo"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   6720
         TabIndex        =   7
         Top             =   3720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "pais_descripcion"
         BoundColumn     =   "pais_codigo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo DtcPais 
         Bindings        =   "AlFrmCreaMaterial.frx":1536A
         DataField       =   "pais_codigo"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   8760
         TabIndex        =   53
         Top             =   3360
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
         Bindings        =   "AlFrmCreaMaterial.frx":15380
         DataField       =   "CodGrupo"
         DataSource      =   "AdoArt"
         Height          =   315
         Left            =   3480
         TabIndex        =   69
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ListField       =   "codigo_unidad"
         BoundColumn     =   "CodGrupo"
         Text            =   "Elige Grupo ..."
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Mínimo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Características del Bien"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Industria (Pais Origen)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   52
         Top             =   3480
         Width           =   1545
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alerta Temprana:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4245
         TabIndex        =   51
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Primer Vencimiento:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   5760
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Total Vendida"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4035
         TabIndex        =   49
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Total Compra"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2025
         TabIndex        =   48
         Top             =   4920
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prec.Venta Cliente"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8085
         TabIndex        =   46
         Top             =   4185
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUB-GRUPO"
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
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Referencia"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   4200
         Width           =   1320
      End
      Begin VB.Label TDBFrame3D6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Actual"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   42
         Top             =   4920
         Width           =   1365
      End
      Begin VB.Label TDBFrame3D7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Compra Base"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2640
         TabIndex        =   41
         Top             =   4185
         Width           =   1455
      End
      Begin VB.Label TDBFrame3D8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Venta Base"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   40
         Top             =   4185
         Width           =   1365
      End
      Begin VB.Label TDBFrame3D5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad de Medida"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3540
         TabIndex        =   36
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label TDBFrame3D9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label TDBFrame3D3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label TDBFrame3D4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Bien o Servicio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   2280
         Width           =   6600
      End
      Begin VB.Label TDBFrame3D1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE DEL BIEN / SERVICIO"
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1995
         Width           =   9375
      End
      Begin VB.Label TDBFrame3D2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GRUPO"
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
         Left            =   120
         TabIndex        =   31
         Top             =   135
         Width           =   7455
      End
   End
   Begin MSAdodcLib.Adodc AdoMontador 
      Height          =   375
      Left            =   3720
      Top             =   7680
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      Left            =   1800
      Top             =   7680
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Left            =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   15240
      TabIndex        =   26
      Top             =   10515
      Width           =   15240
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
   Begin MSComctlLib.ImageList imlMaterial 
      Left            =   4200
      Top             =   4200
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
            Picture         =   "AlFrmCreaMaterial.frx":15397
            Key             =   "Grupos"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrmCreaMaterial.frx":15C71
            Key             =   "NoElegido"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AlFrmCreaMaterial.frx":1654B
            Key             =   "Elegido"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   375
      Left            =   7920
      Top             =   7680
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Left            =   5880
      Top             =   7680
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Top             =   8160
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
   Begin MSDataGridLib.DataGrid tdbgArt 
      Bindings        =   "AlFrmCreaMaterial.frx":16E25
      Height          =   5295
      Left            =   0
      TabIndex        =   62
      Top             =   1920
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   9340
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
      Caption         =   "LISTA DE PRODUCTOS"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "CodGrupo"
         Caption         =   "Grupo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "COD_montador"
         Caption         =   "SubGrupo"
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
         DataField       =   "unidad"
         Caption         =   "Unidad_Med"
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
         DataField       =   "CodDetalle"
         Caption         =   "Cod-Prod"
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
         DataField       =   "DescDetalle"
         Caption         =   "Descripcion"
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
         DataField       =   "Estado_Registro"
         Caption         =   "Aprob."
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
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   510.236
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CryBBSS 
      Left            =   600
      Top             =   8160
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
      Top             =   8160
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
Dim rsUnidad As ADODB.Recordset
Dim rsMontador As ADODB.Recordset

Dim rsgrupo As ADODB.Recordset
Dim RsArt, rsPais As ADODB.Recordset
Dim rsNada As ADODB.Recordset
'--------
Dim estado As Integer ' 0 navegar, 1 Agregar, 2 Editar
Dim swnuevo As Boolean
Dim sino As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim marca1 As BookmarkEnum
'--
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim queryinicial As String

Public Sub ALPrincipal(QEstado As Integer)
    '
'    Screen.MousePointer = vbHourglass
'    estado = QEstado
'    '
'    Select Case estado
'        Case 0
'            Set RsArt = New ADODB.Recordset
'            'JQA 04/2008
'            'GlSqlAux = "SELECT * FROM ALCLDetalle WHERE CAST(CODGRUPO AS INT)< 50  AND coddetalle = ISNULL(coddetalle, NULL) ORDER BY CAST (CODGRUPO AS INT)"
'            'GlSqlAux = "SELECT * FROM ALCLDetalle WHERE coddetalle = ISNULL(coddetalle, NULL) ORDER BY CODGRUPO, cod_montador, codDetalle "
'            queryinicial = "SELECT * FROM ALCLDetalle WHERE coddetalle = ISNULL(coddetalle, NULL) ORDER BY CODGRUPO, cod_montador, DescDetalle "
'            RsArt.Open queryinicial, db, adOpenDynamic, adLockOptimistic
'            If RsArt.RecordCount > 0 Then
'               GlHayRegs = True  'Variable global
'            Else
'               GlHayRegs = False
'            End If
'            BotonesNavegar Me
'            FraArticulos.Enabled = False
'            Set AdoArt.Recordset = RsArt
'        Case 1
'
'        Case 2
'
'    End Select
'    '
'    Screen.MousePointer = vbDefault
'    Me.Show
End Sub

Private Sub AdoArt_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Dim Marca As String
'Dim a As Integer
'Dim COD_MARCAx, cod_UMedida As String
If AdoArt.Recordset.BOF Or AdoArt.Recordset.EOF Then
        If AdoArt.Recordset.BOF And AdoArt.Recordset.EOF Then
            TxtGrupo.Text = ""
            TxtDetalle.Text = ""
            TxtDescripcion.Text = ""
            TxtActual.Text = ""
            chkEstado.Value = vbUnchecked
            AdoArt.Caption = "Registro: 0 de 0"
'            BuscaNodo "Grupo"
        Else
            Exit Sub
        End If
Else
'    If swnuevo = False Then
'            If Not (IsNull(AdoMarca.Recordset("cod_marca"))) Then
'                If Not (AdoMarca.Recordset.BOF) Then AdoMarca.Recordset.MoveFirst
'                AdoMarca.Recordset.Find "cod_marca ='" & AdoArt.Recordset!COD_MARCA & "'", , adSearchForward
'                If Not AdoMarca.Recordset.EOF Then
'                    'TDBC_marcas.Item(1) = AdoMarca.Recordset!descripcion
'                    TDBC_marcas.Text = AdoMarca.Recordset!descripcion
'                End If
'            End If
'            If Not (IsNull(AdoMontador.Recordset("cod_montador"))) Then
'                If Not (AdoMontador.Recordset.BOF) Then AdoMontador.Recordset.MoveFirst
'                AdoMontador.Recordset.Find "Cod_montador ='" & AdoArt.Recordset!cod_MONTADOR & "'", , adSearchForward
'                If Not AdoMontador.Recordset.EOF Then
'                    '
'                End If
'            End If
'            If Not (IsNull(AdoMedida.Recordset("Unidad"))) Then
'                If Not (AdoMedida.Recordset.BOF) Then AdoMedida.Recordset.MoveFirst
'                    AdoMedida.Recordset.Find "Unidad ='" & AdoArt.Recordset!Unidad & "'", , adSearchForward
'                If Not AdoMedida.Recordset.EOF Then
'                    '
'                End If
'            End If
        If AdoArt.Recordset!StockMin < AdoArt.Recordset!StockActual Then
            TxtActual.BackColor = &HE0E0E0
        Else
            TxtActual.BackColor = &HFF&
        End If
'    End If
        'TDBC_Montador
    Set Img_Foto = Leer_Imagen(db, "Select Foto From ALCLDetalle Where CodDetalle = '" & AdoArt.Recordset("CodDetalle") & "' ", "Foto")
    Image2 = Img_Foto
    If AdoArt.Recordset!estado_registro = "SI" Then
        chkEstado.Value = vbChecked
        CmdFoto.Visible = True
    Else
        CmdFoto.Visible = False
        chkEstado.Value = vbUnchecked
    End If
        'chkEstado.Value =IIf(CBool(AdoArt.Recordset!estado), vbChecked, vbUnchecked)
'        BuscaNodo AdoArt.Recordset!CodGrupo
    
End If
End Sub

Private Sub CmdAnadir_Click()
    swnuevo = True
    Set tdbgArt.DataSource = rsNada
    AdoArt.Recordset.AddNew
    estado = 1
'    BotonesEditar Me
    FraOpciones.Visible = False
    FraGraba.Visible = True
    tdbgArt.Enabled = False
    FraArticulos.Enabled = True
    TxtGrupo.Enabled = False
    DtcGrupoDes.Enabled = True
    TxtSub.Enabled = False
    TDBC_Montador.Enabled = False
'    trv.SetFocus
'    BuscaNodo "grupo"
    txtStockMin.Text = 0
End Sub

Private Sub cmdAprueba_Click()
   sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If AdoArt.Recordset("Estado") = 0 Then
      If sino = vbYes Then
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
        AdoArt.Recordset("estado") = 1
        AdoArt.Recordset("Estado_Registro") = "SI"
        AdoArt.Recordset("fecha_registro") = Date
        AdoArt.Recordset("usr_usuario") = GlUsuario
        AdoArt.Recordset.Update
      End If
   Else
        MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub CmdBuscar_Click()
'  Set ClBuscaGrid = New ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.QueryUtilizado = GlSqlAux
'  ClBuscaGrid.Título = "Elija un Detalle"
'  ClBuscaGrid.EsTdbGrid = True
'  Set ClBuscaGrid.GridTrabajo = tdbgArt
'  Set ClBuscaGrid.RecordsetTrabajo = AdoArt.Recordset
'  ClBuscaGrid.Ejecutar
''  If ClBuscaGrid.ElegidoCol1 <> "" Then
''    AdoArt.Recordset.Filter = adFilterNone
''    AdoArt.Recordset.MoveFirst
''    AdoArt.Recordset.Find "CodGrupo + '-' + CodDetalle   = " & ClBuscaGrid.ElegidoCol1 & " - " & ClBuscaGrid.ElegidoCol2 & ""
'  End If
  PosibleApliqueFiltro = False
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = tdbgArt
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = AdoArt.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True

End Sub
Private Sub CmdCancelar_Click()

On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If AdoArt.Recordset.EditMode <> adEditNone Then AdoArt.Recordset.CancelUpdate
    Call ABREDETALLE_Click
    Call CARGA
    AdoArt.Caption = "Registro: " & CStr(AdoArt.Recordset.AbsolutePosition) & " de " & AdoArt.Recordset.RecordCount
    'BotonesNavegar Me
    FraOpciones.Visible = True
    FraGraba.Visible = False
    FraArticulos.Enabled = False
    TxtGrupo.Enabled = True
    DtcGrupoDes.Enabled = True
    TxtSub.Enabled = True
    TDBC_Montador.Enabled = True
'    Set tdbgArt.DataSource = AdoArt
    Screen.MousePointer = vbDefault
    estado = 0
    swnuevo = False
    tdbgArt.Enabled = True
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub Cmdeditar_Click()
On Error GoTo Que_Error
    swnuevo = False
    Screen.MousePointer = vbHourglass
    'BotonesEditar Me
    estado = 2
    FraOpciones.Visible = False
    FraGraba.Visible = True
    FraArticulos.Enabled = True
    TxtGrupo.Enabled = False
    TDBC_Montador.Enabled = False
    TxtSub.Enabled = False
    If AdoArt.Recordset!estado_registro = "NO" Then
        DtcGrupoDes.Enabled = True
        TxtDetalle.Enabled = True
    Else
        DtcGrupoDes.Enabled = False
        TxtDetalle.Enabled = False
    End If
    AdoArt.Caption = "Editando Registro..."
    Screen.MousePointer = vbDefault
    tdbgArt.Enabled = False
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo Que_Error
    'ao_adjudica_detalle_D
   If AdoArt.Recordset.RecordCount > 0 Then
      If ExisteDetalle(AdoArt.Recordset!codDetalle) Then MsgBox "No se puede eliminar un BIEN o SERVICIO que ya tiene Registros en COMPRAS o ALMACEN.", vbInformation + vbOKOnly, "Atención": Exit Sub
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         'AdoArt.Recordset.Delete
         AdoArt.Recordset!estado = "ER"
         AdoArt.Recordset.Update
         AdoArt.Recordset.Requery
      End If
   Else
        MsgBox "No existen registros para Anular.", vbExclamation, "Atención"
   End If
   Exit Sub
    
'    If Not GlHayRegs Then
'        MsgBox "No existen registro para Anular", vbExclamation + vbOKOnly, "Atención"
'        Exit Sub
'    End If
'    If ExisteDetalle(AdoArt.Recordset!CodGrupo & "-" & AdoArt.Recordset!codDetalle) Then MsgBox "No se puede eliminar el Detalle seleccionado ya que se tiene registro de Movimientos en Almacen.", vbInformation + vbOKOnly, "Atención": Exit Sub
'    If MsgBox("¿ Está seguro que se va a Anular el registro visualizado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
'        Screen.MousePointer = vbHourglass
'        'AdoArt.Recordset.Delete
'        AdoArt.Recordset!estado = 2
'        AdoArt.Recordset.MoveNext
'        If AdoArt.Recordset.EOF Then
'          If AdoArt.Recordset.RecordCount > 0 Then
'            AdoArt.Recordset.MoveLast
'          Else
'            GlHayRegs = False
'            AdoArt.Refresh
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

Private Sub CmdFoto_Click()
  On Error GoTo QError
    If AdoArt.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\BIENES\" & Trim(AdoArt.Recordset!CodGrupo) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOTB"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(AdoArt.Recordset!iniciales) & "-" & Trim(AdoArt.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\BIENES\" & Trim(AdoArt.Recordset!CodGrupo) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOTB"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(AdoArt.Recordset!iniciales) & "-" & Trim(AdoArt.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(AdoArt.Recordset!iniciales) + "-" + Trim(AdoArt.Recordset("codigo_beneficiario")) + "\" + Trim(AdoArt.Recordset!ARCHIVO_FOTO)
'    Else
        ARCH_FOTO = App.Path + "\BIENES\" + Trim(AdoArt.Recordset!CodGrupo) + "\" + Trim(AdoArt.Recordset!ARCHIVO_FOTO)
'    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + AdoArt.Recordset!codigo_beneficiario + "\" + AdoArt.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    CodBien = AdoArt.Recordset!codDetalle
    If Guardar_Imagen(db, "Select Foto From ALCLDetalle Where CodDetalle= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
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

Private Sub CmdGrabar_Click()
On Error GoTo QError
   If valida Then
      Screen.MousePointer = vbHourglass
        ' Empezar a grabar
        '*********************************
      db.BeginTrans
        'JQA 04/2008
      If swnuevo = True Then
        AdoArt.Recordset!CodGrupo = Trim(TxtGrupo.Text)
        AdoArt.Recordset!cod_MONTADOR = Trim(Montador.Text)
        AdoArt.Recordset!codDetalle = Trim(TxtDetalle.Text)
        AdoArt.Recordset!ARCHIVO_FOTO = "Cargar_Archivo"
        AdoArt.Recordset!DescDetalle = TxtDescripcion.Text + " - " + TxtInicial
      End If
      If swnuevo = False Then
        AdoArt.Recordset!DescDetalle = TxtDescripcion.Text
      End If
        AdoArt.Recordset!Nombre_Anterior = TxtDescripcion2.Text
        AdoArt.Recordset!Unidad = IIf(Unidad.Text = "", "UNI", Unidad.Text)
        AdoArt.Recordset!COD_MARCA = IIf(marcas.Text = "", "S/N", marcas.Text)
        ' Campos no ligados
        'AdoArt.Recordset!estado = IIf(chkEstado.Value = vbChecked, 1, 0)
'        AdoArt.Recordset!StockInicial = IIf(TxtInicial.Text = "", 0, Val(TxtInicial.Text))      'Val(TxtInicial.Text)
        AdoArt.Recordset!cod_ant = TxtInicial.Text
        AdoArt.Recordset!COD_UNIV = DtcGrupoUni.Text
        AdoArt.Recordset!Precio_Compra = IIf(TxtPrecComp.Text = "", 0, CDbl(TxtPrecComp.Text))      'CDbl(TxtPrecComp.Text)
        AdoArt.Recordset!Precio_salon = IIf(TxtPrecVenta.Text = "", 0, CDbl(TxtPrecVenta.Text))      'CDbl(txtStockMin)
        AdoArt.Recordset!Precio_estimado = IIf(TxtPrecEst.Text = "", 0, CDbl(TxtPrecEst.Text))      'CDbl(TxtPrecEst)
        AdoArt.Recordset!StockMin = IIf(txtStockMin.Text = "", 0, CDbl(txtStockMin.Text))      'CDbl(txtStockMin)
        AdoArt.Recordset!pais_codigo = DtcPais.Text
        'AdoArt.Recordset!ARCHIVO_F = Trim(AdoArt.Recordset!cod_MONTADOR) + "-" + Trim(AdoArt.Recordset!codDetalle) + ".JPG"
        AdoArt.Recordset!ARCHIVO_F = Trim(AdoArt.Recordset!codDetalle) + ".JPG"
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
        'BotonesNavegar Me
        FraOpciones.Visible = True
        FraGraba.Visible = False
        FraArticulos.Enabled = False
        TxtGrupo.Enabled = True
        DtcGrupoDes.Enabled = True
        TxtSub.Enabled = True
        TDBC_Montador.Enabled = True
        Screen.MousePointer = vbDefault
        marca1 = AdoArt.Recordset.BookMark
        Call ABREDETALLE_Click
        Call CARGA
        AdoArt.Recordset.Move marca1 - 1
        'AdoArt.Recordset.MoveLast
        'Set tdbgArt.DataSource = AdoArt
        estado = 0
        'CARGA
        swnuevo = False
        tdbgArt.Enabled = True
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

Private Sub CmdLista_Click()
  Dim IResult As Integer
  With CryLista
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True

        .ReportFileName = App.Path & "\Reportes\Almacen\Productos_Todos.rpt"
    IResult = .PrintReport
    If IResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With
End Sub

Private Sub CmdSalir_Click()
    Unload Me
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
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from AL_Montador where CodGrupo='" & CodFuente & "'"
   
   Set Montador.RowSource = Nothing
   Set Montador.RowSource = db.Execute(strConsultaF, , adCmdText)
   Montador.ReFill
   Montador.BoundText = Empty
   
   Set TDBC_Montador.RowSource = Nothing
   Set TDBC_Montador.RowSource = db.Execute(strConsultaF, , adCmdText)
   TDBC_Montador.ReFill
   TDBC_Montador.BoundText = Empty

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
'    Set Nodo = trv.Nodes.Add(, , "Grupo", "Grupos", "Grupos")
'    Nodo.Expanded = True
'    Nodo.Bold = True
    ABREDETALLE = True
    Call ABREDETALLE_Click
    Call CARGA
'    Set rsgrupo = New ADODB.Recordset
'    rsgrupo.Open "SELECT * FROM ALClGrupo ORDER BY CAST (CodGrupo AS INT) ", db, adOpenStatic
'    Set AdoGrupo.Recordset = rsgrupo
'    If rsgrupo.RecordCount > 0 Then
'      rsgrupo.MoveFirst
'      While Not rsgrupo.EOF
'        Set Nodo = trv.Nodes.Add("Grupo", tvwChild, "D" & Trim(rsgrupo!CodGrupo), rsgrupo!descgrupo, "NoElegido", "Elegido")
'        rsgrupo.MoveNext
'      Wend
'    Else
'        trv.Nodes(1).Text = "No Existen Grupos Creados..."
'    End If
    '
    'BotonesNavegar Me
    FraOpciones.Visible = True
    FraGraba.Visible = False
    FraArticulos.Enabled = False
    Screen.MousePointer = vbDefault
	Call SeguridadSet(Me)
End Sub

Private Sub ABREDETALLE_Click()
    Set RsArt = New ADODB.Recordset
    'JQA 04/2008
    If RsArt.State = 1 Then RsArt.Close
    'queryinicial = "SELECT * FROM ALCLDetalle WHERE Estado <> 2 "   'ORDER BY CODGRUPO, cod_montador, DescDetalle
    queryinicial = "SELECT * FROM ALCLDetalle where ESTADO_registro <> 'ER' "   'ORDER BY CODGRUPO, cod_montador, DescDetalle
    RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    RsArt.Sort = "CodGrupo, COD_montador, DescDetalle"
    If RsArt.RecordCount > 0 Then
       GlHayRegs = True  'Variable global
    Else
       GlHayRegs = False
    End If
    Set AdoArt.Recordset = RsArt
    'Set tdbgArt.DataSource = AdoArt.Recordset
'    AdoArt.Recordset.Requery
'    AdoArt.Refresh
    Set tdbgArt.DataSource = AdoArt
End Sub

Private Sub ABREDETALLE2_Click()
    Set RsArt = New ADODB.Recordset
    'JQA 04/2008
    If RsArt.State = 1 Then RsArt.Close
    'queryinicial = "SELECT * FROM ALCLDetalle WHERE Estado <> 2 "   'ORDER BY CODGRUPO, cod_montador, DescDetalle
    queryinicial = "SELECT * FROM ALCLDetalle WHERE Estado_registro='NO' "   'ORDER BY CODGRUPO, cod_montador, DescDetalle
    RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    'RsArt.Sort = "CodGrupo, COD_montador"
    RsArt.Sort = "CodGrupo, COD_montador, DescDetalle"
    If RsArt.RecordCount > 0 Then
       GlHayRegs = True  'Variable global
    Else
       GlHayRegs = False
    End If
    Set AdoArt.Recordset = RsArt
    'Set tdbgArt.DataSource = AdoArt.Recordset
'    AdoArt.Recordset.Requery
'    AdoArt.Refresh
    Set tdbgArt.DataSource = AdoArt
End Sub

Private Function valida() As Boolean
    valida = False
    If Trim(TxtGrupo.Text) = "" Then
        MsgBox "Elija el Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
        DtcGrupoDes.SetFocus
        Exit Function
    End If
    If Trim(TDBC_Montador.Text) = "" Then
        MsgBox "Elija el Sub-Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
        DtcGrupoDes.SetFocus
        Exit Function
    End If
    If Trim(TxtDetalle.Text) = "" Then
        MsgBox "Ingrese el Codigo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtDetalle.SetFocus
        Exit Function
    End If
    If Trim(TxtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtDescripcion.SetFocus
        Exit Function
    End If
    If Trim(Unidad.Text) = "" Then
        MsgBox "Ingrese la Unidad de Medida.", vbExclamation + vbOKOnly, "Atención"
        Unidad.SetFocus
        Exit Function
    End If
    If Trim(TxtPrecComp.Text) = "" Then
        MsgBox "Ingrese EL Precio de Compra del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtPrecComp.SetFocus
        Exit Function
    End If
    If Trim(txtStockMin.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Salon del Detalle.", vbExclamation + vbOKOnly, "Atención"
        txtStockMin.SetFocus
        Exit Function
    End If
    If Trim(TxtPrecEst.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Cliente del Detalle.", vbExclamation + vbOKOnly, "Atención"
        TxtPrecEst.SetFocus
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

'Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
'    If InStr(Node.Key, "G") = 0 Then
'        TxtGrupo.Text = Mid(Node.Key, 2)
'    Else
'        TxtGrupo.Text = ""
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

Private Function ExisteDetalle(codDetalle As String) As Boolean
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ALIngresoAlmDet WHERE CodArt = '" & codDetalle & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_no_objecion_detalle_D WHERE codDetalle = '" & codDetalle & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteDetalle = rs!Cuantos > 0
End Function

Private Sub CARGA()
    Set rsMarcas = New ADODB.Recordset
    If rsMarcas.State = 1 Then rsMarcas.Close
    rsMarcas.Open "SELECT * FROM Al_Marcas ORDER BY descripcion", db, adOpenStatic
    Set AdoMarca.Recordset = rsMarcas
    
    Set rsUnidad = New ADODB.Recordset
    If rsUnidad.State = 1 Then rsUnidad.Close
    rsUnidad.Open "Select * from Al_UnidadMedida order by descripcion", db, adOpenStatic
    Set AdoMedida.Recordset = rsUnidad
    
    Set rsMontador = New ADODB.Recordset
    If rsMontador.State = 1 Then rsMontador.Close
    rsMontador.Open "select * from al_montador order by descripcion", db, adOpenStatic
    Set AdoMontador.Recordset = rsMontador
    
    Set rsgrupo = New ADODB.Recordset
    If rsgrupo.State = 1 Then rsgrupo.Close
    rsgrupo.Open "SELECT * FROM ALClGrupo WHERE activo='S' ", db, adOpenStatic
    Set AdoGRUPO.Recordset = rsgrupo
    
    Set rsPais = New ADODB.Recordset
    If rsPais.State = 1 Then rsPais.Close
    rsPais.Open "SELECT * FROM gc_pais WHERE estado_registro='S' order by pais_descripcion", db, adOpenStatic
    Set AdoPais.Recordset = rsPais
End Sub


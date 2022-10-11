VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form AlFrmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Procesos Administrativos - Almacen - Administracion de Almacenes"
   ClientHeight    =   9405
   ClientLeft      =   1560
   ClientTop       =   1725
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   39313.88
   ScaleMode       =   0  'User
   ScaleWidth      =   1.14734e5
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryLista 
      Left            =   0
      Top             =   9480
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
      Left            =   480
      Top             =   9480
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
      Left            =   960
      Top             =   9480
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
   Begin VB.Frame Frmnavega 
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
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   696
      Width           =   11895
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo Registro"
         Height          =   480
         Left            =   960
         Picture         =   "AlFrmTransfer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Registra Nueva Transferencia"
         Top             =   1080
         Width           =   1245
      End
      Begin MSAdodcLib.Adodc adoTransf 
         Height          =   330
         Left            =   30
         Top             =   1560
         Width           =   3195
         _ExtentX        =   5636
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
         Caption         =   " <--  DESPLAZAR  -->"
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
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   1800
         TabIndex        =   3
         Top             =   795
         Width           =   795
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sin Aprobar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   360
         TabIndex        =   2
         Top             =   795
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DtgVentas 
         Bindings        =   "AlFrmTransfer.frx":058A
         Height          =   1785
         Left            =   3240
         TabIndex        =   9
         Top             =   120
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   3149
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
         Enabled         =   -1  'True
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
         Caption         =   "DATOS GENERALES DE LA TRANSFERENCIA"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "nro_transfer"
            Caption         =   "Nro.Transfer"
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
            DataField       =   "CodDestino"
            Caption         =   "Alm.Origen"
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
            DataField       =   "Codigo_beneficiario"
            Caption         =   "Responsable.Entrega"
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
            DataField       =   "CodDestino2"
            Caption         =   "Alm.Destino"
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
            DataField       =   "Codigo_beneficiario2"
            Caption         =   "Responsable.Recepcion"
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
            DataField       =   "fecha_transfer"
            Caption         =   "Fecha.Transfer"
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
            DataField       =   "estado_registro"
            Caption         =   "Aprob"
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
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1649.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1844.787
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   555.024
            EndProperty
         EndProperty
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
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame FrmDetalle 
      BackColor       =   &H00400000&
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
      Height          =   2295
      Left            =   0
      TabIndex        =   16
      Top             =   6105
      Width           =   11910
      Begin VB.CommandButton CmdListaAnula 
         BackColor       =   &H0080C0FF&
         Caption         =   "Anula Producto a Transferir -->"
         Height          =   585
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Anula Producto Existente"
         Top             =   1395
         Width           =   1365
      End
      Begin VB.CommandButton CmdListaMod 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modifica Prod. a Transferir --->"
         Height          =   585
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Modifica Producto Existente"
         Top             =   825
         Width           =   1380
      End
      Begin VB.CommandButton CmdLista 
         BackColor       =   &H0080C0FF&
         Caption         =   "Nuevo Producto a Transferir --->"
         Height          =   585
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Adicionar Productos"
         Top             =   240
         Width           =   1380
      End
      Begin MSDataGridLib.DataGrid DtGLista 
         Height          =   1980
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   3493
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
         Caption         =   "DATOS DE LOS PRODUCTOS TRANSFERIDOS"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "fecha_registro"
            Caption         =   "Fecha"
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
            DataField       =   "nro_transfer"
            Caption         =   "Nro.Transfer"
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
            DataField       =   "nro_licitacion"
            Caption         =   "Nro.Compra"
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
            DataField       =   "nro_lote"
            Caption         =   "Nro.Lote"
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
            DataField       =   "CodDetalle"
            Caption         =   "Producto"
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
            DataField       =   "DescDetalle"
            Caption         =   "Descripcion del Producto"
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
            DataField       =   "cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3929.953
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3450
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6085
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16576
      TabCaption(0)   =   "DATOS GENERALES DE LA TRANSFERENCIA"
      TabPicture(0)   =   "AlFrmTransfer.frx":05A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrmCabecera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmgrabcabeza"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmabm"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DATOS DE LOS PRODUCTOS TRANSFERIDOS"
      TabPicture(1)   =   "AlFrmTransfer.frx":05BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEdita"
      Tab(1).Control(1)=   "FrmGrabDet"
      Tab(1).ControlCount=   2
      Begin VB.Frame FrmEdita 
         BackColor       =   &H80000010&
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
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   -74880
         TabIndex        =   25
         Top             =   1320
         Width           =   11565
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   310
            HideSelection   =   0   'False
            Left            =   10855
            TabIndex        =   79
            Top             =   1080
            Width           =   265
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   310
            HideSelection   =   0   'False
            Left            =   10910
            TabIndex        =   78
            Top             =   360
            Width           =   265
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   310
            HideSelection   =   0   'False
            Left            =   8200
            TabIndex        =   77
            Top             =   360
            Width           =   265
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   310
            HideSelection   =   0   'False
            Left            =   5800
            TabIndex        =   76
            Top             =   360
            Width           =   265
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   310
            HideSelection   =   0   'False
            Left            =   3590
            TabIndex        =   75
            Top             =   360
            Width           =   265
         End
         Begin VB.CommandButton Cmd_PersNuevo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Elije Stock"
            Height          =   525
            Left            =   9960
            MaskColor       =   &H00C0FFFF&
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Nuevo Personal"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DtcNroCompra 
            Bindings        =   "AlFrmTransfer.frx":05DA
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   2280
            TabIndex        =   38
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "nro_licitacion"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCodMontador 
            Bindings        =   "AlFrmTransfer.frx":05F1
            CausesValidation=   0   'False
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   9840
            TabIndex        =   35
            Top             =   1080
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "cod_montador"
            BoundColumn     =   "CodDetalle"
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
         Begin MSDataListLib.DataCombo DtcGrupo 
            Bindings        =   "AlFrmTransfer.frx":0608
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   9000
            TabIndex        =   33
            Top             =   1080
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "CodGrupo"
            BoundColumn     =   "CodDetalle"
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
         Begin VB.TextBox TxtNroTrf 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            DataField       =   "nro_transfer"
            DataSource      =   "AdoTransf_Det"
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
            ForeColor       =   &H00FFFFC0&
            Height          =   405
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtCantidad 
            Alignment       =   2  'Center
            DataField       =   "cantidad"
            DataSource      =   "AdoTransf_Det"
            Height          =   285
            Left            =   6600
            TabIndex        =   8
            Text            =   "0"
            Top             =   1560
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DtcNroLote 
            Bindings        =   "AlFrmTransfer.frx":061F
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   4320
            TabIndex        =   26
            Top             =   360
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Nro_Lote"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtccodbien 
            Bindings        =   "AlFrmTransfer.frx":0636
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   6600
            TabIndex        =   27
            Top             =   1080
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "CodDetalle"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtcdesbien 
            Bindings        =   "AlFrmTransfer.frx":064D
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1080
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "DescDetalle"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcFechaVenc 
            Bindings        =   "AlFrmTransfer.frx":0664
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   6600
            TabIndex        =   52
            Top             =   360
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "fechaVenc"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcSaldoAct 
            Bindings        =   "AlFrmTransfer.frx":067B
            DataField       =   "CodDetalle"
            DataSource      =   "AdoTransf_Det"
            Height          =   315
            Left            =   9720
            TabIndex        =   55
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "StockActual"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Stock en Almacen Origen:"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   9240
            TabIndex        =   54
            Top             =   120
            Width           =   1860
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   6720
            TabIndex        =   53
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label lbltipoVenta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "ELIJA PRODUCTO A TRANSFERIR ..."
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
            Left            =   225
            TabIndex        =   51
            Top             =   1560
            Width           =   3270
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Compra:"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   2520
            TabIndex        =   39
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Lote"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   4560
            TabIndex        =   37
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-Grupo"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   9960
            TabIndex        =   36
            Top             =   840
            Width           =   885
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   9120
            TabIndex        =   34
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro.Transferencia:"
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
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Productos a Tranferir:"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   4320
            TabIndex        =   31
            Top             =   1560
            Width           =   2205
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción del Producto"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1785
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Producto"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   6720
            TabIndex        =   29
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame FrmGrabDet 
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
         Height          =   900
         Left            =   -74880
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
         Width           =   11565
         Begin VB.CommandButton cmdElige 
            BackColor       =   &H0080C0FF&
            Caption         =   "New Prod"
            Height          =   480
            Left            =   2040
            MaskColor       =   &H00000000&
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.CommandButton CmdGrabaDet 
            Caption         =   "Graba Producto Transferido"
            Height          =   675
            Left            =   3420
            Picture         =   "AlFrmTransfer.frx":0692
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Graba Datos de Producto"
            Top             =   120
            Width           =   2085
         End
         Begin VB.CommandButton CmdCancelaDet 
            Caption         =   "Cancela Prod. Transferido"
            Height          =   675
            Left            =   6105
            Picture         =   "AlFrmTransfer.frx":089C
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Cancela Grabación"
            Top             =   120
            Width           =   2070
         End
      End
      Begin VB.Frame frmabm 
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
         Height          =   930
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   11655
         Begin VB.CommandButton ImprProd 
            Caption         =   "Por Producto"
            Height          =   720
            Left            =   6000
            Picture         =   "AlFrmTransfer.frx":0AA6
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Imprime Nota de Transferencia"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton cmdAprueba 
            Caption         =   "Aprobar"
            Height          =   720
            Left            =   4080
            Picture         =   "AlFrmTransfer.frx":2228
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Aprueba Transferencia"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdEnviar 
            Caption         =   "Entregar"
            Height          =   720
            Left            =   9480
            Picture         =   "AlFrmTransfer.frx":2EF2
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Entrega Producto"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdDesaprueba 
            Caption         =   "Desapro."
            Height          =   720
            Left            =   4080
            Picture         =   "AlFrmTransfer.frx":3BBC
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdSalCabeza 
            Caption         =   "Salir"
            Height          =   720
            Left            =   10440
            Picture         =   "AlFrmTransfer.frx":3DC6
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Salir de Transferencias"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdImpCabeza 
            Caption         =   "Nota Transfer"
            Height          =   720
            Left            =   8520
            Picture         =   "AlFrmTransfer.frx":3FD0
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Imprime Nota de Transferencia"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdDelCabeza 
            Caption         =   "Anular"
            Height          =   720
            Left            =   2160
            Picture         =   "AlFrmTransfer.frx":5752
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Anula Transferencia Existente"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdBusCabeza 
            Caption         =   "Buscar"
            Height          =   720
            Left            =   3120
            Picture         =   "AlFrmTransfer.frx":641C
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Busca un Registro"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdModCabeza 
            Caption         =   "Modificar"
            Height          =   720
            Left            =   1200
            Picture         =   "AlFrmTransfer.frx":6CE6
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Modifica Transferencia Existente"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdAddCabeza 
            Caption         =   "Nuevo"
            Height          =   720
            Left            =   300
            Picture         =   "AlFrmTransfer.frx":75B0
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Registra Nueva Transferencia"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton ImprAlmacen 
            Caption         =   "Por Almacen"
            Height          =   720
            Left            =   5040
            Picture         =   "AlFrmTransfer.frx":E09E
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Ver Estado del Almacen"
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame frmgrabcabeza 
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
         Height          =   900
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Visible         =   0   'False
         Width           =   11685
         Begin VB.CommandButton CmdCanCabeza 
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   5865
            Picture         =   "AlFrmTransfer.frx":ED68
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdGraCabeza 
            Caption         =   "Grabar"
            Height          =   675
            Left            =   4740
            Picture         =   "AlFrmTransfer.frx":EF72
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame FrmCabecera 
         BackColor       =   &H80000010&
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
         Height          =   2070
         Left            =   120
         TabIndex        =   18
         Top             =   1290
         Width           =   11685
         Begin VB.Frame Frasolic 
            BackColor       =   &H8000000A&
            Caption         =   "Almacen Origen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   1335
            Left            =   60
            TabIndex        =   20
            Top             =   600
            Width           =   5485
            Begin MSDataListLib.DataCombo DtcRespNom 
               Bindings        =   "AlFrmTransfer.frx":F17C
               DataField       =   "codigo_beneficiario"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   840
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcResp 
               Bindings        =   "AlFrmTransfer.frx":F197
               DataField       =   "codigo_beneficiario"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   3480
               TabIndex        =   1
               Top             =   600
               Visible         =   0   'False
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483638
               ForeColor       =   -2147483624
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_AlmDes 
               Bindings        =   "AlFrmTransfer.frx":F1B2
               DataField       =   "CODDESTINO"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   5250
               _ExtentX        =   9260
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DESCDESTINO"
               BoundColumn     =   "CODDESTINO"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_Alm 
               Bindings        =   "AlFrmTransfer.frx":F1CB
               DataField       =   "CODDESTINO"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   1800
               TabIndex        =   41
               Top             =   480
               Visible         =   0   'False
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "CODDESTINO"
               BoundColumn     =   "CODDESTINO"
               Text            =   ""
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Responsable Entrega"
               ForeColor       =   &H00400000&
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   600
               Width           =   1530
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000A&
            Caption         =   "Almacen Destino"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   1350
            Left            =   6000
            TabIndex        =   44
            Top             =   600
            Width           =   5535
            Begin MSDataListLib.DataCombo DtcRespNom2 
               Bindings        =   "AlFrmTransfer.frx":F1E4
               DataField       =   "codigo_beneficiario2"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   120
               TabIndex        =   45
               Top             =   840
               Width           =   5370
               _ExtentX        =   9472
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcResp2 
               Bindings        =   "AlFrmTransfer.frx":F202
               DataField       =   "codigo_beneficiario2"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   3720
               TabIndex        =   46
               Top             =   600
               Visible         =   0   'False
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483638
               ForeColor       =   -2147483624
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_AlmDes2 
               Bindings        =   "AlFrmTransfer.frx":F220
               DataField       =   "CodDestino2"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   5370
               _ExtentX        =   9472
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DESCDESTINO"
               BoundColumn     =   "CODDESTINO"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtc_Alm2 
               Bindings        =   "AlFrmTransfer.frx":F23C
               DataField       =   "CodDestino2"
               DataSource      =   "adoTransf"
               Height          =   315
               Left            =   1920
               TabIndex        =   48
               Top             =   480
               Visible         =   0   'False
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "CODDESTINO"
               BoundColumn     =   "CODDESTINO"
               Text            =   ""
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               Caption         =   "Responsable Recepcion"
               ForeColor       =   &H00400000&
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   1755
            End
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Height          =   495
            Left            =   5520
            MaskColor       =   &H80000016&
            Picture         =   "AlFrmTransfer.frx":F258
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   1440
            Width           =   540
         End
         Begin VB.CommandButton Seleccionar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            Height          =   495
            Left            =   5520
            MaskColor       =   &H80000016&
            Picture         =   "AlFrmTransfer.frx":F3E2
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   720
            Width           =   540
         End
         Begin VB.TextBox TxtConcepto 
            DataField       =   "Observaciones"
            DataSource      =   "adoTransf"
            Height          =   285
            Left            =   3360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            Top             =   240
            Width           =   8175
         End
         Begin VB.TextBox txtnrosol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            DataField       =   "nro_transfer"
            DataSource      =   "adoTransf"
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   180
            TabIndex        =   19
            Top             =   280
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "fecha_transfer"
            DataSource      =   "adoTransf"
            Height          =   285
            Left            =   1560
            TabIndex        =   42
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   51773441
            CurrentDate     =   36464
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000010&
            Caption         =   "Nro.Transfer"
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
            Height          =   285
            Left            =   180
            TabIndex        =   24
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Observaciones"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3360
            TabIndex        =   23
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Fecha Transferencia"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1605
            TabIndex        =   22
            Top             =   60
            Width           =   1470
         End
      End
   End
   Begin Crystal.CrystalReport CryV01 
      Left            =   480
      Top             =   8640
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
   Begin MSAdodcLib.Adodc adopuestosol 
      Height          =   330
      Left            =   6000
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "AdoPersonal"
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
   Begin MSAdodcLib.Adodc AdoBeneficiario 
      Height          =   330
      Left            =   4080
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "AdoBeneficiario"
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
   Begin MSAdodcLib.Adodc AdoTransf_Det 
      Height          =   330
      Left            =   1920
      Top             =   8640
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
      Caption         =   "AdoTransf_Det"
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
   Begin MSAdodcLib.Adodc adoac_bienes 
      Height          =   330
      Left            =   7920
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "adoac_bienes"
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
   Begin MSAdodcLib.Adodc AdoAlmDestino 
      Height          =   330
      Left            =   1920
      Top             =   9000
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
      Caption         =   "AdoAlmDestino"
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
   Begin MSAdodcLib.Adodc AdoStock 
      Height          =   330
      Left            =   0
      Top             =   9000
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "AdoStock"
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
   Begin MSAdodcLib.Adodc AdoAlmacen 
      Height          =   330
      Left            =   0
      Top             =   8640
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
      Caption         =   "AdoAlmacen"
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
      Left            =   1920
      Top             =   9480
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
   Begin VB.Label LblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   1200
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label LblUni_descripcion_larga 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      Height          =   225
      Left            =   3360
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADMINISTRACION DE ALMACENES"
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
      Left            =   6150
      TabIndex        =   11
      Top             =   90
      Width           =   5490
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   690
      Left            =   0
      Picture         =   "AlFrmTransfer.frx":F56C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "AlFrmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frenteservicio As String
Dim rsNada As New ADODB.Recordset
Dim rstdetsalalm As New ADODB.Recordset
Dim rstrc_personalSoli, rs_personalUsr As New ADODB.Recordset
Dim rs_Beneficiario, RS_BENEF As New ADODB.Recordset
Dim rs_TipoCambio As New ADODB.Recordset
Dim rstFc_unidad_ejecutora As New ADODB.Recordset
Dim rstacumdet As New ADODB.Recordset
Dim rs_tipoBenef As New ADODB.Recordset
Dim RsAlmacen As New ADODB.Recordset
Dim rs_DestinoAux As New ADODB.Recordset
'Transferencias
Dim rs_Transf As New ADODB.Recordset
Dim rs_Transf_Det As New ADODB.Recordset
Dim rsStock As New ADODB.Recordset
Dim rsAlmDestino As New ADODB.Recordset
Dim rs_Transf_Det2 As New ADODB.Recordset
Dim rsDestinodet2 As New ADODB.Recordset

Dim marca1 As Variant
Dim swgrabar, swnuevo, deta2 As Integer
Dim correlsolic, nroventa, correlv As Integer
Dim correldetalle As Integer
Dim Cobrobs As Double
Dim gestion0 As String
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
Dim queryinicial As String
Dim queryinicial2 As String
'Almacenes
Dim descri_bien As String
Dim Cant_Alm As Integer
Dim CodTrf As Integer
    
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
            txtnosolicitud1.Text = adoTransf.Recordset("codigo_solicitud")
            txtcorrdet.Text = " "
            
            txtsolpeso.Text = 0
        End If
    End If
End Sub

Private Sub adoTransf_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If (Not adoTransf.Recordset.BOF) And (Not adoTransf.Recordset.EOF) Then
   If Not IsNull(adoTransf.Recordset("nro_transfer")) Then
        CodTrf = adoTransf.Recordset("nro_transfer")
        SSTab1.Visible = True
            If (adoTransf.Recordset("estado_registro") = "S") Then
                cmdAprueba.Visible = False
                cmdDesaprueba.Visible = True
                cmdDesaprueba.Enabled = True
                CmdModCabeza.Enabled = False
                CmdDelCabeza.Enabled = False
               ' CmdLista.Enabled = False
            Else
                cmdAprueba.Visible = True
                cmdAprueba.Enabled = True
                cmdDesaprueba.Visible = False
                CmdModCabeza.Enabled = True
                CmdDelCabeza.Enabled = True
            End If
            If adoTransf.Recordset("estado_registro") = "S" Then
                CmdEnviar.Enabled = False
                cmdDesaprueba.Enabled = False
                CmdLista.Enabled = False
                CmdListaMod.Enabled = False
                CmdListaAnula.Enabled = False
            Else
                CmdEnviar.Enabled = True
                CmdLista.Enabled = True
                CmdListaMod.Enabled = True
                CmdListaAnula.Enabled = True
            End If
            
'            If DtcDeudor.Text = "SI" Then
'                DtcDeudor.BackColor = &HFF&
'            Else
'                DtcDeudor.BackColor = &H80000010
'            End If
            'If adoTransf.Recordset("codigo_beneficiario") <> "" And adoTransf.Recordset("codigo_beneficiario") <> "VD" Then
            If adoTransf.Recordset("codigo_beneficiario") <> "" Then
'                Set RS_BENEF = New ADODB.Recordset
'                If RS_BENEF.State = 1 Then RS_BENEF.Close
'                RS_BENEF.Open "select * from FC_beneficiario where codigo_beneficiario = '" & adoTransf.Recordset!codigo_beneficiario & "'  ", db, adOpenKeyset, adLockOptimistic
'                'RS_BENEF.Recordset.Requery
'                If RS_BENEF.RecordCount > 0 Then
''                    If RS_BENEF!emitido = "SI" Then
''                        DtcDeudor.BackColor = &HFF&
''                    Else
''                        DtcDeudor.BackColor = &H80000010
''                    End If
'                End If
                
            End If
        Set rs_Transf_Det = New ADODB.Recordset
        If rs_Transf_Det.State = 1 Then rs_Transf_Det.Close
        If txtnrosol.Text = "" Then txtnrosol.Text = 0
        rs_Transf_Det.Open "select * from AlClDestino_Transf where nro_transfer = " & CodTrf & "  ", db, adOpenKeyset, adLockOptimistic
        'rs_Transf_Det.Open "select * from av_destinoTrf where nro_transfer = " & adoTransf.Recordset!nro_transfer & "  ", db, adOpenKeyset, adLockOptimistic
        'rs_Transf_Det.Open "select * from av_destinoTrf where nro_transfer = '" & txtnrosol.Text & "'  ", db, adOpenKeyset, adLockOptimistic
        Set AdoTransf_Det.Recordset = rs_Transf_Det
        If AdoTransf_Det.Recordset.RecordCount > 0 Then
            deta2 = 1
            Set DtGLista.DataSource = rs_Transf_Det
            Set rs_Transf_Det2 = New ADODB.Recordset
            If rs_Transf_Det2.State = 1 Then rs_Transf_Det2.Close
            rs_Transf_Det2.Open "select * from av_destinoTrf where nro_transfer = '" & CodTrf & "'  ", db, adOpenKeyset, adLockOptimistic
            'Set DtGLista.DataSource = rs_Transf_Det2
        Else
            deta2 = 0
            Set DtGLista.DataSource = rsNada
        End If
        AdoTransf_Det.Recordset.Requery
        FrmDetalle.Caption = "PRODUCTOS DE LA TRANSFERENCIA NRO. " + Str((adoTransf.Recordset("nro_transfer")))
        
'        Else
'            ' por si es nuevo
'            Dtccodbien.Text = " "
'            Dtcdesbien.Text = " "
        End If
    Else
'        CmdModCabeza.Enabled = False
'        CmdDelCabeza.Enabled = False
''        CmdImpCabeza.Enabled = False
'        cmdAprueba.Enabled = False
'        CmdEnviar.Enabled = False
'        Command1.Enabled = False
''        CmdBusCabeza.Enabled = False
        FrmCabecera.Enabled = True
        FrmDetalle.Visible = False
        SSTab1.Visible = False
End If
End Sub

Private Sub Cmd_PersNuevo_Click()
    glPersNew = "P"
    'frmBeneficiario.Show 'vbModal
    frmListaStock.Show vbModal
End Sub

Private Sub CmdAddCabeza_Click()
    FrmCabecera.Enabled = True
    FrmDetalle.Visible = False
    Frmnavega.Enabled = False
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    Frasolic.Enabled = True
    Frame1.Enabled = True
    swgrabar = 1
    'DtgVentas.Visible = False
    'CmdLista.Enabled = False
    'Call cerea
    Dim rstdestino As New ADODB.Recordset
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from AlClDestino_Cabecera where estado_registro <> 'E'", db, adOpenDynamic, adLockOptimistic
    Set adoTransf.Recordset = rstdestino
    adoTransf.Recordset.AddNew
    DTPfechasol.Value = Date
    DTPfechasol.CheckBox = True
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
End Sub

Private Sub CmdAddDetalle_Click()
    FraDetalle.Visible = True
    FraDetalle.Enabled = True
    adoDetalleSolicitud.Recordset.AddNew
    CmdGraDetalle.Enabled = True
    CmdAddDetalle.Enabled = False
    CmdModDetalle.Enabled = False
    CmdSalDetalle.Enabled = False
    CmdCanDetalle.Enabled = True
    txtnosolicitud1.Enabled = False
    txtcorrdet.Enabled = False

    swgrabar = 1
End Sub

Private Sub cmdAprueba_Click()
Dim sinoalmacen As String

If adoTransf.Recordset("CodDestino") = "" Or adoTransf.Recordset("CodDestino2") = "" Or adoTransf.Recordset("Codigo_beneficiario") = "" Or adoTransf.Recordset("Codigo_beneficiario2") = "" Then
   MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
   Exit Sub
Else
  If adoTransf.Recordset("estado_registro") = "N" Then
    sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
    If sino = vbYes Then
        db.Execute "update AlClDestino_Cabecera set AlClDestino_Cabecera.estado_registro = 'S' Where AlClDestino_Cabecera.nro_transfer = " & adoTransf.Recordset("nro_transfer") & "  "
        
'        Dim rstdestino As New ADODB.Recordset
'        Set rstdestino = New ADODB.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from Ao_VENTAS where ges_gestion = '" & adoTransf.Recordset("ges_gestion") & "' and correl_venta = " & adoTransf.Recordset("correl_venta") & " and nro_venta = " & adoTransf.Recordset("nro_venta") & "  ", db, adOpenDynamic, adLockOptimistic
'        If Not rstdestino.BOF Then rstdestino.MoveFirst
'        If Not rstdestino.BOF And Not rstdestino.EOF Then
'            rstdestino("estado_registro") = "S"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
        marca1 = adoTransf.Recordset.BookMark
        adoTransf.Recordset.Requery
'        adoTransf.Refresh
        adoTransf.Recordset.Move marca1 - 1
    End If
    'FALTA CONTABILIZAR !!!!!!!!!!!!!!!!!!!!
  Else
    MsgBox "No se puede APROBAR el registro, verifique los datos y vuelva a intentar ...", , "Atención"
  End If
End If
End Sub

Private Sub CmdBusCabeza_Click()
'JQA
'  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'  Dim ClBuscaSec As ClBuscaSecuencialEnRS
  PosibleApliqueFiltro = False
  Dim rsNada As ADODB.Recordset
  Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = DtgVentas
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = adoTransf.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
End Sub

Private Sub CmdCancelaCobro_Click()
  FrmCobros.Enabled = False
  'swgrabar = 0
  'Call cerea
  swnuevo = 0
  Call OptFilGral1_Click
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    Frmnavega.Enabled = True
    FrmDetalle.Enabled = True
    TxtCobrador.Visible = True
End Sub

Private Sub CmdCancelaDet_Click()
  'TxtNroTrf.Enabled = True
  FrmEdita.Enabled = False
  swgrabar = 0
  'Call cerea
  swnuevo = 0
  'cmdElige.Enabled = False
  marca1 = adoTransf.Recordset.BookMark
  If adoTransf.Recordset("estado_registro") = "S" Then
    Call OptFilGral2_Click
  Else
    Call OptFilGral1_Click
  End If
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    Frmnavega.Enabled = True
    FrmDetalle.Enabled = True
    frmabm.Visible = True
    frmgrabcabeza.Visible = True
    FrmGrabDet.Visible = False
  'AdoTransf_Det.Refresh
  'adoTransf.Recordset.Move marca1 - 1
End Sub

Private Sub cmdDesaprueba_Click()
  sino = MsgBox("Esta seguro de Desaprobar el registro?", vbYesNo, "Confirmando")
  If sino = vbYes Then
    Dim rstdestino As New ADODB.Recordset
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from AlClDestino_Cabecera where nro_transfer = " & adoTransf.Recordset("nro_transfer") & " ", db, adOpenDynamic, adLockOptimistic
    If Not rstdestino.BOF Then rstdestino.MoveFirst
    If Not rstdestino.BOF And Not rstdestino.EOF Then
      rstdestino("estado_registro") = "N"
      rstdestino.Update
    End If
    If rstdestino.State = 1 Then rstdestino.Close
    marca1 = adoTransf.Recordset.BookMark
    Call OptFilGral1_Click
    adoTransf.Recordset.Move marca1 - 1
  End If
End Sub

Private Sub cmdElige_Click()
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
    
End Sub

Private Sub CmdEnviar_Click()
  If adoTransf.Recordset("estado_registro") = "S" Then
    sino = MsgBox("Confirma la entrega de los productos al Cliente ?", vbYesNo, "Confirmando")
    If sino = vbYes Then
        If adoTransf.Recordset("tipo_venta") = "E" Then
            db.Execute "INSERT INTO ao_ventas_cobranzas (nro_venta, correl_venta, ges_gestion, codigo_beneficiario, CI, nombre_cobrador, deuda_cobrada, deuda_cobrada_dol, deuda_dscto, deuda_total, fecha_cobranza, obs_cobranza, nro_cmpbte , Literal, usr_usuario, fecha_registro, hora_registro) VALUES ('" & adoTransf.Recordset!nro_venta & "', '" & adoTransf.Recordset!correl_venta & "', '" & adoTransf.Recordset!ges_gestion & "', '" & adoTransf.Recordset!codigo_beneficiario & "', '" & adoTransf.Recordset!ci & "', '" & Dtcpaternosol + " " + dtcmaternosol + " " + dtcnombresol & "', '" & adoTransf.Recordset!monto_total_Bs & "', '" & adoTransf.Recordset!monto_total_Us & "', '0', '" & adoTransf.Recordset!monto_total_Bs & "', '" & adoTransf.Recordset!fecha_venta & "', 'CANCELADO', '0', '-', '" & GlUsuario & "', '" & Date & "', '" & adoTransf.Recordset!hora_registro & "')"
        End If
        If adoTransf.Recordset("tipo_venta") = "C" Then
            db.Execute "update FC_BENEFICIARIO set emitido = 'SI' where codigo_beneficiario = '" & DtcNIT & "' "
        End If
        Dim rstdestino As New ADODB.Recordset
        Set rstdestino = New ADODB.Recordset
        If rstdestino.State = 1 Then rstdestino.Close
        rstdestino.Open "select * from AlClDestino_Cabecera where ges_gestion = '" & adoTransf.Recordset("ges_gestion") & "' and correl_venta = " & adoTransf.Recordset("correl_venta") & " and nro_venta = " & adoTransf.Recordset("nro_venta") & "  ", db, adOpenDynamic, adLockOptimistic
        If Not rstdestino.BOF Then rstdestino.MoveFirst
        If Not rstdestino.BOF And Not rstdestino.EOF Then
            rstdestino("estado_entregado") = "S"
            rstdestino.Update
        End If
        If rstdestino.State = 1 Then rstdestino.Close
        marca1 = adoTransf.Recordset.BookMark
        adoTransf.Recordset.Requery
        adoTransf.Refresh
        'adoTransf.Recordset.Move marca1 - 1
        'adoTransf.Recordset.MoveLast
        db.Execute "update AlCldetalle set AlCldetalle.stocksalida = av_acumula_venta.cantidad_vendida from AlCldetalle, av_acumula_venta Where AlCldetalle.CodGrupo = av_acumula_venta.CodGrupo And AlCldetalle.cod_MONTADOR = av_acumula_venta.cod_MONTADOR And AlCldetalle.codDetalle = av_acumula_venta.codDetalle"
        db.Execute "update AlCldetalle set StockActual= Stockinicial + stockingreso - StockSalida"
    End If
  Else
    MsgBox "No se puede ENTREGAR!!. Debe Aprobar previamente el registro ...", , "Atención"
  End If
End Sub

Private Sub CmdModDetalle_Click()
  FraDetalle.Visible = True
  FraDetalle.Enabled = True
  txtnosolicitud1.Enabled = False
  txtcorrdet.Enabled = False

  CmdGraDetalle.Enabled = True
  CmdAddDetalle.Enabled = False
  CmdModDetalle.Enabled = False
  CmdSalDetalle.Enabled = False
  CmdCanDetalle.Enabled = True
  swgrabar = 2
End Sub

Private Sub CmdNOunidad_Click()
    swunidad = 0
    Frmunidad.Visible = False
End Sub

Private Sub CmdSalDetalle_Click()
    FraDetalle.Visible = False
    FrmDetalle.Visible = False
    frmabm.Visible = True
    Frmnavega.Enabled = True
    FraDetalle.Enabled = False
    CmdGraDetalle.Enabled = False
    CmdCanDetalle.Enabled = False
End Sub

Private Sub CmdCanDetalle_Click()
'    Adodetallesolicitud.Recordset.CancelUpdate
    adoDetalleSolicitud.Refresh
    CmdGraDetalle.Enabled = False
    CmdAddDetalle.Enabled = True
    CmdModDetalle.Enabled = True
    CmdSalDetalle.Enabled = True
    CmdCanDetalle.Enabled = False
    FraDetalle.Enabled = False
    swgrabar = 0
End Sub

Private Sub CmdGrabaDet_Click()
 If Val(DtcSaldoAct.Text) >= Val(TxtCantidad.Text) Then
      'frmabm.Visible = True
      'frmgrabcabeza.Visible = False
      'TxtNroTrf.Enabled = True
      FrmEdita.Enabled = False
    '  DtGListaN.Enabled = True
      'cmdElige.Enabled = False
    '  Dtccodbien.Visible = False
    '  Dtcdesbien.Visible = False
      'Txtrazon_s.Enabled = False
    If swnuevo = 1 Then
      AdoTransf_Det.Recordset!nro_transfer = adoTransf.Recordset("nro_transfer")
      AdoTransf_Det.Recordset!nro_licitacion = Val(DtcNroCompra.Text)
      AdoTransf_Det.Recordset!codDetalle = Dtccodbien.Text
    End If
      AdoTransf_Det.Recordset!Nro_Lote = DtcNroLote.Text                  'Nro. de Lote de Compra
      AdoTransf_Det.Recordset!cantidad = Val(TxtCantidad)                 'Cantidad a Transferir
      AdoTransf_Det.Recordset!DescDetalle = Dtcdesbien.Text
      AdoTransf_Det.Recordset!usr_usuario = GlUsuario
      AdoTransf_Det.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
      AdoTransf_Det.Recordset!hora_registro = Format(Time, "hh:mm:ss")
      AdoTransf_Det.Recordset.Update
    'db.CommitTrans
    
    Call acumulaMont(adoTransf.Recordset("CodDestino"), adoTransf.Recordset("CodDestino2"), AdoTransf_Det.Recordset("nro_licitacion"), AdoTransf_Det.Recordset("codDetalle"))
    'Call acumulaMont2(adoTransf.Recordset("CodDestino2"), AdoTransf_Det.Recordset("nro_licitacion"), AdoTransf_Det.Recordset("codDetalle"))
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False

    Frmnavega.Enabled = True
    FrmDetalle.Enabled = True
    frmabm.Visible = True
    frmgrabcabeza.Visible = True
    FrmGrabDet.Visible = False
    Call OptFilGral1_Click
    If swnuevo = 1 Then
      'Call abre_ventas_det
      'rs_Transf_Det.Requery
      'AdoTransf_Det.Refresh
      'AdoTransf_Det.Recordset.MoveLast
      
    End If
    swnuevo = 0
  Else
   MsgBox "Saldo insuficiente, Intente nuevamente !..."
  End If

End Sub

Private Sub CmdImpCabeza_Click()
'    Dim IResult As Variant, i%, Y%
'    Dim co As New ADODB.Command
''    Dim rs As New ADODB.Recordset
''    rs.Open "select * from av_ventas_comprobante where ges_gestion='" & Me.adoTransf.Recordset!ges_gestion & "' and " & _
''            "correl_venta=" & Me.adoTransf.Recordset!correl_venta & " and nro_venta=" & Me.adoTransf.Recordset!nro_venta, db, adOpenStatic, adLockReadOnly
''    i = 1
''    y = 1
'    CryV01.ReportFileName = App.Path & "\reportes\ventas\NOTA DE VENTA.rpt"
'    CryV01.WindowShowRefreshBtn = True
'    CryV01.StoredProcParam(0) = Me.adoTransf.Recordset!ges_gestion
'    CryV01.StoredProcParam(1) = Me.adoTransf.Recordset!correl_venta
'    CryV01.StoredProcParam(2) = Me.adoTransf.Recordset!nro_venta
'    IResult = CryV01.PrintReport
'    If IResult <> 0 Then MsgBox CryV01.LastErrorNumber & " : " & CryV01.LastErrorString, vbCritical, "Error de impresión"

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
  marca1 = adoTransf.Recordset.BookMark
  If adoTransf.Recordset!estado_registro = "N" Then
    'If OptFilGral1.Value = True Then Call OptFilGral1_Click
    'If OptFilGral2.Value = True Then Call OptFilGral2_Click
'    adoTransf.Recordset.Move marca1 - 1
    swnuevo = 1
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False
    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    Frmnavega.Enabled = False
    FrmDetalle.Enabled = False
    FrmGrabDet.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = False
     'tipoBenef
   
    AdoTransf_Det.Recordset.AddNew
  Else
    MsgBox "Los productos del registro Aprobado o Entregado, NO pueden ser cambiados !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub CmdListaAnula_Click()
 If adoTransf.Recordset!estado_registro = "N" Then
   sino = MsgBox("Está seguro de ANULAR este registro", vbYesNo + vbQuestion, "Atención ...")
   If sino = vbYes Then
     AdoTransf_Det.Recordset.Delete
     AdoTransf_Det.Recordset.Update
     rs_Transf_Det.Requery
     AdoTransf_Det.Refresh
     'cerea
     AdoTransf_Det.Refresh
   End If
  Else
    MsgBox "Los productos del registro Aprobado , NO pueden ser ANULADOS !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub CmdListaMod_Click()
  If adoTransf.Recordset!estado_registro = "N" Then
    Frmnavega.Enabled = False
    FrmDetalle.Enabled = False
    swgrabar = 0
    swnuevo = 2
    TxtNroTrf.Enabled = False
    marca1 = adoTransf.Recordset.BookMark
    TxtNroTrf.Text = adoTransf.Recordset!nro_transfer  'txtnrosol.Text
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(0) = False
   
    FrmEdita.Visible = True
    FrmEdita.Enabled = True
    FrmGrabDet.Visible = True
    frmabm.Visible = False
    frmgrabcabeza.Visible = False
  Else
    MsgBox "Los productos del registro Aprobado o Entregado, NO pueden ser modificados !! ", vbExclamation, "Atención!"
  End If
End Sub

Private Sub CmdModCabeza_Click()
    FrmCabecera.Enabled = True
    FrmDetalle.Visible = False
     DTPfechasol.SetFocus
    frmabm.Visible = False
    frmabm.Visible = False
    Frmnavega.Enabled = False
    frmgrabcabeza.Visible = True
    Frasolic.Enabled = True
    Frame1.Enabled = True
    
    CmdLista.Enabled = False
    swgrabar = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
End Sub

Private Sub CmdDelCabeza_Click()
  If adoTransf.Recordset("estado_registro") <> "N" Then
    sino = MsgBox("Esta seguro de ANULAR la transferencia registrada ?", vbYesNo, "Confirmando")
    If sino = vbYes Then
        Dim rstdestino As New ADODB.Recordset
        Set rstdestino = New ADODB.Recordset
        If rstdestino.State = 1 Then rstdestino.Close
        rstdestino.Open "select * from AlClDestino_Cabecera where nro_transfer = " & adoTransf.Recordset("nro_transfer") & "   ", db, adOpenDynamic, adLockOptimistic
        If Not rstdestino.BOF Then rstdestino.MoveFirst
        If Not rstdestino.BOF And Not rstdestino.EOF Then
            rstdestino("estado_registro") = "E"
            rstdestino.Update
        End If
        If rstdestino.State = 1 Then rstdestino.Close
        marca1 = adoTransf.Recordset.BookMark
        adoTransf.Recordset.Requery
        adoTransf.Refresh
        adoTransf.Recordset.Move marca1 - 1
    End If
  Else
    MsgBox "NO se puede ANULAR el registro donde el producto ya fue Transferido a Otro Almacen o Anulado ...", , "Atencion"
  End If
End Sub

Private Sub CmdNuevo_Click()
    FrmCabecera.Enabled = True
    FrmDetalle.Visible = False
    Frmnavega.Enabled = False
    frmabm.Visible = False
    frmgrabcabeza.Visible = True
    Frasolic.Enabled = True
    Frame1.Enabled = True
    swgrabar = 1
    'DtgVentas.Visible = False
    'CmdLista.Enabled = False
    'Call cerea
    Dim rstdestino As New ADODB.Recordset
    Set rstdestino = New ADODB.Recordset
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from AlClDestino_Cabecera where estado_registro <> 'E'", db, adOpenDynamic, adLockOptimistic
    Set adoTransf.Recordset = rstdestino
    adoTransf.Recordset.AddNew
    DTPfechasol.Value = Date
    DTPfechasol.CheckBox = True
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False

End Sub

Private Sub CmdSalCabeza_Click()
    sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
    If sino = vbYes Then
'        adoTransf.Recordset.Close
        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rs_Transf_Det.State = 1 Then rs_Transf_Det.Close
'        If rs_Ventas.State = 1 Then rs_Ventas.Close
        Unload Me
    End If
End Sub

Private Sub CmdDetCabeza_Click()
    frmabm.Visible = False
    FrmDetalle.Visible = True
    FraDetalle.Visible = True
    Frmnavega.Enabled = False
    If Not (adoDetalleSolicitud.Recordset.BOF) Then adoDetalleSolicitud.Recordset.MoveFirst
End Sub

Private Sub CmdGraCabeza_Click()
 
    FrmCabecera.Enabled = False
    Call grabar
    frmabm.Visible = True
    frmgrabcabeza.Visible = False
'    adoTransf.Recordset.CancelUpdate
    Frmnavega.Enabled = True
    FrmCabecera.Enabled = False
    Frasolic.Enabled = True
    DtgVentas.Visible = True
    FrmDetalle.Visible = True
     Frame1.Enabled = True
'    adoao_solicitud_detalle.Refresh
    CmdLista.Enabled = True

    'adoTransf.Refresh
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
End Sub

Private Sub CmdCanCabeza_Click()
  'adoTransf.Refresh
  frmabm.Visible = True
'  frmdetalle.Visible = False
  frmgrabcabeza.Visible = False
  marca1 = adoTransf.Recordset.BookMark
  If adoTransf.Recordset("estado_registro") = "S" Then
    Call OptFilGral2_Click
  Else
    Call OptFilGral1_Click
  End If
  Frmnavega.Enabled = True
  FrmCabecera.Enabled = False
  Frasolic.Enabled = True
  FrmDetalle.Visible = True
  DtgVentas.Visible = True
  CmdLista.Enabled = True

  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = True

  'adoTransf.Recordset.Move marca1 - 1
End Sub



Private Sub DtcBenef_Click(Area As Integer)
    DtgNomBenef.BoundText = DtcBenef.BoundText
End Sub

Private Sub DtcCI_Click(Area As Integer)
    DtcPaterno.BoundText = DtcCI.BoundText
    DtcMaterno.BoundText = DtcCI.BoundText
    DtcNombre.BoundText = DtcCI.BoundText
End Sub

Private Sub Dtccibe_Click(Area As Integer)
    Dtcpaternobe.Text = Dtccibe.BoundText
    If Not (IsNull(Dtccibe.Text)) And (Trim(Dtccibe.Text) <> "") Then
        If Not (adopuestobe.Recordset.BOF) Then adopuestobe.Recordset.MoveFirst
        adopuestobe.Recordset.Find "ci = '" & Trim(Dtccibe.Text) & "' ", , adSearchForward
        If Not adoTransf.Recordset.EOF Then
            Dtcmaternobe.Text = IIf(IsNull(adopuestobe.Recordset("materno")) = True, " ", adopuestobe.Recordset("materno"))
            Dtcnombrebe.Text = IIf(IsNull(adopuestobe.Recordset("nombres")) = True, " ", adopuestobe.Recordset("nombres"))
        End If
    End If

End Sub


Private Sub cmdVerifica_existencia_Click()
' verifica existencia  del almacen
Cant_Alm = 0
AlFrmExistencia_Almacen.Show

DE.dbo_albSacaDetalleMaterial Mid(txtCodigo, 3, 12), descri_bien, Cant_Alm
Txtcant_alm = Cant_Alm
If Cant_Alm >= TxtCantPedi Then
        optSi = True
    Else
        optNo = True
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




Private Sub Dtc_Alm_Click(Area As Integer)
    Dtc_AlmDes.BoundText = Dtc_Alm.BoundText
End Sub

Private Sub Dtc_AlmDes_Click(Area As Integer)
    Dtc_Alm.BoundText = Dtc_AlmDes.BoundText
End Sub

Private Sub Dtccodbien_Click(Area As Integer)
    DtcFechaVenc.BoundText = Dtccodbien.BoundText
    DtcCodMontador.BoundText = Dtccodbien.BoundText
    DtcSaldoAct.BoundText = Dtccodbien.BoundText
    Dtcdesbien.BoundText = Dtccodbien.BoundText
    DtcGrupo.BoundText = Dtccodbien.BoundText
'    DtcPrecioCli.BoundText = Dtccodbien.BoundText
    DtcNroLote.BoundText = Dtccodbien.BoundText
'    DtcCorrelAdjDet.BoundText = Dtccodbien.BoundText
    DtcNroCompra.BoundText = Dtccodbien.BoundText
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

Private Sub DtcCodMontador_Click(Area As Integer)
    DtcFechaVenc.BoundText = DtcCodMontador.BoundText
    Dtccodbien.BoundText = DtcCodMontador.BoundText
    DtcSaldoAct.BoundText = DtcCodMontador.BoundText
    Dtcdesbien.BoundText = DtcCodMontador.BoundText
    DtcGrupo.BoundText = DtcCodMontador.BoundText
'    DtcPrecioCli.BoundText = DtcCodMontador.BoundText
    DtcNroLote.BoundText = DtcCodMontador.BoundText
'    DtcCorrelAdjDet.BoundText = DtcCodMontador.BoundText
    DtcNroCompra.BoundText = DtcCodMontador.BoundText
End Sub


Private Sub Dtcdesbien_Click(Area As Integer)
    DtcCodMontador.BoundText = Dtcdesbien.BoundText
    DtcFechaVenc.BoundText = Dtcdesbien.BoundText
    Dtccodbien.BoundText = Dtcdesbien.BoundText
    DtcSaldoAct.BoundText = Dtcdesbien.BoundText
    DtcGrupo.BoundText = Dtcdesbien.BoundText
'    DtcPrecioCli.BoundText = Dtcdesbien.BoundText
    DtcNroLote.BoundText = Dtcdesbien.BoundText
'    DtcCorrelAdjDet.BoundText = Dtcdesbien.BoundText
    DtcNroCompra.BoundText = Dtcdesbien.BoundText
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

Private Sub DtcdesNIT_LostFocus()
    'If AdoBeneficiario.Recordset!emitido = "SI" Then
'    If DtcDeudor.Text = "SI" Then
'        DtcDeudor.BackColor = &HFF&
'    Else
'        DtcDeudor.BackColor = &H80000010
'    End If
    
End Sub

Private Sub DtcMaterno_Click(Area As Integer)
    DtcCI.BoundText = DtcMaterno.BoundText
    DtcPaterno.BoundText = DtcMaterno.BoundText
    DtcNombre.BoundText = DtcMaterno.BoundText
End Sub

Private Sub DtcNombre_Click(Area As Integer)
    DtcCI.BoundText = DtcNombre.BoundText
    DtcMaterno.BoundText = DtcNombre.BoundText
    DtcPaterno.BoundText = DtcNombre.BoundText
End Sub

Private Sub DtcPaterno_Click(Area As Integer)
    DtcCI.BoundText = DtcPaterno.BoundText
    DtcMaterno.BoundText = DtcPaterno.BoundText
    DtcNombre.BoundText = DtcPaterno.BoundText
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

Private Sub Dtcmaternobe_Click(Area As Integer)
    Dtcpaternobe.BoundText = Dtcmaternobe.BoundText
    Dtcnombrebe.BoundText = Dtcmaternobe.BoundText
    Dtccibe.Text = Dtcpaternobe.BoundText
End Sub

Private Sub DtcFechaVenc_Click(Area As Integer)
    DtcCodMontador.BoundText = DtcFechaVenc.BoundText
    DtcGrupo.BoundText = DtcFechaVenc.BoundText
    Dtccodbien.BoundText = DtcFechaVenc.BoundText
    Dtcdesbien.BoundText = DtcFechaVenc.BoundText
    DtcSaldoAct.BoundText = DtcFechaVenc.BoundText
'    DtcPrecioCli.BoundText = DtcFechaVenc.BoundText
    DtcNroLote.BoundText = DtcFechaVenc.BoundText
'    DtcCorrelAdjDet.BoundText = DtcFechaVenc.BoundText
    DtcNroCompra.BoundText = DtcFechaVenc.BoundText
End Sub

Private Sub DtcGrupo_Click(Area As Integer)
    DtcCodMontador.BoundText = DtcGrupo.BoundText
    DtcFechaVenc.BoundText = DtcGrupo.BoundText
    Dtccodbien.BoundText = DtcGrupo.BoundText
    Dtcdesbien.BoundText = DtcGrupo.BoundText
    DtcSaldoAct.BoundText = DtcGrupo.BoundText
'    DtcPrecioCli.BoundText = DtcGrupo.BoundText
    DtcNroLote.BoundText = DtcGrupo.BoundText
'    DtcCorrelAdjDet.BoundText = DtcGrupo.BoundText
    DtcNroCompra.BoundText = DtcGrupo.BoundText
End Sub


Private Sub Dtcnombrebe_Click(Area As Integer)
    Dtcpaternobe.BoundText = Dtcnombrebe.BoundText
    Dtcmaternobe.BoundText = Dtcnombrebe.BoundText
    Dtccibe.Text = Dtcpaternobe.BoundText
End Sub


Private Sub Dtcpaternobe_Click(Area As Integer)
    Dtccibe.Text = Dtcpaternobe.BoundText
    If Not (IsNull(Dtccibe.Text)) And (Trim(Dtccibe.Text) <> "") Then
        If Not (adopuestobe.Recordset.BOF) Then adopuestobe.Recordset.MoveFirst
        adopuestobe.Recordset.Find "ci = '" & Trim(Dtccibe.Text) & "' ", , adSearchForward
        If Not adoTransf.Recordset.EOF Then
            Dtcmaternobe.Text = IIf(IsNull(adopuestobe.Recordset("materno")) = True, " ", adopuestobe.Recordset("materno"))
            Dtcnombrebe.Text = IIf(IsNull(adopuestobe.Recordset("nombres")) = True, " ", adopuestobe.Recordset("nombres"))
        End If
    End If
End Sub


Private Sub DtcNroCompra_Click(Area As Integer)
    DtcCodMontador.BoundText = DtcNroCompra.BoundText
    DtcFechaVenc.BoundText = DtcNroCompra.BoundText
    Dtccodbien.BoundText = DtcNroCompra.BoundText
    DtcSaldoAct.BoundText = DtcNroCompra.BoundText
    DtcGrupo.BoundText = DtcNroCompra.BoundText
'    DtcPrecioCli.BoundText = DtcNroCompra.BoundText
    DtcNroLote.BoundText = DtcNroCompra.BoundText
'    DtcCorrelAdjDet.BoundText = DtcNroCompra.BoundText
    Dtcdesbien.BoundText = DtcNroCompra.BoundText
End Sub

Private Sub DtcNroLote_Click(Area As Integer)
    DtcGrupo.BoundText = DtcNroLote.BoundText
    DtcCodMontador.BoundText = DtcNroLote.BoundText
    DtcFechaVenc.BoundText = DtcNroLote.BoundText
    Dtccodbien.BoundText = DtcNroLote.BoundText
    Dtcdesbien.BoundText = DtcNroLote.BoundText
    DtcSaldoAct.BoundText = DtcNroLote.BoundText
'    DtcPrecioCli.BoundText = DtcNroLote.BoundText
'    DtcCorrelAdjDet.BoundText = DtcNroLote.BoundText
    DtcNroCompra.BoundText = DtcNroLote.BoundText
End Sub

Private Sub DtgNomBenef_Click(Area As Integer)
    DtcBenef.BoundText = DtgNomBenef.BoundText
End Sub

Private Sub DtcSaldoAct_Click(Area As Integer)
    DtcFechaVenc.BoundText = DtcSaldoAct.BoundText
    DtcCodMontador.BoundText = DtcSaldoAct.BoundText
    Dtccodbien.BoundText = DtcSaldoAct.BoundText
    Dtcdesbien.BoundText = DtcSaldoAct.BoundText
    DtcGrupo.BoundText = DtcSaldoAct.BoundText
'    DtcPrecioCli.BoundText = DtcSaldoAct.BoundText
    DtcNroLote.BoundText = DtcSaldoAct.BoundText
'    DtcCorrelAdjDet.BoundText = DtcSaldoAct.BoundText
    DtcNroCompra.BoundText = DtcSaldoAct.BoundText
End Sub

'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub

'Private Sub DTPfechasol_LostFocus()
'    Set rs_TipoCambio = New ADODB.Recordset
'    If rs_TipoCambio.State = 1 Then rs_TipoCambio.Close
'    rs_TipoCambio.Open "select * from ac_tipo_cambio WHERE Fecha_Cambio='" & DTPfechasol & "'  ", db, adOpenKeyset, adLockReadOnly
'    If rs_TipoCambio.RecordCount > 0 Then
'        txtTDC.Text = rs_TipoCambio!cambio_oficial
'    End If
'    adopuestosol.Refresh
'End Sub

Private Sub Form_Load()
  'GlNombFor = "F04"
  LblUsuario.Caption = GlUsuario
  marca1 = 1
  deta2 = 0
   
    Set rstrc_personalSoli = New ADODB.Recordset
    If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
    rstrc_personalSoli.Open "select * from fc_beneficiario WHERE TIPO_beneficiario=1 or TIPO_beneficiario=6 or TIPO_beneficiario=7 ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockReadOnly
    'rstrc_personalSoli.Open "select * from fv_beneficiarioUsr WHERE usuario='" & GlUsuario & "' ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockReadOnly
    Set adopuestosol.Recordset = rstrc_personalSoli
    If adopuestosol.Recordset.RecordCount > 0 Then
    End If
    adopuestosol.Refresh
    
    Set rs_Beneficiario = New ADODB.Recordset
    If rs_Beneficiario.State = 1 Then rs_Beneficiario.Close
    rs_Beneficiario.Open "select * from fc_beneficiario WHERE (TIPO_beneficiario=1 or TIPO_beneficiario=6 or TIPO_beneficiario=7) ORDER BY denominacion_beneficiario ", db, adOpenKeyset, adLockReadOnly
    Set AdoBeneficiario.Recordset = rs_Beneficiario
    AdoBeneficiario.Refresh
    
    Set RsAlmacen = New ADODB.Recordset
    If RsAlmacen.State = 1 Then RsAlmacen.Close
    RsAlmacen.Open "select * from ALCLDestinos", db, adOpenKeyset, adLockReadOnly
    Set AdoAlmacen.Recordset = RsAlmacen
    AdoAlmacen.Refresh
    
    Set rsAlmDestino = New ADODB.Recordset
    If rsAlmDestino.State = 1 Then rsAlmDestino.Close
    rsAlmDestino.Open "select * from ALCLDestinos ", db, adOpenKeyset, adLockReadOnly     'where nro_transfer = '" & TxtNroTrf.Text & "'
    Set AdoAlmDestino.Recordset = rsAlmDestino
    AdoAlmDestino.Refresh

    Set rstac_bienes = New ADODB.Recordset
    If rstac_bienes.State = 1 Then rstac_bienes.Close
    rstac_bienes.Open "select * from ALCLDetalle", db, adOpenKeyset, adLockReadOnly
    Set adoac_bienes.Recordset = rstac_bienes
    adoac_bienes.Refresh
    
    Set rsStock = New ADODB.Recordset
    If rsStock.State = 1 Then rsStock.Close
    rsStock.Open "select * from av_DestinoDet", db, adOpenKeyset, adLockReadOnly
    Set AdoStock.Recordset = rsStock
    AdoStock.Refresh
     
    Call OptFilGral1_Click
    
    FrmEdita.Enabled = False
    FrmCabecera.Enabled = False
    FrmGrabDet.Visible = False
    swnuevo = 0
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
End Sub

Private Sub grabar()
    'db.BeginTrans
    If swgrabar = 1 Then
'      Dim rstdestino As New ADODB.Recordset
'      Set rstdestino = New ADODB.Recordset
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select tipo_tramite, numero_correlativo from fc_correl WHERE tipo_tramite='ventas'", db, adOpenDynamic, adLockOptimistic
'      If rstdestino.RecordCount <> 0 Then
'        adoTransf.Recordset("nro_transfer") = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino!numero_correlativo = (CDbl(rstdestino!numero_correlativo) + 1)
'        rstdestino.Update
'      Else
'        adoTransf.Recordset("nro_transfer") = 1
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
      'adoTransf.Recordset("nro_transfer") = adoTransf.Recordset.RecordCount
      'rstdestino.AddNew
    End If
       'adoTransf.Recordset("nro_transfer") = adoTransf.Recordset.RecordCount + 1
       adoTransf.Recordset("CodDestino") = Dtc_Alm.Text
       adoTransf.Recordset("codigo_beneficiario") = DtcResp.Text
       adoTransf.Recordset("CodDestino2") = Dtc_Alm2
       adoTransf.Recordset("Codigo_beneficiario2") = DtcResp2
       adoTransf.Recordset("fecha_transfer") = DTPfechasol.Value
       adoTransf.Recordset("Observaciones") = TxtConcepto.Text            'E=Efectivo, C=Credito
       adoTransf.Recordset("estado_registro") = "N"
       adoTransf.Recordset("usr_usuario") = GlUsuario
       adoTransf.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
       adoTransf.Recordset("hora_registro") = Format(Time, "hh/mm/ss")
        
    adoTransf.Recordset.Update
 
    'adoTransf.Recordset.Requery
    'If rstdestino.State = 1 Then rstdestino.Close
    'db.CommitTrans
    If adoTransf.Recordset.RecordCount > 0 Then
       marca1 = adoTransf.Recordset.BookMark
       Call OptFilGral1_Click
       'adoTransf.Refresh
       'adoTransf.Recordset.Move marca1 - 1
        If swgrabar = 1 Then
            adoTransf.Refresh
            adoTransf.Recordset.MoveLast
        End If
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


Private Sub ImprAlmacen_Click()
    ALFrmAlmacen.Show
''para asignar del almacen
''De.dbo_alb_graba_CabDetalle adoTransf.Recordset("codigo_solicitud"), rs_Transf_Det!tipo_cambio, rs_Transf_Det!codigo_poa, txtCodigo, txtDesc, TxtCantPedi
'DE.dbo_alb_graba_CabDetalle adoTransf.Recordset("codigo_solicitud"), rs_Transf_Det!tipo_cambio, rs_Transf_Det!codigo_poa, rs_Ventas_lista!ci, rs_Ventas_lista!profesion, rs_Ventas_lista!aplanilla
'CmdEnviar.Enabled = False
'Command1.Enabled = False
'db.Execute " UPDATE AO_SOLICITUD SET APROBADO = 1 ,ESTATUS='S' ,DURACION_ESTIMADA_TIEMPO='ALMACEN'" & _
'    "WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.adoTransf.Recordset!ges_gestion & "' and " & _
'    "(ao_Solicitud.codigo_unidad) = '" & Me.adoTransf.Recordset!codigo_unidad & "' and " & _
'    "(ao_Solicitud.codigo_solicitud) =  " & Me.adoTransf.Recordset!codigo_solicitud & ""
'adoTransf.Refresh
'MsgBox "Solicitud APROBADA / registrada en Entrega Almacen ", vbInformation
''AlmFrmSalidaMaterialF11.Show

' ALMACEN FISICO ***********************************************************
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

Private Sub ImprProd_Click()
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

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
   Set rs_Transf = New ADODB.Recordset
   If rs_Transf.State = 1 Then rs_Transf.Close
   queryinicial = "select * from AlClDestino_Cabecera where estado_registro = 'N' "
   rs_Transf.Open queryinicial, db, adOpenKeyset, adLockOptimistic
   Set adoTransf.Recordset = rs_Transf
   adoTransf.Recordset.Requery
   If adoTransf.Recordset.RecordCount > 0 Then
      adoTransf.Recordset.Move marca1 - 1
      'adoTransf.Recordset.MoveLast
      Set DtgVentas.DataSource = rs_Transf
      CmdNuevo.Visible = False
   Else
      CmdNuevo.Visible = True
   End If
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
  Set rs_Transf = New ADODB.Recordset
  If rs_Transf.State = 1 Then rs_Transf.Close
  queryinicial = "select * from AlClDestino_Cabecera "
   rs_Transf.Open queryinicial, db, adOpenKeyset, adLockOptimistic
   Set adoTransf.Recordset = rs_Transf
   adoTransf.Recordset.Requery
   If adoTransf.Recordset.RecordCount > 0 Then
      adoTransf.Recordset.Move marca1 - 1
      'adoTransf.Recordset.MoveLast
      Set DtgVentas.DataSource = rs_Transf
      CmdNuevo.Visible = False
   Else
      CmdNuevo.Visible = True
   End If
End Sub

'Private Sub Option1_Click()
'    Frame2.Visible = True
'End Sub
'

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
  txtnrosol = " "
  dtccisol.Text = " "
  Dtcpaternosol.Text = " "  'dtccisol.BoundText
'  dtcmaternosol.Text = " "
'  dtcnombresol.Text = " "
  txtCantTotal = "0"
  TxtMontoBs = "0"
  TxtMontoUs = "0"
  TxtConcepto = ""
  DtcNIT = ""
  DtcdesNIT = ""
  txtTDC.Text = GlTipoCambioOficial
  
'  DtCDenominacion_moneda = ""
'  TxtMonto_bolivianos = 0
'  Txtmonto_dolares = 0
'  TxtMonto_bolivianos_contra = 0
'  Txtmonto_dolares_contra = 0
'  DtCOrg_descripcion = ""
'  txtjustifica = ""
'  txtnrosol = ""
'  txtterref = ""
End Sub
Private Sub fbuscaunidad()
  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
  If rstFc_unidad_ejecutora.RecordCount > 0 Then
    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
  Else
    LblUni_descripcion_larga.Caption = ""
  End If
  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
End Sub

''ALB

Sub creaVista()
db.Execute "drop view vwF04"

db.Execute "create view vwF04 as " & _
            "select  ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.tipo_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.grado_instruccion, ao_solicitud_lista.profesion, ao_solicitud_lista.paterno, ao_solicitud_lista.materno, ao_solicitud_lista.nombres, ao_solicitud_lista.telefono, ao_solicitud_lista.razon_s, ao_solicitud.codigo_solicitud, ao_solicitud.justificacion_solicitud, ao_solicitud.duracion_estimada_numero, ao_solicitud.por_tiempo, ao_solicitud.fecha_estimada_inicio, ao_solicitud.caracteristicas, ao_solicitud.duracion_estimada_tiempo, " & _
            "ao_solicitud.tr_adjuntos AS docAdjunta, " & _
            "ao_solicitud.codigo_bien, ac_bienes.bie_descripcion , ao_solicitud.observaciones, fc_unidad_ejecutora.uni_descripcion_larga, ao_solicitud.fecha_solicitud, " & _
            "(rc_personal.paterno) + ' ' + (rc_personal.materno) + ' ' +(rc_personal.nombres) + ' [' + ao_solicitud.ci + ']' AS pmn " & _
            "from ao_solicitud_lista  ,     " & _
                 "ao_solicitud       ,     " & _
                 "fc_unidad_ejecutora,     " & _
                 "rc_personal,             " & _
                 "ac_bienes                " & _
            "where  ao_solicitud_lista.ges_Gestion       = '" & Me.adoTransf.Recordset!ges_gestion & "' and " & _
                    "ao_solicitud_lista.codigo_unidad    = '" & Me.adoTransf.Recordset!codigo_unidad & "' and " & _
                    "ao_solicitud_lista.codigo_solicitud =  " & Me.adoTransf.Recordset!codigo_solicitud & " and " & _
                    "ao_solicitud_lista.ges_Gestion      = ao_solicitud.ges_gestion            and " & _
                    "ao_solicitud_lista.codigo_unidad    = ao_solicitud.codigo_unidad          and " & _
                    "ao_solicitud_lista.codigo_solicitud = ao_solicitud.codigo_solicitud       and " & _
                    "ao_solicitud.codigo_unidad          = fc_unidad_ejecutora.codigo_unidad   and " & _
                    "ao_solicitud.codigo_bien            = ac_bienes.codigo_bien               and " & _
                    "ao_solicitud.ci                     = rc_personal.ci                      " & _
            "GROUP BY ao_solicitud_lista.id_beneficiario, ao_solicitud_lista.doc_identidad, ao_solicitud_lista.tipo_beneficiario, " & _
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
    "ALCLGRUPO.DescGrupo, RC_Personal.paterno, RC_Personal.materno, RC_Personal.nombres, " & _
    "ao_Solicitud.observaciones, ao_Solicitud.caracteristicas, " & _
    "ao_Solicitud.tr_adjuntos, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, " & _
    "ao_Solicitud.duracion_estimada_numero, ao_Solicitud.duracion_estimada_tiempo, " & _
    "ao_solicitud_lista.codDetalle AS ci_material,  ao_solicitud_lista.profesion, ao_solicitud_lista.Aplanilla, " & _
    "ao_solicitud_lista.razon_s, ao_solicitud_lista.Nro_pagos, ao_solicitud_lista.Monto_solicitud_dl, ao_solicitud_lista.AUnidad " & _
"FROM ao_Solicitud, ao_Solicitud_detalle, ALCLGRUPO, RC_Personal, ao_solicitud_lista " & _
"WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.adoTransf.Recordset!ges_gestion & "' and " & _
    "(ao_Solicitud.codigo_unidad) = '" & Me.adoTransf.Recordset!codigo_unidad & "' and " & _
    "(ao_Solicitud.codigo_solicitud) =  " & Me.adoTransf.Recordset!codigo_solicitud & " and " & _
    "ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud AND " & _
    "ao_Solicitud.codigo_unidad = ao_Solicitud_lista.codigo_unidad AND " & _
    "ao_Solicitud.codigo_solicitud = ao_Solicitud_lista.codigo_solicitud AND " & _
    "ao_Solicitud.CodGrupo = ALCLGRUPO.CodGrupo AND " & _
    "ao_Solicitud.ci = RC_Personal.ci"
End Sub

Private Sub acumulaMont(alm, almd, corr, nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
'  LblGestion
'  lblcorrelVenta
'  lblNroVenta
  'rstacumdet.Open "select sum(precio_total_venta) as totbs, sum (precio_total_venta_US) as totdl , sum (cantidad_vendida) as cantot from AlClDestino_Transf where CodDestino = '" & alm & "' and correl_venta = '" & corr & "' and nro_venta = " & nro, db, adOpenKeyset, adLockOptimistic
  rstacumdet.Open "select sum(cantidad) as cantot from av_DestinoTrf where CodDestino = '" & alm & "' and nro_licitacion = '" & corr & "' and CodDetalle = '" & nro & "' ", db, adOpenKeyset, adLockOptimistic
  db.Execute "update AlClDestino_Det set AlClDestino_Det.StockSalida = AlClDestino_Det.StockSalida + " & rstacumdet!cantot & " Where AlClDestino_Det.CodDestino = '" & alm & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
  db.Execute "update AlClDestino_Det set AlClDestino_Det.StockActual = AlClDestino_Det.StockIngreso-AlClDestino_Det.StockSalida Where AlClDestino_Det.CodDestino = '" & alm & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
  'db.Execute "update ao_ventas set ao_ventas.monto_total_Bs = " & rstacumdet!totbs & " , ao_ventas.monto_cobrado = " & rstacumdet!totbs & ", ao_ventas.monto_total_Us = " & rstacumdet!totdl & ", ao_ventas.cantidad_total_vendida = " & rstacumdet!cantot & ", ao_ventas.saldo_p_cobrar = ao_ventas.monto_total_Bs - ao_ventas.deuda_cobrada Where ao_ventas.ges_gestion = '" & ges & "' And ao_ventas.nro_venta = " & nro & " "
  'db.Execute "update AlClDestino_Cabecera set AlClDestino_Cabecera.monto_total_Bs = " & rstacumdet!totbs & " , AlClDestino_Cabecera.monto_total_Us = " & rstacumdet!totdl & ", AlClDestino_Cabecera.cantidad_total_vendida = " & rstacumdet!cantot & ", AlClDestino_Cabecera.saldo_p_cobrar = " & adoTransf.Recordset!monto_total_Bs - adoTransf.Recordset!deuda_cobrada & " Where AlClDestino_Cabecera.ges_gestion = '" & ges & "' And AlClDestino_Cabecera.nro_venta = " & nro & " "
  
  If rstacumdet.State = 1 Then rstacumdet.Close
  
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  rstacumdet.Open "select sum(cantidad) as cantot from av_DestinoTrf where CodDestino2 = '" & almd & "' and nro_licitacion = '" & corr & "' and CodDetalle = '" & nro & "' ", db, adOpenKeyset, adLockOptimistic
  Set rs_DestinoAux = New ADODB.Recordset
  If rs_DestinoAux.State = 1 Then rs_DestinoAux.Close
  rs_DestinoAux.Open "select * from AlClDestino_Det where CodDestino = '" & almd & "' and nro_licitacion = '" & corr & "' and CodDetalle = '" & nro & "' ", db, adOpenKeyset, adLockOptimistic
  If rs_DestinoAux.RecordCount > 0 Then
    db.Execute "update AlClDestino_Det set AlClDestino_Det.StockIngreso = AlClDestino_Det.StockIngreso + " & rstacumdet!cantot & " Where AlClDestino_Det.CodDestino = '" & almd & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
    db.Execute "update AlClDestino_Det set AlClDestino_Det.StockActual = AlClDestino_Det.StockIngreso - AlClDestino_Det.StockSalida Where AlClDestino_Det.CodDestino = '" & almd & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
  Else
    'db.Execute "INSERT INTO AlClDestino_Det (CodDestino, nro_licitacion, CodDetalle, Nro_Lote, fechaVenc, CodGrupo, COD_montador, StockIngreso, StockSalida, StockActual) VALUES ('" & almd & "', " & corr & ", '" & nro & "', '" & nro & "', '" & rstacumdet!Nro_Lote & "', '" & rstacumdet!fechaVenc & "', '" & rstacumdet!CodGrupo & "', '" & rstacumdet!cod_MONTADOR & "', " & rstacumdet!cantidad & ", 0, " & rstacumdet!cantidad & ") "
    
    db.Execute "INSERT INTO AlClDestino_Det(CodDestino, nro_licitacion, CodDetalle, Nro_Lote, fechaVenc, CodGrupo, COD_montador, StockIngreso, StockSalida, StockActual) " & _
    "select '" & almd & "', av_DestinoDet.nro_licitacion, av_DestinoDet.CodDetalle, av_DestinoDet.Nro_Lote, av_DestinoDet.fechaVenc, av_DestinoDet.CodGrupo, av_DestinoDet.COD_montador, " & rstacumdet!cantot & ", 0, " & rstacumdet!cantot & " from av_DestinoDet where av_DestinoDet.CodDestino = '" & alm & "' and av_DestinoDet.nro_licitacion = '" & corr & "' and av_DestinoDet.CodDetalle = '" & nro & "' "
    
'    Set rsDestinodet2 = New ADODB.Recordset
'    If rsDestinodet2.State = 1 Then rsDestinodet2.Close
'    rsDestinodet2.Open "select * from AlClDestino_Det ", db, adOpenKeyset, adLockOptimistic
'    rsDestinodet2.AddNew
'    rsDestinodet2.CodDestino = almd
'    rsDestinodet2.nro_licitacion = corr
'    rsDestinodet2.codDetalle = nro
'    rsDestinodet2.Nro_Lote = AdoTransf_Det.Recordset!Nro_Lote
'    rsDestinodet2.fechaVenc
'    rsDestinodet2.CodGrupo
'    rsDestinodet2.cod_MONTADOR
'    rsDestinodet2.StockIngreso = cantot
'    rsDestinodet2.StockSalida = 0
'    rsDestinodet2.StockActual = cantot
  End If
  If rstacumdet.State = 1 Then rstacumdet.Close
  
End Sub

Private Sub acumulaMont2(alm, corr, nro)
  Set rstacumdet = New ADODB.Recordset
  If rstacumdet.State = 1 Then rstacumdet.Close
  rstacumdet.Open "select sum(cantidad) as cantot from av_DestinoTrf where CodDestino2 = '" & alm & "' and nro_licitacion = '" & corr & "' and CodDetalle = '" & nro & "' ", db, adOpenKeyset, adLockOptimistic
  
  Set rs_DestinoAux = New ADODB.Recordset
  If rs_DestinoAux.State = 1 Then rs_DestinoAux.Close
  rs_DestinoAux.Open "select * from AlClDestino_Det where CodDestino = '" & alm & "' and nro_licitacion = '" & corr & "' and CodDetalle = '" & nro & "' ", db, adOpenKeyset, adLockOptimistic
  If rs_DestinoAux.RecordCount > 0 Then
    db.Execute "update AlClDestino_Det set AlClDestino_Det.StockIngreso = AlClDestino_Det.StockIngreso + " & rstacumdet!cantot & " Where AlClDestino_Det.CodDestino = '" & alm & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
    db.Execute "update AlClDestino_Det set AlClDestino_Det.StockActual = AlClDestino_Det.StockIngreso - AlClDestino_Det.StockSalida Where AlClDestino_Det.CodDestino = '" & alm & "' And AlClDestino_Det.nro_licitacion = " & corr & " and AlClDestino_Det.CodDetalle =  '" & nro & "' "
  Else
    db.Execute "INSERT INTO AlClDestino_Det (CodDestino, nro_licitacion, CodDetalle, Nro_Lote, fechaVenc, CodGrupo, COD_montador, StockIngreso, StockSalida, StockActual) VALUES ('" & alm & "', " & corr & ", '" & nro & "', '" & nro & "', '" & rstacumdet!Nro_Lote & "', '" & rstacumdet!fechaVenc & "', '" & rstacumdet!CodGrupo & "', '" & rstacumdet!cod_MONTADOR & "', " & rstacumdet!cantidad & ", 0, " & rstacumdet!cantidad & ") "
'    db.Execute "INSERT INTO ao_ventas_cobranzas (nro_venta, correl_venta, ges_gestion, codigo_beneficiario, CI, nombre_cobrador, deuda_cobrada, deuda_cobrada_dol, deuda_dscto, deuda_total, fecha_cobranza, obs_cobranza, nro_cmpbte , Literal, usr_usuario, fecha_registro, hora_registro) VALUES ('" & adoTransf.Recordset!nro_venta & "', '" & adoTransf.Recordset!correl_venta & "', '" & adoTransf.Recordset!ges_gestion & "', '" & adoTransf.Recordset!codigo_beneficiario & "', '" & adoTransf.Recordset!ci & "', '" & Dtcpaternosol + " " + dtcmaternosol + " " + dtcnombresol & "', '" & adoTransf.Recordset!monto_total_Bs & "', '" & adoTransf.Recordset!monto_total_Us & "', '0', '" & adoTransf.Recordset!monto_total_Bs & "', '" & adoTransf.Recordset!fecha_venta & "', 'CANCELADO', '0', '-', '" & GlUsuario & "', '" & Date & "', '" & adoTransf.Recordset!hora_registro & "')"
  End If
  If rstacumdet.State = 1 Then rstacumdet.Close
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


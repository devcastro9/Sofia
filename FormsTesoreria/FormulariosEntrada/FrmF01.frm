VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmF01 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos Administrativos - Compras - Solicitudes de cARGO DE cUENTA"
   ClientHeight    =   9480
   ClientLeft      =   30
   ClientTop       =   2070
   ClientWidth     =   14310
   ClipControls    =   0   'False
   HasDC           =   0   'False
   Icon            =   "FrmF01.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   23603.34
   ScaleMode       =   0  'User
   ScaleWidth      =   44924.09
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmGrabaDet 
      BackColor       =   &H00400000&
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
      Height          =   2595
      Left            =   0
      TabIndex        =   110
      Top             =   6000
      Visible         =   0   'False
      Width           =   1020
      Begin VB.CommandButton cmdElige 
         BackColor       =   &H00FFFFC0&
         Caption         =   "New Prod"
         Height          =   720
         Left            =   120
         MaskColor       =   &H80000004&
         Picture         =   "FrmF01.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Registro Nuevo Producto"
         Top             =   1800
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton CmdCancelaDet 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   120
         Picture         =   "FrmF01.frx":6C94
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Cancela Grabación"
         Top             =   960
         Width           =   825
      End
      Begin VB.CommandButton CmdGrabaDet 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Grabar"
         Height          =   735
         Left            =   120
         Picture         =   "FrmF01.frx":6F9E
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Graba Datos del Producto"
         Top             =   120
         Width           =   825
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   3240
      TabIndex        =   17
      Top             =   720
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421504
      ForeColor       =   49152
      TabCaption(0)   =   "CABECERA SOLICITUD COMPRA"
      TabPicture(0)   =   "FrmF01.frx":72A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmgrabcabeza"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmabm"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DETALLE DEL REQUERIMIENTO"
      TabPicture(1)   =   "FrmF01.frx":72C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrmEditaDet"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrmEditaDet 
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
         Height          =   4720
         Left            =   -74900
         TabIndex        =   72
         Top             =   360
         Width           =   8440
         Begin MSDataListLib.DataCombo DtcPrecioUV 
            Bindings        =   "FrmF01.frx":72E0
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   3405
            TabIndex        =   97
            Top             =   1980
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_estimado"
            BoundColumn     =   "CodDetalle"
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
         Begin VB.Frame Frame1 
            BackColor       =   &H80000010&
            Caption         =   "Registro de Datos del Producto"
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
            Height          =   1900
            Left            =   20
            TabIndex        =   74
            Top             =   2805
            Width           =   8380
            Begin VB.TextBox Txtrazon_s 
               CausesValidation=   0   'False
               DataField       =   "razon_s"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   525
               Left            =   1920
               MaxLength       =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   78
               Top             =   1275
               Width           =   6375
            End
            Begin VB.TextBox TxtCantidad 
               Alignment       =   2  'Center
               DataField       =   "APlanilla"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   285
               Left            =   840
               TabIndex        =   77
               Text            =   "0"
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox TxtPrecioU 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               DataField       =   "nro_pagos"
               DataSource      =   "adoao_solicitud_lista"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3795
               TabIndex        =   76
               Text            =   "0"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TxtPrecioC 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               DataField       =   "Precio_compra"
               DataSource      =   "adoao_solicitud_lista"
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
               Height          =   360
               Left            =   6960
               TabIndex        =   75
               Text            =   "0"
               Top             =   800
               Width           =   1335
            End
            Begin MSDataListLib.DataCombo Dtccodbien 
               Bindings        =   "FrmF01.frx":72F9
               DataField       =   "CodDetalle"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   315
               Left            =   840
               TabIndex        =   79
               Top             =   360
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   741
               _Version        =   393216
               BackColor       =   -2147483632
               ForeColor       =   16777152
               ListField       =   "CodDetalle"
               BoundColumn     =   "CodDetalle"
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo Dtcdesbien 
               Bindings        =   "FrmF01.frx":7312
               DataField       =   "CodDetalle"
               DataSource      =   "adoao_solicitud_lista"
               Height          =   315
               Left            =   2595
               TabIndex        =   80
               Top             =   360
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DescDetalle"
               BoundColumn     =   "CodDetalle"
               Text            =   ""
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "Producto:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   375
               Width           =   690
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "Características Comple- mentarias del Producto:"
               ForeColor       =   &H00800000&
               Height          =   390
               Left            =   120
               TabIndex        =   84
               Top             =   1260
               Width           =   2175
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label27 
               BackColor       =   &H80000010&
               Caption         =   "Cantidad:"
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   120
               TabIndex        =   83
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label Label24 
               BackColor       =   &H80000010&
               Caption         =   "Precio Referen. Actual:"
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   2115
               TabIndex        =   82
               Top             =   840
               Width           =   2085
            End
            Begin VB.Label Label23 
               BackColor       =   &H80000010&
               Caption         =   "Precio Compra Actual:"
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   5325
               TabIndex        =   81
               Top             =   840
               Width           =   1725
            End
         End
         Begin MSDataListLib.DataCombo DtcCodUniv 
            Bindings        =   "FrmF01.frx":732B
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   3400
            TabIndex        =   73
            Top             =   1980
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "cod_univ"
            BoundColumn     =   "CodDetalle"
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
         Begin MSDataListLib.DataCombo DtcSubgrupoDes 
            Bindings        =   "FrmF01.frx":7344
            DataField       =   "COD_MONTADOR"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   2160
            TabIndex        =   86
            Top             =   1365
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "descripcion"
            BoundColumn     =   "COD_MONTADOR"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcGrupo 
            Bindings        =   "FrmF01.frx":735E
            DataField       =   "CodGrupo"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   1800
            TabIndex        =   87
            Top             =   960
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "DescGrupo"
            BoundColumn     =   "CodGrupo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcSubgrupo 
            Bindings        =   "FrmF01.frx":7375
            DataField       =   "COD_MONTADOR"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   960
            TabIndex        =   88
            Top             =   1365
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "COD_MONTADOR"
            BoundColumn     =   "COD_MONTADOR"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCodGrupo 
            Bindings        =   "FrmF01.frx":738F
            DataField       =   "CodGrupo"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   960
            TabIndex        =   89
            Top             =   960
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "CodGrupo"
            BoundColumn     =   "CodGrupo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcCodGrupoP 
            Bindings        =   "FrmF01.frx":73A6
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   7320
            TabIndex        =   90
            Top             =   960
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483633
            ListField       =   "CodGrupo"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcSubgrupoP 
            Bindings        =   "FrmF01.frx":73BF
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   7320
            TabIndex        =   91
            Top             =   1320
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483633
            ListField       =   "COD_MONTADOR"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcPrecioC 
            Bindings        =   "FrmF01.frx":73D8
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   7005
            TabIndex        =   92
            Top             =   1980
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_Compra"
            BoundColumn     =   "CodDetalle"
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
         Begin MSDataListLib.DataCombo DtcPrecioU 
            Bindings        =   "FrmF01.frx":73F1
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   5235
            TabIndex        =   93
            Top             =   1980
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Precio_Salon"
            BoundColumn     =   "CodDetalle"
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
         Begin MSDataListLib.DataCombo DtcCodAnt 
            Bindings        =   "FrmF01.frx":740A
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   1725
            TabIndex        =   94
            Top             =   1980
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Cod_Ant"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtcdesAnt 
            Bindings        =   "FrmF01.frx":7423
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   2160
            TabIndex        =   95
            Top             =   2400
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Nombre_Anterior"
            BoundColumn     =   "CodDetalle"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo Dtc_UniMed 
            Bindings        =   "FrmF01.frx":743C
            DataField       =   "CodDetalle"
            DataSource      =   "adoao_solicitud_lista"
            Height          =   315
            Left            =   120
            TabIndex        =   96
            Top             =   1980
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   -2147483632
            ForeColor       =   16777152
            ListField       =   "Unidad"
            BoundColumn     =   "CodDetalle"
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
         Begin VB.Label lbltipoVenta 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRANDO PRODUCTOS PARA PEDIDO AL PROVEEDOR  ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   270
            Left            =   960
            TabIndex        =   109
            Top             =   240
            Width           =   6975
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "SubGrupo:"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   1425
            Width           =   765
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Grupo :"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   120
            TabIndex        =   104
            Top             =   1005
            Width           =   525
         End
         Begin VB.Label Label18 
            BackColor       =   &H80000010&
            Caption         =   "Nombre Anterior Producto:"
            ForeColor       =   &H00400000&
            Height          =   240
            Left            =   120
            TabIndex        =   103
            Top             =   2400
            Width           =   1890
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Precio Referen."
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   5220
            TabIndex        =   102
            Top             =   1740
            Width           =   1230
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Código Alternativo"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   3405
            TabIndex        =   101
            Top             =   1740
            Width           =   1305
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Precio Compra"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   7035
            TabIndex        =   100
            Top             =   1740
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Unidad Medida"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   120
            TabIndex        =   99
            Top             =   1740
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Código Anterior"
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   1740
            TabIndex        =   98
            Top             =   1740
            Width           =   1200
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmabm 
         BackColor       =   &H00404080&
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
         TabIndex        =   58
         Top             =   360
         Width           =   8430
         Begin VB.CommandButton Command2 
            Caption         =   "Imprimir"
            Height          =   720
            Left            =   6495
            Picture         =   "FrmF01.frx":7455
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Imprime Solicitud"
            Top             =   120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdAprueba 
            Caption         =   "Aprobar"
            Height          =   720
            Left            =   2880
            Picture         =   "FrmF01.frx":8BD7
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Aprueba Solicitud"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdEnviar 
            Caption         =   "Enviar"
            Height          =   720
            Left            =   3765
            Picture         =   "FrmF01.frx":98A1
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Envia Solicitud al Proveedor"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdAddCabeza 
            Caption         =   "Adicionar"
            Height          =   720
            Left            =   225
            Picture         =   "FrmF01.frx":A56B
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Adiciona Nueva Solicitud"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdModCabeza 
            Caption         =   "Modificar"
            Height          =   720
            Left            =   1110
            Picture         =   "FrmF01.frx":11059
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Modifica Solicitud Existente"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdDelCabeza 
            Caption         =   "Anular"
            Height          =   720
            Left            =   1995
            Picture         =   "FrmF01.frx":11923
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Anula Solicitud Existente"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdImpCabeza 
            Caption         =   "Imprimir"
            Height          =   720
            Left            =   4650
            Picture         =   "FrmF01.frx":125ED
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Imprime Solicitud"
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdSalCabeza 
            Caption         =   "Salir"
            Height          =   720
            Left            =   7440
            Picture         =   "FrmF01.frx":13D6F
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Salir de ""Solicitudes de Compra"""
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton cmdDesaprueba 
            Caption         =   "Desapro."
            Height          =   720
            Left            =   2880
            Picture         =   "FrmF01.frx":13F79
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Almacen"
            Height          =   720
            Left            =   3765
            Picture         =   "FrmF01.frx":14183
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdBusCabeza 
            Caption         =   "Buscar"
            Height          =   720
            Left            =   5580
            Picture         =   "FrmF01.frx":14E4D
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Busca una Solicitud"
            Top             =   120
            Width           =   765
         End
         Begin Crystal.CrystalReport CryF01 
            Left            =   6525
            Top             =   285
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
      Begin VB.Frame frmgrabcabeza 
         BackColor       =   &H00404080&
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
         TabIndex        =   69
         Top             =   360
         Visible         =   0   'False
         Width           =   8430
         Begin VB.CommandButton CmdCanCabeza 
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4110
            Picture         =   "FrmF01.frx":15717
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdGraCabeza 
            Caption         =   "Grabar"
            Height          =   675
            Left            =   3345
            Picture         =   "FrmF01.frx":15921
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   3870
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   8430
         Begin VB.Frame Frame2 
            BackColor       =   &H80000010&
            Caption         =   " Datos Complementarios"
            ForeColor       =   &H00000080&
            Height          =   1760
            Left            =   -60
            TabIndex        =   19
            Top             =   2040
            Width           =   8595
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0.000%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   6480
               TabIndex        =   21
               Text            =   "0"
               Top             =   1410
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.CheckBox ChkTdr 
               BackColor       =   &H80000010&
               Caption         =   "Se Adjuntan Especificaciones o Doc.s ?"
               ForeColor       =   &H00000040&
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   1440
               Value           =   1  'Checked
               Width           =   3495
            End
            Begin VB.TextBox txtterref 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   3360
               MaxLength       =   1
               TabIndex        =   26
               Top             =   1380
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.TextBox Txt_porcentaje 
               Alignment       =   2  'Center
               DataField       =   "por_tiempo"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   7335
               TabIndex        =   25
               Top             =   1410
               Width           =   780
            End
            Begin VB.TextBox txtjustifica 
               DataField       =   "justificacion_solicitud"
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   1560
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Top             =   1080
               Visible         =   0   'False
               Width           =   6855
            End
            Begin VB.TextBox Txtcaracteristicas 
               DataField       =   "caracteristicas"
               DataSource      =   "adosolicitud"
               Height          =   450
               Left            =   1560
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   600
               Width           =   6855
            End
            Begin VB.TextBox Txtobservaciones 
               DataField       =   "observaciones"
               DataSource      =   "adosolicitud"
               Height          =   285
               Left            =   1560
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   1080
               Width           =   6855
            End
            Begin MSDataListLib.DataCombo DtcPOADes 
               Bindings        =   "FrmF01.frx":15B2B
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2760
               TabIndex        =   20
               Top             =   240
               Width           =   5640
               _ExtentX        =   9948
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "descripcion_poa"
               BoundColumn     =   "codigo_poa"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtcdesbien1 
               Bindings        =   "FrmF01.frx":15B40
               DataField       =   "codGrupo"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   2400
               TabIndex        =   28
               Top             =   600
               Visible         =   0   'False
               Width           =   5880
               _ExtentX        =   10372
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "DescGrupo"
               BoundColumn     =   "CodGrupo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtccodbien1 
               Bindings        =   "FrmF01.frx":15B5B
               DataField       =   "codGrupo"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   1560
               TabIndex        =   29
               Top             =   600
               Visible         =   0   'False
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               BackColor       =   14737632
               ListField       =   "CodGrupo"
               BoundColumn     =   "codGrupo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcPOA 
               Bindings        =   "FrmF01.frx":15B76
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   1560
               TabIndex        =   30
               Top             =   240
               Width           =   1470
               _ExtentX        =   2593
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   -2147483624
               ListField       =   "codigo_poa"
               BoundColumn     =   "codigo_poa"
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
            Begin MSDataListLib.DataCombo DtcMarca 
               Bindings        =   "FrmF01.frx":15B8B
               DataField       =   "codigo_poa"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   6360
               TabIndex        =   106
               Top             =   480
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               BackColor       =   12632256
               ForeColor       =   12648447
               ListField       =   "ent_codigo"
               BoundColumn     =   "codigo_poa"
               Text            =   ""
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000010&
               Caption         =   "Actividad (POA) :"
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   240
               TabIndex        =   36
               Top             =   300
               Width           =   1275
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   "S=Si  /  N=No"
               ForeColor       =   &H00800000&
               Height          =   165
               Left            =   3720
               TabIndex        =   35
               Top             =   1470
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000010&
               Caption         =   "Porcentaje de Descuento:"
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   5415
               TabIndex        =   34
               Top             =   1440
               Width           =   1860
            End
            Begin VB.Label Label16 
               BackColor       =   &H80000010&
               Caption         =   "Observaciones :"
               ForeColor       =   &H00000040&
               Height          =   285
               Left            =   240
               TabIndex        =   33
               Top             =   1140
               Width           =   1275
            End
            Begin VB.Label Label26 
               BackColor       =   &H80000010&
               Caption         =   "Descripción o Caract. Grales.:"
               ForeColor       =   &H00000040&
               Height          =   405
               Left            =   240
               TabIndex        =   32
               Top             =   600
               Width           =   1260
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   " % "
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
               Height          =   165
               Left            =   8160
               TabIndex        =   31
               Top             =   1440
               Width           =   225
            End
         End
         Begin VB.Frame Frame2A 
            BackColor       =   &H80000010&
            Caption         =   $"FrmF01.frx":15BA0
            ForeColor       =   &H00000040&
            Height          =   960
            Left            =   -60
            TabIndex        =   37
            Top             =   1320
            Width           =   8595
            Begin MSDataListLib.DataCombo Dtcpaternobe 
               Bindings        =   "FrmF01.frx":15C2A
               DataField       =   "ci_aprueba"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   120
               TabIndex        =   38
               Top             =   285
               Width           =   6075
               _ExtentX        =   10716
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "denominacion_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo Dtccibe 
               Bindings        =   "FrmF01.frx":15C44
               DataField       =   "ci_aprueba"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   6180
               TabIndex        =   39
               Top             =   285
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "codigo_beneficiario"
               BoundColumn     =   "codigo_beneficiario"
               Text            =   ""
            End
         End
         Begin VB.Frame Frasolic 
            BackColor       =   &H80000010&
            Caption         =   " Apellidos y Nombres del Responsable de la Solicitud ---------------------------------------------------------- Doc. de Identidad"
            ForeColor       =   &H00000040&
            Height          =   825
            Left            =   -60
            TabIndex        =   47
            Top             =   660
            Width           =   8595
            Begin MSDataListLib.DataCombo dtccisol 
               Bindings        =   "FrmF01.frx":15C5E
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   6510
               TabIndex        =   48
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "CI"
               BoundColumn     =   "ci"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo dtcnombresol 
               Bindings        =   "FrmF01.frx":15C79
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   330
               Left            =   4215
               TabIndex        =   49
               Top             =   240
               Width           =   2560
               _ExtentX        =   4524
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   -2147483624
               ListField       =   "Nombres"
               BoundColumn     =   "CI"
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo dtcmaternosol 
               Bindings        =   "FrmF01.frx":15C94
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   330
               Left            =   2385
               TabIndex        =   50
               Top             =   240
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   741
               _Version        =   393216
               Enabled         =   0   'False
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   -2147483624
               ListField       =   "materno"
               BoundColumn     =   "CI"
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSDataListLib.DataCombo Dtcpaternosol 
               Bindings        =   "FrmF01.frx":15CAF
               DataField       =   "ci"
               DataSource      =   "adosolicitud"
               Height          =   315
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "paterno"
               BoundColumn     =   "ci"
               Text            =   ""
            End
         End
         Begin VB.Frame FrmApertura 
            Caption         =   "APERTURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   705
            Left            =   360
            TabIndex        =   41
            Top             =   2400
            Visible         =   0   'False
            Width           =   8055
            Begin VB.ComboBox cmbSubCta2 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "FrmF01.frx":15CCA
               Left            =   6240
               List            =   "FrmF01.frx":15CD4
               TabIndex        =   42
               Top             =   255
               Visible         =   0   'False
               Width           =   1695
            End
            Begin MSDataListLib.DataCombo DtCvalor1 
               Bindings        =   "FrmF01.frx":15CEA
               Height          =   315
               Left            =   1680
               TabIndex        =   43
               Top             =   255
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "valor1"
               BoundColumn     =   "TIPO"
               Text            =   ""
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Trámite:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   46
               Top             =   285
               Width           =   1395
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Cargo de Cuenta:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4200
               TabIndex        =   45
               Top             =   285
               Width           =   2055
            End
            Begin VB.Label Lbltipo_bien_Cta_doc1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   6105
               TabIndex        =   44
               Top             =   345
               Width           =   1875
            End
         End
         Begin MSComCtl2.DTPicker DTPfechasol 
            DataField       =   "fecha_solicitud"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   6765
            TabIndex        =   40
            Top             =   150
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   89915393
            CurrentDate     =   36464
         End
         Begin MSDataListLib.DataCombo DtcUnidad 
            Bindings        =   "FrmF01.frx":15D09
            DataField       =   "codigo_unidad"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   3240
            TabIndex        =   52
            Top             =   150
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   741
            _Version        =   393216
            ListField       =   "codigo_unidad"
            BoundColumn     =   "codigo_unidad"
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
         Begin MSDataListLib.DataCombo DtcUnidadDes 
            Bindings        =   "FrmF01.frx":15D21
            DataField       =   "codigo_unidad"
            DataSource      =   "adosolicitud"
            Height          =   315
            Left            =   3240
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Uni_descripcion_larga"
            BoundColumn     =   "codigo_unidad"
            Text            =   ""
         End
         Begin VB.Label txtnrosol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label29"
            DataField       =   "codigo_solicitud"
            DataSource      =   "adosolicitud"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   1080
            TabIndex        =   108
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Lbltipo_bien_Cta_doc 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   2880
            TabIndex        =   57
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000010&
            Caption         =   "Código Unidad:"
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   2640
            TabIndex        =   56
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000010&
            Caption         =   "Fecha Solicitud:"
            ForeColor       =   &H00000040&
            Height          =   405
            Left            =   5910
            TabIndex        =   55
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000010&
            Caption         =   "No. de Solicitud:"
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
            Height          =   405
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   1005
         End
      End
   End
   Begin VB.Frame Frmnavega 
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
      Height          =   5130
      Left            =   0
      TabIndex        =   5
      Top             =   715
      Width           =   3300
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H80000010&
         Caption         =   "Sin Enviar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H80000010&
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Top             =   120
         Width           =   795
      End
      Begin MSAdodcLib.Adodc adosolicitud 
         Height          =   330
         Left            =   60
         Top             =   4770
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
         BackColor       =   12648384
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
         Caption         =   "Cabecera Solicitud Compra"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmF01.frx":15D39
         Height          =   4275
         Left            =   60
         TabIndex        =   2
         Top             =   480
         Width           =   3170
         _ExtentX        =   5583
         _ExtentY        =   7541
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   32768
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
         Caption         =   "CABECERA SOLICITUD COMPRA"
         ColumnCount     =   42
         BeginProperty Column00 
            DataField       =   "codigo_unidad"
            Caption         =   "UNIDAD"
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
            DataField       =   "codigo_solicitud"
            Caption         =   "Nro.Sol."
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
            DataField       =   "Ges_Gestion"
            Caption         =   "Ges_Gestion"
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
            DataField       =   "estado_aprobado"
            Caption         =   "APR"
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
            DataField       =   "estado_enviado"
            Caption         =   "ENV"
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
            DataField       =   "aprobado"
            Caption         =   "Aprobado"
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
            DataField       =   "tipo_formulario"
            Caption         =   "TIPO"
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
            DataField       =   "justificacion_solicitud"
            Caption         =   "justificacion_solicitud"
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
            DataField       =   "CI"
            Caption         =   "CI"
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
            DataField       =   "Codigo_puesto"
            Caption         =   "Codigo_puesto"
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
            DataField       =   "CI_aprueba"
            Caption         =   "CI_aprueba"
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
            DataField       =   "Fecha_recepción"
            Caption         =   "Fecha_recepción"
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
         BeginProperty Column12 
            DataField       =   "fecha_solicitud"
            Caption         =   "fecha_solicitud"
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
         BeginProperty Column13 
            DataField       =   "codigo_poa"
            Caption         =   "codigo_poa"
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
         BeginProperty Column14 
            DataField       =   "tipo_moneda"
            Caption         =   "tipo_moneda"
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
         BeginProperty Column15 
            DataField       =   "monto_bolivianos"
            Caption         =   "monto_bolivianos"
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
         BeginProperty Column16 
            DataField       =   "monto_dolares"
            Caption         =   "monto_dolares"
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
         BeginProperty Column17 
            DataField       =   "Tipo_cambio"
            Caption         =   "Tipo_cambio"
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
         BeginProperty Column18 
            DataField       =   "monto_bolivianos_contra"
            Caption         =   "monto_bolivianos_contra"
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
         BeginProperty Column19 
            DataField       =   "monto_dolares_contra"
            Caption         =   "monto_dolares_contra"
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
         BeginProperty Column20 
            DataField       =   "org_codigo_contra"
            Caption         =   "org_codigo_contra"
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
         BeginProperty Column21 
            DataField       =   "Uni_codigo"
            Caption         =   "Uni_codigo"
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
         BeginProperty Column22 
            DataField       =   "consultor_empresa"
            Caption         =   "consultor_empresa"
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
         BeginProperty Column23 
            DataField       =   "nacional_extranjero"
            Caption         =   "nacional_extranjero"
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
         BeginProperty Column24 
            DataField       =   "funcion_actividad"
            Caption         =   "funcion_actividad"
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
         BeginProperty Column25 
            DataField       =   "duracion_estimada_numero"
            Caption         =   "duracion_estimada_numero"
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
         BeginProperty Column26 
            DataField       =   "duracion_estimada_tiempo"
            Caption         =   "duracion_estimada_tiempo"
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
         BeginProperty Column27 
            DataField       =   "impuestos"
            Caption         =   "impuestos"
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
         BeginProperty Column28 
            DataField       =   "por_tiempo"
            Caption         =   "por_tiempo"
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
         BeginProperty Column29 
            DataField       =   "fecha_estimada_inicio"
            Caption         =   "fecha_estimada_inicio"
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
         BeginProperty Column30 
            DataField       =   "tr_adjuntos"
            Caption         =   "tr_adjuntos"
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
         BeginProperty Column31 
            DataField       =   "observaciones"
            Caption         =   "observaciones"
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
         BeginProperty Column32 
            DataField       =   "codigo_bien"
            Caption         =   "codigo_bien"
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
         BeginProperty Column33 
            DataField       =   "caracteristicas"
            Caption         =   "caracteristicas"
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
         BeginProperty Column34 
            DataField       =   "usr_usuario"
            Caption         =   "usr_usuario"
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
         BeginProperty Column35 
            DataField       =   "pas_viat"
            Caption         =   "pas_viat"
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
         BeginProperty Column36 
            DataField       =   "fecha_registro"
            Caption         =   "fecha_registro"
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
         BeginProperty Column37 
            DataField       =   "hora_registro"
            Caption         =   "hora_registro"
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
         BeginProperty Column38 
            DataField       =   "usuario_aprueba"
            Caption         =   "usuario_aprueba"
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
         BeginProperty Column39 
            DataField       =   "fecha_aprueba"
            Caption         =   "fecha_aprueba"
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
         BeginProperty Column40 
            DataField       =   "hora_aprueba"
            Caption         =   "hora_aprueba"
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
         BeginProperty Column41 
            DataField       =   "Lista_adjunta"
            Caption         =   "Lista_adjunta"
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
            EndProperty
            BeginProperty Column21 
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
            BeginProperty Column26 
            EndProperty
            BeginProperty Column27 
            EndProperty
            BeginProperty Column28 
            EndProperty
            BeginProperty Column29 
            EndProperty
            BeginProperty Column30 
            EndProperty
            BeginProperty Column31 
            EndProperty
            BeginProperty Column32 
            EndProperty
            BeginProperty Column33 
            EndProperty
            BeginProperty Column34 
            EndProperty
            BeginProperty Column35 
            EndProperty
            BeginProperty Column36 
            EndProperty
            BeginProperty Column37 
            EndProperty
            BeginProperty Column38 
            EndProperty
            BeginProperty Column39 
            EndProperty
            BeginProperty Column40 
            EndProperty
            BeginProperty Column41 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      Picture         =   "FrmF01.frx":15D54
      ScaleHeight     =   675
      ScaleWidth      =   11835
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      Begin VB.Label LblUni_descripcion_larga 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   5160
      End
      Begin VB.Label label7 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblUni_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUDES DE CARGO DE CUENTA"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   405
         Left            =   5655
         TabIndex        =   9
         Top             =   120
         Width           =   5910
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin MSAdodcLib.Adodc Adocc_parametros 
      Height          =   330
      Left            =   0
      Top             =   9120
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
      Caption         =   "Adocc_parametros"
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
   Begin MSAdodcLib.Adodc adopuestosol 
      Height          =   330
      Left            =   2160
      Top             =   9120
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
      Caption         =   "adopuestosol"
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
   Begin MSAdodcLib.Adodc adopuestobe 
      Height          =   330
      Left            =   2160
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
      Caption         =   "adopuestobe"
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
   Begin MSAdodcLib.Adodc AdoUnidad 
      Height          =   330
      Left            =   4320
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoUnidad"
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
   Begin MSAdodcLib.Adodc AdoPOA 
      Height          =   330
      Left            =   6360
      Top             =   8760
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "AdoPOA"
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
   Begin MSAdodcLib.Adodc adoao_solicitud_detalle 
      Height          =   330
      Left            =   8400
      Top             =   8760
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
      Caption         =   "Navegar"
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
   Begin MSAdodcLib.Adodc ado_bienes 
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
      Caption         =   "ado_bienes"
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2880
      Left            =   0
      TabIndex        =   13
      Top             =   5880
      Width           =   11880
      Begin VB.CommandButton CmdListaAnula 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Anular -->"
         Height          =   705
         Left            =   120
         Picture         =   "FrmF01.frx":1A844
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   1920
         Width           =   865
      End
      Begin VB.CommandButton CmdListaMod 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modificar->"
         Height          =   705
         Left            =   120
         Picture         =   "FrmF01.frx":1AC86
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   1080
         Width           =   865
      End
      Begin VB.CommandButton CmdLista 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Adicionar->"
         Height          =   705
         Left            =   120
         Picture         =   "FrmF01.frx":1B0C8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Adiciona Producto"
         Top             =   240
         Width           =   865
      End
      Begin MSDataGridLib.DataGrid DtGLista 
         Bindings        =   "FrmF01.frx":1B50A
         Height          =   2295
         Left            =   1080
         TabIndex        =   4
         Top             =   180
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777152
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         Caption         =   "DETALLE DE REQUERIMIENTO"
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CodGrupo"
            Caption         =   "Grupo"
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
            DataField       =   "cod_montador"
            Caption         =   "Sub-Grupo"
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
            DataField       =   "CodDetalle"
            Caption         =   "BBySS"
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
            DataField       =   "profesion"
            Caption         =   "Denominación"
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
            DataField       =   "cantidad"
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
         BeginProperty Column05 
            DataField       =   "precio_venta"
            Caption         =   "Prec.Bruto"
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
            DataField       =   "total_venta"
            Caption         =   "Total Bruto"
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
            DataField       =   "precio_compra"
            Caption         =   "Prec.Neto"
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
            DataField       =   "total_compra"
            Caption         =   "Total Neto"
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
            DataField       =   "razon_s"
            Caption         =   "Caracteristicas del Bien"
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
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DtGao_solicitud_detalle 
         Bindings        =   "FrmF01.frx":1B52E
         Height          =   1275
         Left            =   1320
         TabIndex        =   14
         Top             =   300
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2249
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "codigo_solicitud"
            Caption         =   "Solicitud"
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
            DataField       =   "codigo_detalle"
            Caption         =   "Nro.Det."
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
            DataField       =   "codigo_unidad"
            Caption         =   "Unidad"
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
            DataField       =   "codigo_poa"
            Caption         =   "Frente Servic."
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
            DataField       =   "monto_bolivianos"
            Caption         =   "Monto_Bs. (B)"
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
            DataField       =   "monto_dolares"
            Caption         =   "Monto_$US (B)"
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
            DataField       =   "Tipo_cambio"
            Caption         =   "TDC"
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
            DataField       =   "monto_bolivianos_contra"
            Caption         =   "Monto_Bs. (I)"
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
            DataField       =   "monto_dolares_contra"
            Caption         =   "Monto_$US (I)"
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
            DataField       =   "tipo_moneda"
            Caption         =   "Moneda"
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
            DataField       =   "org_codigo_ext"
            Caption         =   "Fin_principal"
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
            DataField       =   "org_codigo_contra"
            Caption         =   "Fin_Impuesto"
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
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoao_solicitud_lista 
         Height          =   330
         Left            =   1080
         Top             =   2520
         Width           =   10755
         _ExtentX        =   18971
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
         BackColor       =   16777152
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
         Caption         =   "Detalle de Productos"
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
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   5
         X1              =   0
         X2              =   11880
         Y1              =   60
         Y2              =   60
      End
   End
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   330
      Left            =   4320
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
   Begin MSAdodcLib.Adodc AdoMontador 
      Height          =   330
      Left            =   6360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Caption         =   "AdoMontador"
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
Attribute VB_Name = "FrmF01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim rscc_parametros As New adodb.Recordset
'Dim rstdestino As New adodb.Recordset
'Dim conv1, conv2, conv_nal, CONVE As String
'Dim cta1, txtGes_gestion As String
'Dim cat_nal, CATEG, SOLISTA As String
'Dim parametro, parametro2 As String
'Dim GCODIGO_PAGO As String
''
'Dim rstdetsalalm As New adodb.Recordset
'Dim rstAo_solicitud As New adodb.Recordset
'Dim rstao_solicitud_detalle As New adodb.Recordset
'Dim rstpoa As New adodb.Recordset
'Dim rstpoaAux As New adodb.Recordset
'Dim rstrc_personalSoli As New adodb.Recordset
'Dim rstrc_personalCargo As New adodb.Recordset
'Dim rstfc_partida_gasto As New adodb.Recordset
'Dim rstFc_unidad_ejecutora As New adodb.Recordset
'Dim rstac_bienes As New adodb.Recordset
'Dim rstfc_relacionador_poa_ppto As New adodb.Recordset
'Dim rstOrganismo_finanExt As New adodb.Recordset
'Dim rstao_solicitud_lista As New adodb.Recordset
'Dim rs_ResponsableAaux As New adodb.Recordset
'Dim rs_soldetaux As New adodb.Recordset
'Dim rstacumdet As New adodb.Recordset
'Dim rs_Bienes As New adodb.Recordset
'Dim rs_montador As New adodb.Recordset
'Dim rsgrupo As New adodb.Recordset
'Dim rstcodigo_detalle As New adodb.Recordset
'Dim rsdetalle As New adodb.Recordset
'
'Dim swgrabar, valida As Integer
'Dim correlsolic As Integer
'Dim correldetalle As Integer
'Dim swunidad, tot_form, prev_dev As Integer
'Dim marca1 As BookmarkEnum
'Dim ext1, tgn1 As Double
'Dim precuni, precTot, precSln, precTotV As Double
'Dim cantTot As Integer
''==== busquedas ====
'Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim PosibleApliqueFiltro As Boolean
'Dim msgSalir As String
'Dim queryinicial As String
'Dim queryinicial2 As String
'Dim sino As Integer
''MODI ALB
'Dim V_accion As String
''Para PCE, Pagos_espera y Pagos
'  'Dim rstdestino As New ADODB.Recordset
'  Dim rstorigen As New adodb.Recordset
'  Dim rstpagos As New adodb.Recordset
'  Dim rstpago_detalle As New adodb.Recordset
'  Dim rscorrelativo As New adodb.Recordset
'
'  Dim Proyecto1 As String
'  Dim Par_Codigo1 As String
'  Dim Organismo1 As String
'  Dim fte_codigo1 As String
'  Dim Org_Codigo1 As String
'  Dim pro_Programa1 As String
''  Dim Pro_SubPrograma1 As String
'  Dim Pro_Proyecto1 As String
'  Dim Pro_Actividad1 As String
'  Dim gestion1 As String
'  Dim uni_codigo1 As String
'  Dim COD_SOL As Integer
'  Dim codigo_categoria1 As String
'  Dim codigo_convenio1 As String
'  Dim Fte_contraparte1 As String
'  Dim Org_Contraparte1 As String
'
'  Dim por_fte_ext1 As Double
'  Dim por_fte_nal1 As Double
'  Dim codigo_pago1 As Double
'  Dim ges_gestion1 As String
'
'  Dim swpresup As Integer
'  Dim i As Integer
'  Dim j As Integer
'  Dim k As Integer
'  Dim v_por_fte(3, 3)
'  Dim tot_reg As Integer
'  Dim rectot As Integer
'
'  Dim rssolista As New adodb.Recordset
'  Dim rstao_solicitud_recibido As New adodb.Recordset
'  Dim swSubir As String
'  Dim swnuevo As Integer
'
'Private Sub adosalalm_MoveComplete(ByVal adReason As adodb.EventReasonEnum, ByVal pError As adodb.Error, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
'    If pRecordset.EOF Or pRecordset.BOF Then Exit Sub
'        Select Case pRecordset.EditMode
'        Case adEditNone
'            If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'            rstdetsalalm.Open "Select * from ao_detallesalidaalmacen where correlativo_salida = '" & pRecordset("correlativo_salida") & "'", db, adOpenDynamic, adLockOptimistic
''            Set DataGrid2.DataSource = Nothing
''            Set DataGrid2.DataSource = rstdetsalalm
''            DataGrid2.ReBind
'        End Select
'End Sub
'
'Private Sub adosolicitud_MoveComplete(ByVal adReason As adodb.EventReasonEnum, ByVal pError As adodb.Error, adStatus As adodb.EventStatusEnum, ByVal pRecordset As adodb.Recordset)
''JQA JUN/2008
'   If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then
'      If Not IsNull(adosolicitud.Recordset("codigo_solicitud")) And (adosolicitud.Recordset("ci") <> " ") Then
'         If adosolicitud.Recordset("tr_adjuntos") = "S" Then ChkTdr.Value = 1
'         If adosolicitud.Recordset("tr_adjuntos") = "N" Then ChkTdr.Value = 0
'         If adosolicitud.Recordset("tr_adjuntos") = "E" Then ChkTdr.Value = 2
'         If Not (IsNull(adosolicitud.Recordset("ci"))) Then
'            lblUni_codigo = IIf(IsNull(adosolicitud.Recordset("codigo_unidad")) = True, "ADMIN", adosolicitud.Recordset("codigo_unidad"))
'            Set rstao_solicitud_detalle = New adodb.Recordset
'            If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'            queryinicial2 = "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud
'            rstao_solicitud_detalle.Open queryinicial2, db, adOpenKeyset, adLockReadOnly
'            If rstao_solicitud_detalle.RecordCount > 0 Then
'                'DtGao_solicitud_detalle.Visible = True
'                DtGao_solicitud_detalle.Visible = False
'            Else
'                DtGao_solicitud_detalle.Visible = False
'            End If
'            Set adoao_solicitud_detalle.Recordset = rstao_solicitud_detalle
'            adoao_solicitud_detalle.Refresh
'         End If
'         If (adosolicitud.Recordset("estado_aprobado") <> "N") Then 'Or (Not IsNull(adosolicitud.Recordset("estado_aprobado")))
'                cmdAprueba.Visible = False
'                cmdDesaprueba.Visible = True
'                cmdDesaprueba.Enabled = True
'                CmdModCabeza.Enabled = False
'                CmdDelCabeza.Enabled = False
'               ' CmdLista.Enabled = False
'         Else
'                cmdAprueba.Visible = True
'                cmdAprueba.Enabled = True
'                cmdDesaprueba.Visible = False
'                CmdModCabeza.Enabled = True
'                CmdDelCabeza.Enabled = True
'         End If
'         If adosolicitud.Recordset("estado_aprobado") = "S" Then
'                cmdAprueba.Visible = False
'                cmdDesaprueba.Visible = True
'         Else
'                cmdAprueba.Visible = True
'                cmdDesaprueba.Visible = False
'         End If
'         If (adosolicitud.Recordset("estado_aprobado") = "E") Then
'                cmdAprueba.Visible = False
'                cmdDesaprueba.Visible = False
'         End If
'''''''''      If adosolicitud.Recordset!migrado <> "S" Then
'''''''''        cmdDesaprueba.Enabled = True
'''''''''      Else
'''''''''        cmdDesaprueba.Enabled = False
'''''''''      End If
'         If adosolicitud.Recordset!estado_enviado = "S" Or adosolicitud.Recordset!ESTADO_APROBADO = "E" Then
'                CmdEnviar.Enabled = False
'                CmdDelCabeza.Enabled = False
'                Command1.Enabled = False
'         Else
'                CmdEnviar.Enabled = True
'                CmdDelCabeza.Enabled = True
'                Command1.Enabled = True
'         End If
'            'jqa DIC-2008
'         'If adosolicitud.Recordset!Lista_adjunta = "S" Then
'         If SOLISTA = "A" Or adosolicitud.Recordset!Lista_adjunta = "S" Then
'            'If swnuevo <> 1 Then
'              DtGLista.Visible = True
'              'Frame10.Enabled = True
'              adoao_solicitud_lista.Visible = True
'              Set rstao_solicitud_lista = New adodb.Recordset
'              If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
'              rstao_solicitud_lista.Open "select * from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
'              Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
''              MsgBox Me.adoao_solicitud_lista.Recordset.RecordCount
'              If adoao_solicitud_lista.Recordset.RecordCount > 0 Then
'                SOLISTA = "A"
'                Set rstacumdet = New adodb.Recordset
'                If rstacumdet.State = 1 Then rstacumdet.Close
'                rstacumdet.Open "select sum(precio_compra) as precuni, sum(precio_venta) as precSln, sum(total_compra) as precTot, sum(total_venta) as precTotV, sum(cantidad) as cantTot from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' ", db, adOpenKeyset, adLockReadOnly
'                adoao_solicitud_lista.Caption = "TOTALES -->  Cantidad= '" & CStr(rstacumdet!cantTot) & "'   Precio Bruto= '" & CStr(rstacumdet!precSln) & "'   Total Bruto= '" & CStr(rstacumdet!precTotV) & "'  Precio Neto= '" & CStr(rstacumdet!precuni) & "'   Total Neto= '" & CStr(rstacumdet!precTot) & "'"
'                If rstacumdet.State = 1 Then rstacumdet.Close
'                DtGLista.Caption = "PRODUCTOS DE LA SOLICITUD Nro. " + Str(adosolicitud.Recordset("codigo_solicitud"))
'              Else
'                SOLISTA = "B"
'                DtGLista.Caption = ""
'              End If
'              adoao_solicitud_lista.Refresh
'            'End If
'         Else
'            DtGLista.Visible = False
'            'Frame10.Enabled = False
'            adoao_solicitud_lista.Visible = False
'         End If
'         'DtGLista.Caption = "DETALLE PRODUCTOS - SOLICITUD NRO. " + Str((adosolicitud.Recordset("CODIGO_SOLICITUD")))
'            'jqa
'      Else
'            ' por si es nuevo
'      End If
'      If IsNull(adosolicitud.Recordset!por_tiempo) Then
'        Text1.Text = "100"
'      Else
'        Text1.Text = CDbl(adosolicitud.Recordset!por_tiempo) * 100
'      End If
''      Text1.Text = IIf(IsNull(adosolicitud.Recordset!por_tiempo), 100, CDbl(adosolicitud.Recordset!por_tiempo) * 100)
'   Else
'        CmdModCabeza.Enabled = False
'        CmdDelCabeza.Enabled = False
''        CmdImpCabeza.Enabled = False
'        cmdAprueba.Enabled = False
'        CmdEnviar.Enabled = False
'        Command1.Enabled = False
''        CmdBusCabeza.Enabled = False
'   End If
''JQA JUN/2008
'
''ALB MODI PARA CAPTURA
''  If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then
''    If DtCvalor1.BoundText = "CC" Then
''        Label3.Visible = True
''        cmbSubCta2.Visible = True
''      Else
''        Label3.Visible = False
''        cmbSubCta2.Visible = False
''    End If
''      '''ALB
''    Select Case adosolicitud.Recordset!tipo_bien_Cta_doc
''        Case "A" 'Fondo Rotatorio Apertura
''          Lbltipo_bien_Cta_doc = "APERTURA"
''          If adosolicitud.Recordset!codigo_unidad_ant = "X" Then
''            Lbltipo_bien_Cta_doc1 = "...Cerrado"
''          Else
''            Lbltipo_bien_Cta_doc1 = ""
''          End If
'''        Case "AA" 'Fondo Rotatorio Apertura
'''          Lbltipo_bien_Cta_doc = "APERTURA"
'''          If adosolicitud.Recordset!codigo_unidad_ant = "X" Then
'''            Lbltipo_bien_Cta_doc1 = "...Cerrado"
'''          Else
'''            Lbltipo_bien_Cta_doc1 = " BALANCE"
'''          End If
''        Case "R"  'Fondo Rotatorio Rendición
''          Lbltipo_bien_Cta_doc = "RENDICION"
''          Lbltipo_bien_Cta_doc1 = " a: " & adosolicitud.Recordset!codigo_unidad_ant & "/" & adosolicitud.Recordset!codigo_solicitud_ant
''        Case "RR"  'Fondo Rotatorio Rendición
''          Lbltipo_bien_Cta_doc = "RENDICION"
''          Lbltipo_bien_Cta_doc1 = " a: " & adosolicitud.Recordset!codigo_unidad_ant & "/" & adosolicitud.Recordset!codigo_solicitud_ant
''        Case "C" 'Fondo Rotatorio Cierre
''          Lbltipo_bien_Cta_doc = "CIERRE"
''          Lbltipo_bien_Cta_doc1 = " a: " & adosolicitud.Recordset!codigo_unidad_ant & "/" & adosolicitud.Recordset!codigo_solicitud_ant
''        Case "CC" 'Fondo Rotatorio Cierre
''          Lbltipo_bien_Cta_doc = "CIERRE"
''          Lbltipo_bien_Cta_doc1 = "B.A. " '" a: " & adosolicitud.Recordset!codigo_unidad_ant & "/" & adosolicitud.Recordset!codigo_solicitud_ant
''    End Select
''
''    If (adosolicitud.Recordset("aprobado") > 0) Then 'Or (Not IsNull(adosolicitud.Recordset("aprobado")))
''        cmdAprueba.Visible = False
''        cmdDesaprueba.Visible = True
''        CmdEnviar.Enabled = True
''        CmdModCabeza.Enabled = False
''        CmdDelCabeza.Enabled = False
''      Else
''        cmdAprueba.Visible = True
''        cmdAprueba.Enabled = True
''        cmdDesaprueba.Visible = False
''        CmdModCabeza.Enabled = True
''        CmdDelCabeza.Enabled = True
''    End If
''
''    If (adosolicitud.Recordset("aprobado") = 2) Then
''        cmdAprueba.Visible = False
''        cmdDesaprueba.Visible = False
''    End If
''
''    If IsNull(adosolicitud.Recordset!subcta2) Then
''        cmbSubCta2.Text = ""
''      Else
''        Select Case adosolicitud.Recordset!subcta2
''          Case "01" '"Regulares" 'Cargos de Cuenta Regulares
''            Me.cmbSubCta2.Text = "Regulares"
''          Case "02" '"Otros" 'Cargos de Cuenta Otros
''            Me.cmbSubCta2.Text = "Otros"
''          Case "03"  '"PASE" 'Cargos de Cuenta PASE
''            Me.cmbSubCta2.Text = "PASE"
''        End Select
''    End If
''
''    If Not IsNull(adosolicitud.Recordset("codigo_solicitud")) Then
''       DtCvalor1.BoundText = IIf(IsNull(adosolicitud.Recordset!TipoF1), "", adosolicitud.Recordset!TipoF1)
'''            Set rstao_solicitud_detalle = New ADODB.Recordset
'''            If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'''            rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & Trim(adosolicitud.Recordset("ges_gestion")) & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " , db, adOpenKeyset, adLockOptimistic"
'''            Set adoao_solicitud_detalle.Recordset = rstao_solicitud_detalle
''
''
''      If adosolicitud.Recordset("tr_adjuntos") = "S" Then ChkTdr.Value = 1
''      If adosolicitud.Recordset("tr_adjuntos") = "N" Then ChkTdr.Value = 0
''      If adosolicitud.Recordset("tr_adjuntos") = "E" Then ChkTdr.Value = 2
''
''
''       txtnrosol.Text = IIf(IsNull(adosolicitud.Recordset("codigo_solicitud")) = True, False, adosolicitud.Recordset("codigo_solicitud"))
'''       DTPfechasol.Value = IIf(IsNull(adosolicitud.Recordset("fecha_solicitud")) = True, False, adosolicitud.Recordset("fecha_solicitud"))
''       txtjustifica.Text = IIf(IsNull(adosolicitud.Recordset("justificacion_solicitud")) = True, " ", adosolicitud.Recordset("justificacion_solicitud"))
''       txtterref.Text = IIf(IsNull(adosolicitud.Recordset("tr_adjuntos")) = True, " ", adosolicitud.Recordset("tr_adjuntos"))
''
''       dtccisol.Text = IIf(IsNull(adosolicitud.Recordset("ci")) = True, " ", adosolicitud.Recordset("ci"))
''       lblUni_codigo.Caption = adosolicitud.Recordset("codigo_unidad")
''       Dtcpaternosol.Text = dtccisol.BoundText
''       If Not (IsNull(adosolicitud.Recordset("ci"))) Then
''          If Not (adopuestosol.Recordset.BOF) Then adopuestosol.Recordset.MoveFirst
''                adopuestosol.Recordset.Find "ci = '" & Trim(dtccisol.Text) & "' ", , adSearchForward
''             If Not adopuestosol.Recordset.EOF Then
''                    dtcmaternosol.Text = IIf(IsNull(adopuestosol.Recordset("materno")) = True, " ", adopuestosol.Recordset("materno"))
''                    dtcnombresol.Text = IIf(IsNull(adopuestosol.Recordset("nombres")) = True, " ", adopuestosol.Recordset("nombres"))
'''                    dtccodpuesto.Text = IIf(IsNull(adopuestosol.Recordset("codigo_puesto")) = True, " ", adopuestosol.Recordset("codigo_puesto"))
'''                    dtcdenopuesto.Text = dtccodpuesto.BoundText
'''
'''                    dtccoduni.Text = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, " ", adopuestosol.Recordset("codigo_unidad"))
'''                    dtcdescripuni.Text = dtccoduni.BoundText
''
''             End If
''          End If
''
'''            Dtccibe.Text = IIf(IsNull(adosolicitud.Recordset("ci_aprueba")) = True, " ", adosolicitud.Recordset("ci_aprueba"))
'''            Dtcpaternobe.Text = Dtccibe.BoundText
'''            If Not (IsNull(adosolicitud.Recordset("ci_aprueba"))) Then
'''                If Not (adopuestobe.Recordset.BOF) Then adopuestobe.Recordset.MoveFirst
'''                adopuestobe.Recordset.Find "ci = '" & Trim(Dtccibe.Text) & "' ", , adSearchForward
'''                If Not adopuestobe.Recordset.EOF Then
'''                    Dtcmaternobe.Text = IIf(IsNull(adopuestobe.Recordset("materno")) = True, " ", adopuestobe.Recordset("materno"))
'''                    Dtcnombrebe.Text = IIf(IsNull(adopuestobe.Recordset("nombres")) = True, " ", adopuestobe.Recordset("nombres"))
'''                End If
'''            End If
''
'''            DtCDenominacion_moneda.BoundText = IIf(IsNull(adosolicitud.Recordset("tipo_moneda")) = True, "", adosolicitud.Recordset("tipo_moneda"))
'''            TxtTipo_cambio.Text = IIf(IsNull(adosolicitud.Recordset("tipo_caMBIO")) = True, 0, adosolicitud.Recordset("tipo_caMBIO"))
'''            TxtMonto_bolivianos.Text = IIf(IsNull(adosolicitud.Recordset("Monto_bolivianos")) = True, 0, adosolicitud.Recordset("Monto_bolivianos"))
'''            Txtmonto_dolares.Text = IIf(IsNull(adosolicitud.Recordset("monto_dolares")) = True, 0, adosolicitud.Recordset("monto_dolares"))
'''            TxtMonto_bolivianos_contra.Text = IIf(IsNull(adosolicitud.Recordset("monto_bolivianos_contra")) = True, 0, adosolicitud.Recordset("monto_bolivianos_contra"))
'''            Txtmonto_dolares_contra.Text = IIf(IsNull(adosolicitud.Recordset("monto_dolares_contra")) = True, 0, adosolicitud.Recordset("monto_dolares_contra"))
'''            DtCOrg_descripcion.BoundText = IIf(IsNull(adosolicitud.Recordset("org_codigo_contra")) = True, "", adosolicitud.Recordset("org_codigo_contra"))
'''            DtCOrg_descripcionExt.BoundText = IIf(IsNull(adosolicitud.Recordset("org_codigo_ext")) = True, "", adosolicitud.Recordset("org_codigo_ext"))
''
''
''          Set rstao_solicitud_detalle = New ADODB.Recordset
''          If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'''          If rstao_solicitud_detalle.State = 1 Then
''             queryinicial2 = "select * from ao_solicitud_detalle where (ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "') and (codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "') and (codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & ") "
''             rstao_solicitud_detalle.Open queryinicial2 & " order by codigo_detalle ", db, adOpenKeyset, adLockReadOnly
'''            rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " ", db, adOpenKeyset, adLockOptimistic
''             Set adoao_solicitud_detalle.Recordset = rstao_solicitud_detalle
'''          End If
''          adoao_solicitud_detalle.Refresh
''
''          If adosolicitud.Recordset("estado_aprobacion") = "S" Then
''            If (adosolicitud.Recordset!estatus = "S" Or adosolicitud.Recordset!estatus = "A") Or adosolicitud.Recordset!aprobado = 2 Then
''               CmdEnviar.Enabled = False
''               CmdDelCabeza.Enabled = False
''               cmdAprueba.Visible = False
''               cmdDesaprueba.Visible = False
''               CmdModCabeza.Enabled = False
''            Else
''               CmdEnviar.Enabled = True
''               CmdDelCabeza.Enabled = False
''               cmdAprueba.Visible = False
''               cmdDesaprueba.Visible = True
''               CmdModCabeza.Enabled = False
''            End If
''          Else
''             cmdAprueba.Visible = True
''             CmdEnviar.Enabled = True
''             cmdDesaprueba.Visible = False
''             CmdModCabeza.Enabled = True
''             CmdDelCabeza.Enabled = True
''          End If
''       Else
''            ' por si es nuevo
'''            Set rstao_solicitud_detalle = New ADODB.Recordset
'''            If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'''            rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle WHERE ges_gestion = '" & 0 & "' ", db, ad0OpenKeyset, adLockOptimistic
'''            Set Adodetallesolicitud.Recordset = rstao_solicitud_detalle
'''            Adodetallesolicitud.Refresh
''
'''            dtccodpoa.Text = " "
'''            dtcdespoa.Text = dtccodpoa.BoundText
'''
''            dtccisol.Text = " "
''            Dtcpaternosol.Text = dtccisol.BoundText
''
''            dtcmaternosol.Text = " "
''            dtcnombresol.Text = " "
''
'''            dtccodpuesto.Text = " "
'''            dtcdenopuesto.Text = dtccodpuesto.BoundText
'''
'''            dtccoduni.Text = " "
'''            dtcdescripuni.Text = dtccoduni.BoundText
''
'''            Dtccibe.Text = " "
'''            Dtcpaternobe.Text = Dtccibe.BoundText
''
'''            Dtcmaternobe.Text = " "
'''            Dtcnombrebe.Text = " "
''            txtjustifica.Text = ""
'''            TxtMonto_bolivianos.Text = 0
''
'''            Lblaprobado.Visible = False
'''            Lblestatus.Visible = False
''
''       End If
'''        CmdModCabeza.Enabled = True
'''        CmdDelCabeza.Enabled = True
'''        CmdImpCabeza.Enabled = True
'''        CmdLista.Enabled = True
'''        CmdAprueba.Enabled = True
'''        CmdEnviar.Enabled = True
''       CmdBusCabeza.Enabled = True
''       CmdImpCabeza.Enabled = True
''    Else
'''        CmdModCabeza.Enabled = False
'''        CmdDelCabeza.Enabled = False
'''alb
''        CmdImpCabeza.Enabled = False
'''        CmdLista.Enabled = False
'''        CmdAprueba.Enabled = False
'''        CmdEnviar.Enabled = False
''        CmdBusCabeza.Enabled = False
''    End If
'End Sub
'
'Private Sub CmdCancelaDet_Click()
'  swgrabar = 0
'  Call cerea
'  swnuevo = 0
'  Frmnavega.Enabled = True
'  Frame10.Enabled = True
'  'DtGLista.Enabled = True
'  FrmEditaDet.Enabled = False
'  'CmdLista.Enabled = True
'  'CmdListaMod.Enabled = True
'  Call OptFilGral1_Click
'  rstao_solicitud_lista.Requery
'  adoao_solicitud_lista.Refresh
'  SSTab1.Tab = 0
'  SSTab1.TabEnabled(1) = True
'  SSTab1.TabEnabled(0) = True
'  SOLISTA = "B"
'  FrmGrabaDet.Visible = False
'End Sub
'
'Private Sub cmdElige_Click()
'    AlFrmCreaMaterial.Show
'End Sub
'
'Private Sub CmdListaAnula_Click()
'  If adosolicitud.Recordset("estado_aprobado") = "N" And adosolicitud.Recordset("estado_enviado") = "N" Then
'    sino = MsgBox("Está seguro de eliminar este registro", vbYesNo + vbQuestion, "Atención ...")
'    If sino = vbYes Then
'      adoao_solicitud_lista.Recordset.Delete
'      adoao_solicitud_lista.Recordset.Update
'      rstao_solicitud_lista.Requery
'      adoao_solicitud_lista.Refresh
'    End If
'  Else
'    MsgBox "No se puede ANULAR un registro Aprobado ó Enviado !! ", vbExclamation
'  End If
'End Sub
'
'Private Sub sstab1_Click(PreviousTab As Integer)
'    If SSTab1.Tab = 0 Then
'        'SSTab1.TabEnabled(0) = True
'        'SSTab1.TabEnabled(1) = False
'    Else
'      'If adoao_solicitud_lista.Recordset!codigo_solicitud = 0 Then
'      If swnuevo = 1 Then
'        'MsgBox "ERR"
'        FrmEditaDet.Visible = True
'        'DtGLista.Visible = True
'        Frame10.Enabled = True
'        adoao_solicitud_lista.Visible = True
'      Else
'        If adosolicitud.Recordset!Lista_adjunta = "S" Then
'        'If SOLISTA = "A" Then
'           FrmEditaDet.Visible = True
'           'DtGLista.Visible = True
'           Frame10.Enabled = True
'           adoao_solicitud_lista.Visible = True
'         Else
'           FrmEditaDet.Visible = False
'           'DtGLista.Visible = False
'           Frame10.Enabled = False
'           adoao_solicitud_lista.Visible = False
'         End If
'      End If
'    End If
'End Sub
'
'Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']" Or KeyAscii = 8, KeyAscii, 0)
'End Sub
'
'Private Sub valida2()
'   Set rssolista = New adodb.Recordset
'   If rssolista.State = 1 Then rssolista.Close
'   rssolista.Open "select * from ao_solicitud_lista where codigo_solicitud = " & COD_SOL & " and codigo_unidad = '" & uni_codigo1 & "' and CodDetalle = '" & Dtccodbien & "' ", db, adOpenKeyset, adLockOptimistic
''   Set AdoUnidad.Recordset = rssolista
''   AdoUnidad.Refresh
'    If rssolista.RecordCount > 0 Then
'        valida = 2
'        MsgBox "El producto ya fue registrado. Intente nuevamente !! ", vbExclamation + vbOKOnly, "Validación de Datos"
'        Exit Sub
'    End If
'End Sub
'
'Private Sub CmdGrabaDet_Click()
' valida = 1
' Call valida2
' If valida = 1 Then
'  If Not IsNumeric(TxtCantidad.Text) Then
'     MsgBox "El dato registrado debe ser un Valor Numérico Válido.", vbExclamation + vbOKOnly, "Validación de Datos"
'     Exit Sub
'  End If
'  If Dtccodbien <> "" And Val(TxtCantidad.Text) >= 0 Then
'    db.BeginTrans
'    If swnuevo = 1 Then
'      adoao_solicitud_lista.Recordset!ges_gestion = gestion1    'adosolicitud.Recordset("ges_gestion")     'Year(Date)
'      adoao_solicitud_lista.Recordset!codigo_unidad = uni_codigo1   'adosolicitud.Recordset("codigo_UNIDAD")    'Trim(DtcUnidad.Text)
'      adoao_solicitud_lista.Recordset!codigo_solicitud = COD_SOL    'adosolicitud.Recordset("codigo_solicitud") 'Trim(txtnrosol.Text)
''      adoao_solicitud_lista.Recordset!id_beneficiario = id_beneficiario1
'    End If
''    If swnuevo = 2 Then
''      rstdestino.Open "select * from ao_solicitud_lista where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud & " and id_beneficiario = " & adoao_solicitud_lista.Recordset!id_beneficiario, db, adOpenKeyset, adLockOptimistic
''    End If
'      adoao_solicitud_lista.Recordset!CodGrupo = DtcCodGrupo.Text                        'Grupo Bien
'      adoao_solicitud_lista.Recordset!cod_MONTADOR = DtcSubgrupo.Text                    'Sub-Grupo Bien
'      adoao_solicitud_lista.Recordset!codDetalle = Trim(Dtccodbien.Text)                 'Codigo Bien
'      adoao_solicitud_lista.Recordset!doc_identidad = Trim(Dtccodbien.Text)              'Codigo de Bien
'      adoao_solicitud_lista.Recordset!profesion = Trim(Dtcdesbien.Text)      'Descripcion del Bien
'      adoao_solicitud_lista.Recordset!DescDetalle = Trim(Dtcdesbien.Text)
'      adoao_solicitud_lista.Recordset!razon_s = Trim(Txtrazon_s)                         'Caracteristicas del Bien
'      adoao_solicitud_lista.Recordset!grado_instruccion = DtcdesAnt                      'Nombre Antiguo del Producto
'      adoao_solicitud_lista.Recordset!aplanilla = IIf(TxtCantidad = "", 1, TxtCantidad)  'Cantidad Solicitada del Bien
'      adoao_solicitud_lista.Recordset!cantidad = IIf(TxtCantidad = "", 1, TxtCantidad)  'Cantidad Solicitada del Bien
'      adoao_solicitud_lista.Recordset!Nro_pagos = Val(TxtPrecioU)                        'Precio Unitario Bruto del Bien
'      adoao_solicitud_lista.Recordset!Precio_venta = Val(TxtPrecioU)                        'Precio Unitario Bruto del Bien
'      adoao_solicitud_lista.Recordset!total_venta = adoao_solicitud_lista.Recordset!cantidad * adoao_solicitud_lista.Recordset!Precio_venta         'Precio Total Bruto
'      adoao_solicitud_lista.Recordset!Precio_Compra = (Val(TxtPrecioU) * (1 - Val(Txt_porcentaje) / 100))      'Precio Unitario Neto del Bien       'CDbl(TxtPrecioC.Text)   'DtcPrecioC.Text  '
'      adoao_solicitud_lista.Recordset!Total_compra = adoao_solicitud_lista.Recordset!cantidad * adoao_solicitud_lista.Recordset!Precio_Compra      'Precio Total Neto
'      adoao_solicitud_lista.Recordset!Monto_solicitud_dl = adoao_solicitud_lista.Recordset!cantidad * adoao_solicitud_lista.Recordset!Precio_Compra      'Precio Total Neto
'      adoao_solicitud_lista.Recordset!aunidad = Dtc_UniMed.Text                          'Unidad de Medidad del Bien
'
'      adoao_solicitud_lista.Recordset!Tipo_Beneficiario = "F"  'Trim(lbltipo_beneficiario)
'      adoao_solicitud_lista.Recordset!usr_usuario = glusuario
'      adoao_solicitud_lista.Recordset!fecha_registro = Format(Date, "dd/mm/yyyy")
'      adoao_solicitud_lista.Recordset!hora_registro = Format(Time, "hh:mm:ss")
'      adoao_solicitud_lista.Recordset.Update
'      'JQA 04/2008
''     adoao_solicitud_lista.Recordset = rstdestino
'      db.CommitTrans
'    adosolicitud.Recordset("Lista_adjunta") = "S"
'    If swnuevo = 1 Then
''     Call abre_solicitud_lista
'      'rstao_solicitud_lista.Requery
'      'adoao_solicitud_lista.Refresh
'      adoao_solicitud_lista.Recordset.MoveLast
'    End If
'    If swnuevo = 2 Then
'      marca1 = adoao_solicitud_lista.Recordset.Bookmark
''     Call abre_solicitud_lista
''     rstao_solicitud_lista.Update
''     rstao_solicitud_lista.Requery
''     Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
'      If rstao_solicitud_lista.RecordCount > 0 Then
'        adoao_solicitud_lista.Recordset.Move marca1 - 1
'      End If
'    End If
'    Frmnavega.Enabled = True
'    Frame10.Enabled = True
'    'DtGLista.Enabled = True
'    FrmEditaDet.Enabled = False
'    'CmdLista.Enabled = True
'    'CmdListaMod.Enabled = True
'    swnuevo = 0
'    'rstAo_solicitud!Lista_adjunta = "S"
'    Call GRABADET
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(1) = False
'    SSTab1.TabEnabled(0) = True
'    SOLISTA = "A"
'    TxtPrecioU.Enabled = False
'    TxtPrecioC.Enabled = False
'    FrmGrabaDet.Visible = False
'  Else
'    MsgBox "Error en registro de datos, vuelva a intentar.!", vbCritical, ""
'  End If
' End If
'End Sub
'
'Private Sub CmdListaMod_Click()
'  If adosolicitud.Recordset("estado_enviado") = "N" And adosolicitud.Recordset!Lista_adjunta = "S" Then
'    'marca1 = adosolicitud.Recordset.BookMark
'    'marca1 = adoao_solicitud_lista.Recordset.BookMark
'    Frmnavega.Enabled = False
'    Frame10.Enabled = False
'    swgrabar = 0
'    swnuevo = 2
'    'adoao_solicitud_lista.Recordset.Move marca1 - 1
'    SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'    FrmEditaDet.Visible = True
'    FrmEditaDet.Enabled = True
'    FrmGrabaDet.Visible = True
'    If GlSistema = "C" Or GlSistema = "Z" Then
'        TxtPrecioU.Enabled = False
'        TxtPrecioC.Enabled = True
'    End If
'  Else
'    MsgBox "No se puede Modificar un registro Aprobado, Enviado o Inexistente!! ", vbExclamation
'  End If
'End Sub
'
'Private Sub ChkTdr_Click()
'    If ChkTdr.Value = 0 Then txtterref.Text = "N"
'    If ChkTdr.Value = 1 Then txtterref.Text = "S"
'    If ChkTdr.Value = 2 Then txtterref.Text = "E"
'End Sub
'
'Private Sub cmdAprueba_Click()
'If adosolicitud.Recordset!Lista_adjunta = "S" Then
'    sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'    If sino = vbYes Then
'        Dim rstdestino As New adodb.Recordset
'        Set rstdestino = New adodb.Recordset
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
'        If Not rstdestino.BOF Then rstdestino.MoveFirst
'        If Not rstdestino.BOF And Not rstdestino.EOF Then
'            rstdestino("estado_aprobado") = "S"
'            rstdestino.Update
'        End If
'        If rstdestino.State = 1 Then rstdestino.Close
'        marca1 = adosolicitud.Recordset.Bookmark
'        adosolicitud.Recordset.Requery
'        adosolicitud.Refresh
'        adosolicitud.Recordset.Move marca1 - 1
'    End If
'Else
'    MsgBox "No se puede APROBAR. Debe registrar el detalle del registro ...", , "Atención"
'End If
'End Sub
'
'Private Sub CmdBusCabeza_Click()
''JQA
''  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
''  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'  PosibleApliqueFiltro = False
'  Dim rsNada As adodb.Recordset
'  Dim GrSqlAux As String
'  Set ClBuscaGrid = New ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.EsTdbGrid = False
'  Set ClBuscaGrid.GridTrabajo = DataGrid1
'  ClBuscaGrid.QueryUtilizado = queryinicial
'  Set ClBuscaGrid.RecordsetTrabajo = adosolicitud.Recordset
'  ClBuscaGrid.CamposVisibles = "110"
'  ClBuscaGrid.Ejecutar
'  PosibleApliqueFiltro = True
'End Sub
'
'Private Sub cmdDesaprueba_Click()
'  If adosolicitud.Recordset!estado_enviado = "S" Then
'    MsgBox "No se puede DESAPROBAR si el registro está ENVIADO ...", vbCritical, "Advertencia !"
'  Else
'    If adosolicitud.Recordset!ESTADO_APROBADO = "S" Then
'       sino = MsgBox("Esta seguro de DESAPROBAR el registro ?", vbYesNo, "Confirmando")
'       If sino = vbYes Then
'          Dim rstdestino As New adodb.Recordset
'          Set rstdestino = New adodb.Recordset
'          If rstdestino.State = 1 Then rstdestino.Close
'          rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
'          If Not rstdestino.BOF Then rstdestino.MoveFirst
'          If Not rstdestino.BOF And Not rstdestino.EOF Then
'            rstdestino("estado_aprobado") = "N"
'            rstdestino.Update
'          End If
'          If rstdestino.State = 1 Then rstdestino.Close
'          marca1 = adosolicitud.Recordset.Bookmark
'          'adosolicitud.Recordset.AddNew
'          adosolicitud.Recordset.Cancel
'          adosolicitud.Refresh
'          adosolicitud.Recordset.Move marca1 - 1
'       End If
'    Else
'        MsgBox "No se puede DESAPROBAR si el registro NO está APROBADO ...", vbCritical, "Advertencia !"
'    End If
'  End If
'End Sub
'
'Private Sub CmdDetallePoa_Click()
'  If adosolicitud.Recordset.RecordCount > 0 Then
'  marca1 = adosolicitud.Recordset.Bookmark
'   ''''' ALB
'  FrmPoasCapturaALB.Lblformulario = "F01"
'  FrmPoasCapturaALB.lblges_gestion = adosolicitud.Recordset!ges_gestion
'  FrmPoasCapturaALB.lblcodigo_unidad = adosolicitud.Recordset!codigo_unidad
'  FrmPoasCapturaALB.lblcodigo_solicitud = adosolicitud.Recordset!codigo_solicitud
'  FrmPoasCapturaALB.lbltipo_beneficiario = "N" 'adosolicitud.Recordset!tipo_beneficiario
'  FrmPoasCapturaALB.tXTaprobado = adosolicitud.Recordset!aprobado
'  FrmPoasCapturaALB.Lbltipo_bien_Cta_doc = adosolicitud.Recordset!tipo_bien_Cta_doc
'  FrmPoasCapturaALB.Lblcategoria_Cta_doc = adosolicitud.Recordset!subcta2
'  FrmPoasCapturaALB.Show vbModal
'  adosolicitud.Refresh
'  Else
'    MsgBox "No Existen Registros ", vbInformation, "Formulario 1"
'  End If
'  If adosolicitud.Recordset.RecordCount > 0 Then
'    adosolicitud.Recordset.Move marca1 - 1
'  End If
'End Sub
'
'Private Sub CmdEnviar_Click()
'    If adosolicitud.Recordset!ESTADO_APROBADO = "S" Then
'      swunidad = 0
'      sino = MsgBox("Esta seguro de ENVIAR el registro Aprobado ? (Nota: Ya no se podrá Desaprobar!)...", vbYesNo, "Confirmando ...")
'      If sino = vbYes Then
'        'JQA 04/2008
'        marca1 = adosolicitud.Recordset.Bookmark
'        Call val_presupF01(adosolicitud.Recordset, GlNombFor)
'        Set rs_soldetaux = New adodb.Recordset
'        If rs_soldetaux.State = 1 Then rs_soldetaux.Close
'        rs_soldetaux.Open "select ges_gestion, codigo_unidad, codigo_solicitud, org_codigo_contra, sum(monto_bolivianos) as monto_bolivianos, sum(monto_dolares) as monto_dolares, sum(monto_bolivianos_contra) as monto_bolivianos_contra, sum(monto_dolares_contra) as monto_dolares_contra, tipo_moneda from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & "  GROUP BY ges_gestion, codigo_unidad, codigo_solicitud, org_codigo_contra, tipo_moneda ", db, adOpenKeyset, adLockOptimistic
'        If rs_soldetaux.RecordCount > 0 Then
'            deCD.dbo_ap_Graba_No_Objecion_D2 rs_soldetaux!ges_gestion, rs_soldetaux!codigo_solicitud, "PD", "FIN_PROPIO", "10", "00", "00", "00", rstAo_solicitud!caracteristicas, rstAo_solicitud!Observaciones, glusuario, Format(Date, "dd/mm/yyyy"), Format(Time, "hh:mm:ss"), rs_soldetaux!org_codigo_contra, 0, rstAo_solicitud!formulario, 0, "D", "10", rstAo_solicitud!codigo_unidad, rstAo_solicitud!codigo_unidad, rs_soldetaux!monto_dolares, rs_soldetaux!monto_dolares_contra, rs_soldetaux!monto_Bolivianos, rs_soldetaux!monto_bolivianos_contra, rs_soldetaux!tipo_moneda, "0"
'
'            db.Execute " UPDATE AO_SOLICITUD SET estado_enviado='S' " & _
'            "WHERE (ao_Solicitud.Ges_Gestion) = '" & Me.adosolicitud.Recordset!ges_gestion & "' and " & _
'            "(ao_Solicitud.codigo_unidad) = '" & Me.adosolicitud.Recordset!codigo_unidad & "' and " & _
'            "(ao_Solicitud.codigo_solicitud) =  " & Me.adosolicitud.Recordset!codigo_solicitud & ""
'            adosolicitud.Refresh
'            If marca1 > 1 Then
'                adosolicitud.Recordset.Move marca1 - 1
'            End If
''            db.Execute "update AlCldetalle set AlCldetalle.stockingreso= av_acumula_compra.cantidad_cotizada from AlCldetalle, av_acumula_compra Where AlCldetalle.CodGrupo = av_acumula_compra.CodGrupo And AlCldetalle.cod_MONTADOR = av_acumula_compra.cod_MONTADOR And AlCldetalle.codDetalle = av_acumula_compra.codDetalle"
''            db.Execute "update AlCldetalle set StockActual= Stockinicial + stockingreso - StockSalida"
'        Else
'            MsgBox "NO se registro el detalle del registro, intente nuevamente...", vbExclamation, "-"
'        End If
'        adosolicitud.Refresh
'        'JQA 04/2008
'      End If
'    Else
'        MsgBox "No se puede ENVIAR. Debe Aprobar previamente el registro ...", , "Atención"
'    End If
'End Sub
'
'Private Sub CmdImporta_Click()
''    FrmImporta.Show
'End Sub
'
'Private Sub ActualStock()
'
'End Sub
'
''Private Sub CmdOKunidad_Click()
''    swunidad = 1
''        If swunidad = 1 Then
''            Dim rstpagos As New ADODB.Recordset
''            Set rstpagos = New ADODB.Recordset
''            If rstpagos.State = 1 Then rstpagos.Close
''            rstpagos.Open "select * from pagos where GES_gestion = '5000'", db, adOpenKeyset, adLockOptimistic
''
''            rstpagos.AddNew
''                rstpagos("ges_gestion") = adosolicitud.Recordset("ges_gestion")
''                rstpagos("org_codigo") = DataCombo1.Text   'adosolicitud.Recordset("formulario")
''                rstpagos("codigo_pago") = "" 'genera jorge
''                rstpagos("codigo_solicitud") = adosolicitud.Recordset("codigo_solicitud")
''                rstpagos("formulario") = adosolicitud.Recordset("formulario")
''                rstpagos("codigo_unidad") = adosolicitud.Recordset("codigo_unidad")
''                rstpagos("monto_bolivianos") = adosolicitud.Recordset("monto_bolivianos")
''                rstpagos("estado_compromiso") = "N"
''                rstpagos("justificacion") = adosolicitud.Recordset("justificacion_solicitud")
''             rstpagos.Update
''        End If
''End Sub
'
''Private Sub CmdVinculado_Click()
''
''   adosolicitud.Refresh
''    Data1.DatabaseName = " "
''    Data1.DatabaseName = App.path & "\pragma.mdb"  '"C:\Captura\pragma.mdb"
''    Set fs = CreateObject("Scripting.FileSystemObject")
''    Set A = fs.CreateTextFile(App.path & "\TMPBANCOTXT.txt", True)   'fs.CreateTextFile("c:\captura\TMPBANCOTXT.txt", True)
'''    a.WriteLine ("Pruebita de grabar texto")
''    A.Close
''    Data1.Refresh
''    If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then adosolicitud.Recordset.MoveFirst
''    While Not adosolicitud.Recordset.EOF
''        Data1.Recordset.AddNew
''        Data1.Recordset("campo1") = adosolicitud.Recordset("ges_gestion")
''        Data1.Recordset("campo2") = adosolicitud.Recordset("formulario")
''        Data1.Recordset("campo3") = adosolicitud.Recordset("codigo_unidad")
''        Data1.Recordset("campo4") = adosolicitud.Recordset("justificacion_solicitud")
''        Data1.Recordset("campo5") = adosolicitud.Recordset("ci")
''        Data1.Recordset("campo6") = adosolicitud.Recordset("Codigo_puesto")
''        Data1.Recordset("campo7") = adosolicitud.Recordset("CI_aprueba")
''        Data1.Recordset("campo8") = adosolicitud.Recordset("estado_aprobacion")
''        Data1.Recordset("campo9") = adosolicitud.Recordset("codigo_poa")
''        Data1.Recordset("campo10") = adosolicitud.Recordset("tipo_moneda")
''
''        Data1.Recordset("campo11") = adosolicitud.Recordset("codigo_solicitud")
''        Data1.Recordset("campo12") = adosolicitud.Recordset("monto_bolivianos")
''        Data1.Recordset("campo13") = adosolicitud.Recordset("monto_dolares")
''        Data1.Recordset("campo14") = adosolicitud.Recordset("Tipo_cambio")
''        Data1.Recordset("campo15") = adosolicitud.Recordset("duracion_estimada_numero")
''        Data1.Recordset("campo15") = adosolicitud.Recordset("por_tiempo")
''
''        Data1.Recordset.Update
''        adosolicitud.Recordset.MoveNext
''    Wend
''   Frmexporta.Show
''   Unload Me
''
''End Sub
'
'Private Sub CmdAddCabeza_Click()
'    Frame3.Enabled = True
'    Frame10.Enabled = False
'    Frame10.Visible = False
''    Frame1.Visible = False
'    frmabm.Visible = False
'    Frmnavega.Enabled = False
'    frmgrabcabeza.Visible = True
'    Frasolic.Enabled = True
'    swgrabar = 1
'    Call cerea
'    adosolicitud.Refresh
'    adosolicitud.Recordset.AddNew
'    'FrmApertura.Visible = True
'    DTPfechasol.Value = Date
'    DTPfechasol.CheckBox = True
'    Txt_porcentaje.Text = 0.25
'    DtcUnidad.Enabled = True
'    If GlSistema = "A" Or GlSistema = "Z" Then
'        DtcUnidad.Text = "ALFAPARF"
'        Dtccodbien.Text = "27"
'        Dtcdesbien.Text = "DIVISION PROFESIONAL - ALFAPARF"
'        DtcPOA.Text = "2.3.1.2.1"
'        DtcPOADes.Text = "Productos de Belleza ALFAPARF"
'        Dtccibe.Text = "9000000001"
'        Dtcpaternobe.Text = "ARGENBOL S.A."
'    End If
'    If GlSistema = "B" Or GlSistema = "Z" Then
'        DtcUnidad.Text = "ISSUE"
'        Dtccodbien.Text = "27"
'        Dtcdesbien.Text = "DIVISION PROFESIONAL - ISSUE"
'        DtcPOA.Text = "2.3.2.2.1"
'        DtcPOADes.Text = "Productos de Belleza ISSUE"
'        Dtccibe.Text = "9000000007"
'        Dtcpaternobe.Text = "PRODUCTOS ISSUE"
'    End If
'    If GlSistema = "C" Or GlSistema = "Z" Then
'        DtcUnidad.Text = "NUTRIPET"
'        Dtccodbien.Text = "27"
'        Dtcdesbien.Text = "DIVISION PROFESIONAL - NUTRIPET"
'        DtcPOA.Text = "1.1.1.2.1"
'        DtcPOADes.Text = "NUTRIPET Comida para perros"
'        Dtccibe.Text = "9000000005"
'        Dtcpaternobe.Text = "NUTRIPET - INDUSTRIA DE ALIMENTOS PARA ANIMALES"
'    End If
'    If GlSistema = "D" Or GlSistema = "Z" Then
'        DtcUnidad.Text = "PLANALTO"
'        Dtccodbien.Text = "27"
'        Dtcdesbien.Text = "PRODUCTOS VARIOS EN TIENDA"
'        DtcPOA.Text = "1.1.1.2.1"
'        DtcPOADes.Text = "Productos varios en tienda"
'        Dtccibe.Text = "9000000010"
'        Dtcpaternobe.Text = "PROVEEDOR DE VARIOS PRODUCTOS"
'    End If
'    Set rs_ResponsableAaux = New adodb.Recordset
'    If rs_ResponsableAaux.State = 1 Then rs_ResponsableAaux.Close
'    rs_ResponsableAaux.Open "select * from unidad_responsable WHERE status='S' AND codigo_unidad = '" & DtcUnidad.Text & "' AND apellido_esposo = '-' ", db, adOpenKeyset, adLockOptimistic
'    If rs_ResponsableAaux.RecordCount > 0 Then
'        dtccisol.Text = rs_ResponsableAaux!CI
'        Dtcpaternosol.Text = rs_ResponsableAaux!paterno
'        dtcmaternosol.Text = rs_ResponsableAaux!materno
'        dtcnombresol.Text = rs_ResponsableAaux!nombres
'    End If
'    DataGrid1.Visible = False
'    'ALB
'    Lbltipo_bien_Cta_doc.Visible = False
'    Lbltipo_bien_Cta_doc.Caption = "APERTURA"
'
'
'
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'
''    If SSTab1.Tab = 0 Then
''        SSTab1.TabEnabled = True
''        FrmEditaDet.Visible = False
''    Else
''        SSTab1.TabEnabled = True
''        FrmEditaDet.Visible = False
''    End If
'
''  Select Case V_accion
''    Case "APE" 'Fondo Rotatorio Apertura
''      Call cerea
''      Frasolic.Enabled = True
''      Frame2.Enabled = True
'''      FraBoleta.Visible = False
''      ' ini antiguo boletas 30/05/2001
''      'Optboleta1.Value = False
''      ' fin antiguo boletas 30/05/2001
''    Case "REN" 'Fondo Rotatorio Rendición
''      Frasolic.Enabled = False
''      If adosolicitud.Recordset!tipo_bien_Cta_doc = "A" Then
''        Frame2.Enabled = False
''      End If
'''      FraBoleta.Visible = False
'''      If adosolicitud.Recordset!tipo_bien_Cta_doc = "AA" Then
'''        Frame2.Enabled = True
'''      End If
''    Case "CIE" 'Fondo Rotatorio Cierre
''      Frasolic.Enabled = False
''      If adosolicitud.Recordset!tipo_bien_Cta_doc = "A" Then
''        Frame2.Enabled = False
''      End If
'''      FraBoleta.Visible = True
'''      If adosolicitud.Recordset!tipo_bien_Cta_doc = "AA" Then
'''        Frame2.Enabled = True
'''      End If
''    Case "CIE_BA"
''      Frasolic.Enabled = True
''      Frame2.Enabled = True
''      Lbltipo_bien_Cta_doc.Caption = "CIERRE B.A."
'''      FraBoleta.Visible = True
''  End Select
'End Sub
'
'Private Sub CmdGraBoleta_Click()
''  FraModBoleta.Visible = False
'End Sub
'
'Private Sub CmdImpCabeza_Click()
''JQA JUN/2008
'If adosolicitud.Recordset!Lista_adjunta = "S" Then
'    Dim co As New adodb.Command
''    Dim rs As New ADODB.Recordset
''    rs.Open "select * from ao_solicitud_detalle where ges_gestion='" & Me.adosolicitud.Recordset!ges_gestion & "' and " & _
''            "codigo_unidad='" & Me.adosolicitud.Recordset!codigo_unidad & "' and " & _
''            "codigo_solicitud=" & Me.adosolicitud.Recordset!codigo_solicitud, db, adOpenStatic, adLockReadOnly
'    CryF01.ReportFileName = App.Path & "\formularioSeNTRADA\C01_F11.rpt"
'    CryF01.WindowShowRefreshBtn = True
'    'MsgBox rs.RecordCount
'      CryF01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
'      CryF01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
'    'Call CREAVISTAF11          'JQA JUN-2008
'    CryF01.StoredProcParam(0) = Me.adosolicitud.Recordset!ges_gestion
'    CryF01.StoredProcParam(1) = Me.adosolicitud.Recordset!codigo_unidad
'    CryF01.StoredProcParam(2) = Me.adosolicitud.Recordset!codigo_solicitud
'    iresult = CryF01.PrintReport
'    If iresult <> 0 Then MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical, "Error de impresión"
'Else
'    MsgBox "No se puede Imprimir. Debe registrar el detalle del registro ...", , "Atención"
'End If
'' ... (Jorge)
''  Dim V_cmbSubCta2 As String
''  Dim IResult As Variant
''  Dim PaternoS, MaternoS, NombreS, PaternoB, MaternoB, NombreB, UnidadT As String
''  Dim rsunidad As New ADODB.Recordset
''  Set rsunidad = New ADODB.Recordset
''
''  adoao_solicitud_detalle.Refresh
''  '---- ini version actual
'''  db.Execute "drop view av_F01"
'''  db.Execute "create view av_F01 as SELECT ao_Solicitud.ges_gestion, ao_Solicitud.codigo_unidad, ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, ao_Solicitud.fecha_solicitud, ao_Solicitud_detalle.tipo_moneda, ao_Solicitud_detalle.monto_bolivianos, ao_Solicitud_detalle.monto_dolares, ao_Solicitud.fecha_registro, ao_Solicitud.CI_aprueba, ao_Solicitud_detalle.monto_bolivianos_contra, ao_Solicitud_detalle.monto_dolares_contra, ao_Solicitud_detalle.Tipo_cambio " & _
'''            " FROM ao_Solicitud INNER JOIN ao_Solicitud_detalle ON (ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud) AND (ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad) AND (ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion) " & _
'''            " WHERE (((ao_Solicitud.codigo_solicitud)= " & txtnrosol & ") AND ((ao_Solicitud.codigo_unidad)='" & Trim(lblUni_codigo) & "')) "
''  '---- fin version actual
''
''  db.Execute "drop view av_F01"
''  db.Execute "create view av_F01 as SELECT ao_Solicitud.ges_gestion, ao_Solicitud.codigo_unidad, ao_Solicitud.codigo_solicitud, ao_Solicitud.formulario, ao_Solicitud.justificacion_solicitud, ao_Solicitud.CI, ao_Solicitud.fecha_solicitud, ao_Solicitud_detalle.tipo_moneda, ao_Solicitud_detalle.monto_bolivianos, ao_Solicitud_detalle.monto_dolares, ao_Solicitud.fecha_registro, ao_Solicitud.CI_aprueba, ao_Solicitud.estatus, ao_Solicitud.estado_aprobacion, ao_Solicitud_detalle.monto_bolivianos_contra, ao_Solicitud_detalle.monto_dolares_contra, ao_Solicitud_detalle.Tipo_cambio , ao_Solicitud_detalle.codigo_convenio, por_fte_ext, por_fte_nal  " & _
''            " FROM ao_Solicitud INNER JOIN ao_Solicitud_detalle ON (ao_Solicitud.codigo_solicitud = ao_Solicitud_detalle.codigo_solicitud) AND (ao_Solicitud.codigo_unidad = ao_Solicitud_detalle.codigo_unidad) AND (ao_Solicitud.Ges_Gestion = ao_Solicitud_detalle.Ges_Gestion) " & _
''            " WHERE (((ao_Solicitud.codigo_solicitud)= " & txtnrosol & ") AND ((ao_Solicitud.codigo_unidad)='" & Trim(lblUni_codigo) & "')) "
''
''    'and codigo_unidad='" & adosolicitud.Recordset!codigo_unidad & "'
''
''  rsunidad.Open "select * from fc_unidad_ejecutora where codigo_unidad='" & adosolicitud.Recordset("codigo_unidad") & "' ", db, adOpenKeyset, adLockReadOnly
''  CryF01.WindowShowRefreshBtn = True
''  If rsunidad.RecordCount > 0 Then
''     CryF01.Formulas(0) = "UnidadT = '" & rsunidad("uni_descripcion_larga") & "' "
''  Else
''     CryF01.Formulas(0) = "UnidadT = '" & "-" & "' "
''  End If
''  CryF01.Formulas(1) = "PaternoS = '" & Dtcpaternosol.Text & "' "
''  CryF01.Formulas(2) = "MaternoS = '" & dtcmaternosol.Text & "' "
''  CryF01.Formulas(3) = "NombreS = '" & dtcnombresol.Text & "' "
''  CryF01.Formulas(4) = "PaternoB = '" & Dtcpaternobe.Text & "' "
''  CryF01.Formulas(5) = "MaternoB = '" & Dtcmaternobe.Text & "' "
''  CryF01.Formulas(6) = "NombreB = '" & Dtcnombrebe.Text & "' "
''  CryF01.Formulas(7) = "Tipo = '" & Lbltipo_bien_Cta_doc.Caption & "' "
''  V_cmbSubCta2 = " : " & cmbSubCta2.Text
''  CryF01.Formulas(14) = "tipof1 = '" & DtCvalor1.Text & " " & IIf(cmbSubCta2.Visible = False, "", V_cmbSubCta2) & "' "
''  If cmbSubCta2.Visible = True And cmbSubCta2.Text = "PASE" Then
''    CryF01.Formulas(15) = "titmunicipio = 'UNIDAD EDUCATIVA:' "
''    CryF01.Formulas(16) = "codmuni = '" & adoao_solicitud_detalle.Recordset!aux3 & "' "
''    CryF01.Formulas(17) = "desmuni = '" & fbusmuni(adoao_solicitud_detalle.Recordset!aux3) & "' "
''  Else
''    CryF01.Formulas(15) = "titmunicipio = ''"
''    CryF01.Formulas(16) = "codmuni = '' "
''    CryF01.Formulas(17) = "desmuni = '' "
''  End If
''      If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "CC" Then
'''        FraBoleta.Visible = True
''      Else
'''        FraBoleta.Visible = False
''      End If
''
''  If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "CC" Then
''    CryF01.Formulas(8) = "boletaTit = 'BOLETA BANCARIA :'"
''
''
''    ' aqui 30/05/2001
'''    CryF01.Formulas(9) = "boletanumero = '" & Txtnro_boleta.Text & "' " 'TxtPlanilla_depto
''    ' aqui 30/05/2001
''
''
'''    CryF01.Formulas(10) = "boletaCtaTit = 'Cuenta :'"
'''    CryF01.Formulas(11) = "boletaCta = '" & DtCcta_codigo.Text & "' "  'DtCBco_codigo.Text & "' "
'''    CryF01.Formulas(12) = "boletamontoTit = 'Monto Bs.:'"
'''    CryF01.Formulas(13) = "boletamonto = '" & TDBNmontoBs & "' " 'TDBNnro_pagos & "' "
''  Else
''    CryF01.Formulas(8) = "boletaTit = ' '"
''    CryF01.Formulas(9) = "boletanumero = ' ' "
''    CryF01.Formulas(10) = "boletaCtaTit = ' '"
''    CryF01.Formulas(11) = "boletaCta = ' '"
''    CryF01.Formulas(12) = "boletamontoTit = ' '"
''    CryF01.Formulas(13) = "boletamonto = ' '"
''  End If
''  CryF01.ReportFileName = App.Path & "\FormulariosEntrada\S01_F01.rpt"
''  IResult = CryF01.PrintReport
''  If IResult <> 0 Then
''     MsgBox CryF01.LastErrorNumber & " : " & CryF01.LastErrorString, vbCritical + vbOKOnly, "Error..."
''  End If
'End Sub
'
'Private Sub CmdLista_Click()
'  marca1 = adosolicitud.Recordset.Bookmark
'  If adosolicitud.Recordset("estado_enviado") = "N" Then
'    'marca1 = adosolicitud.Recordset.BookMark
'   ' frmabm.Visible = False
'   ' frmgrabcabeza.Visible = True
'    swgrabar = 1
''    Call cerea
'    swnuevo = 1
'    'rstao_solicitud_lista.AddNew
'    Frmnavega.Enabled = False
'    Frame10.Enabled = False
'    FrmEditaDet.Visible = True
'    FrmEditaDet.Enabled = True
'    FrmGrabaDet.Visible = True
'    gestion1 = adosolicitud.Recordset("ges_gestion")
'    uni_codigo1 = adosolicitud.Recordset("CODIGO_UNIDAD")
'    COD_SOL = adosolicitud.Recordset("codigo_solicitud")
'    'adosolicitud.Recordset.Move marca1 - 1
'   'parametro = "ges_gestion" + " <> " + "'2000'"
'   parametro2 = "Cod_marca = '" & DtcMarca.Text & "' "
'   parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
'   Call ABRE_SOL_LISTA
'   SSTab1.Tab = 1
'    SSTab1.TabEnabled(1) = True
'    SSTab1.TabEnabled(0) = False
'
'   ' Call ABRE_SOL_LISTA
'    adoao_solicitud_lista.Recordset.AddNew
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
'End Sub
'
'Private Sub CmdModCabeza_Click()
'    Frame3.Enabled = True
'    Frame10.Enabled = False
'    Frame10.Visible = False
''    txtnrosol.Enabled = False
'    DTPfechasol.SetFocus
'    DTPfechasol.CheckBox = True
'    frmabm.Visible = False
'    Frmnavega.Enabled = False
'    frmgrabcabeza.Visible = True
'    DtcUnidad.Enabled = False
'    swgrabar = 0
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = False
'End Sub
'
'Private Sub CmdDelCabeza_Click()
'If adosolicitud.Recordset!ESTADO_APROBADO = "N" Then
'  Dim rsterr As New adodb.Recordset
'    sino = MsgBox("Està seguro de eliminar este registro", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'      Set rsterr = New adodb.Recordset
'      If rsterr.State = 1 Then rsterr.Close
'      rsterr.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'      If rsterr.RecordCount > 0 Then
'        rsterr!ESTADO_APROBADO = "E"
'        rsterr!estado_enviado = "E"
'        rsterr.Update
'      End If
'      If rsterr.State = 1 Then rsterr.Close
'      marca1 = adosolicitud.Recordset.Bookmark
'      'rstAo_solicitud.Requery
'      Set adosolicitud.Recordset = rstAo_solicitud
'      adosolicitud.Refresh
'      Set adosolicitud.Recordset = rstAo_solicitud
'      If marca1 > 1 Then
'        adosolicitud.Recordset.Move marca1 - 1
'      End If
'    End If
'Else
'    MsgBox "No se puede ANULAR. El registro ya esta Aprobado y/o Enviado ...", , "Atención"
'End If
'End Sub
'
'Private Sub CmdSalCabeza_Click()
'    sino = MsgBox("Esta Seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
'    If sino = vbYes Then
''        adosolicitud.Recordset.Close
'        If rstdetsalalm.State = 1 Then rstdetsalalm.Close
'        If rstpoa.State = 1 Then rstpoa.Close
'        If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'        If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'        If rstfc_partida_gasto.State = 1 Then rstfc_partida_gasto.Close
''        If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
''        If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'        Unload Me
'    End If
'End Sub
'
''Private Sub CmdDetCabeza_Click()
''    frmabm.Visible = False
''    frmdetalle.Visible = True
''    FraDetalle.Visible = True
''    Frmnavega.Enabled = False
''    If Not (Adodetallesolicitud.Recordset.BOF) Then Adodetallesolicitud.Recordset.MoveFirst
''
''End Sub
'
'Private Sub CmdGraCabeza_Click()
'    Frame3.Enabled = False
'    Call grabar
'    Frame10.Enabled = True
'    Frame10.Visible = True
'    DataGrid1.Visible = True
'    frmabm.Visible = True
'    frmgrabcabeza.Visible = False
'    Frmnavega.Enabled = True
'    Frame3.Enabled = False
'    FrmApertura.Visible = False
'    Frasolic.Enabled = True
'    DtcUnidad.Enabled = True
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = True
'End Sub
'
'Private Sub CmdCanCabeza_Click()
'    adosolicitud.Refresh
'    frmabm.Visible = True
''    frmdetalle.Visible = False
'    frmgrabcabeza.Visible = False
'    If adosolicitud.Recordset.RecordCount > 0 Then
'        adosolicitud.Recordset.CancelUpdate
'    End If
'    Frmnavega.Enabled = True
'    Frame3.Enabled = False
'    FrmApertura.Visible = False
'    DataGrid1.Visible = True
'    DtcUnidad.Enabled = True
'    Frame10.Enabled = True
'    Frame10.Visible = True
'    adosolicitud.Refresh
'    SSTab1.Tab = 0
'    SSTab1.TabEnabled(0) = True
'    SSTab1.TabEnabled(1) = True
'End Sub
'
'Private Sub Dtc_UniMed_Click(Area As Integer)
'    Dtccodbien.BoundText = Dtc_UniMed.BoundText
'    Dtcdesbien.BoundText = Dtc_UniMed.BoundText
'    DtcPrecioU.BoundText = Dtc_UniMed.BoundText
'    DtcPrecioUV.BoundText = Dtc_UniMed.BoundText
'    DtcdesAnt.BoundText = Dtc_UniMed.BoundText
'    DtcCodAnt.BoundText = Dtc_UniMed.BoundText
'    DtcCodUniv.BoundText = Dtc_UniMed.BoundText
'    DtcPrecioC.BoundText = Dtc_UniMed.BoundText
'    DtcCodGrupoP.BoundText = Dtc_UniMed.BoundText
'    DtcSubgrupoP.BoundText = Dtc_UniMed.BoundText
'End Sub
'
'Private Sub Dtccibe_Click(Area As Integer)
'    Dtcpaternobe.BoundText = Dtccibe.BoundText
''    Dtcmaternobe.BoundText = Dtccibe.BoundText
''    Dtcnombrebe.BoundText = Dtccibe.BoundText
'End Sub
'
''Private Sub dtccisol_Change()
'''  lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
'''  Call fbuscaunidad
''End Sub
'
'Private Sub dtccisol_Click(Area As Integer)
'    Dtcpaternosol.BoundText = dtccisol.BoundText
'    dtcmaternosol.BoundText = dtccisol.BoundText
'    dtcnombresol.BoundText = dtccisol.BoundText
'
''    Dtcpaternosol.Text = dtccisol.BoundText
''    If Not (IsNull(dtccisol.Text)) And Trim(dtccisol.Text) <> "" Then
''        If Not (adopuestosol.Recordset.BOF) Then adopuestosol.Recordset.MoveFirst
''        adopuestosol.Recordset.Find "ci = '" & Trim(dtccisol.Text) & "' ", , adSearchForward
''        If Not adopuestosol.Recordset.EOF Then
''            dtcmaternosol.Text = IIf(IsNull(adopuestosol.Recordset("materno")) = True, " ", adopuestosol.Recordset("materno"))
''            dtcnombresol.Text = IIf(IsNull(adopuestosol.Recordset("nombres")) = True, " ", adopuestosol.Recordset("nombres"))
'''            lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
'''            Call fbuscaunidad
'''            dtccodpuesto.Text = IIf(IsNull(adopuestosol.Recordset("codigo_puesto")) = True, " ", adopuestosol.Recordset("codigo_puesto"))
'''            dtcdenopuesto.Text = dtccodpuesto.BoundText
'''
'''            dtccoduni.Text = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, " ", adopuestosol.Recordset("codigo_unidad"))
'''            dtcdescripuni.Text = dtccoduni.BoundText
''        End If
''    End If
'End Sub
'
'Private Sub DtcCodAnt_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcCodAnt.BoundText
'    Dtcdesbien.BoundText = DtcCodAnt.BoundText
'    DtcPrecioU.BoundText = DtcCodAnt.BoundText
'    DtcPrecioUV.BoundText = DtcCodAnt.BoundText
'    DtcdesAnt.BoundText = DtcCodAnt.BoundText
'    Dtc_UniMed.BoundText = DtcCodAnt.BoundText
'    DtcCodUniv.BoundText = DtcCodAnt.BoundText
'    DtcPrecioC.BoundText = DtcCodAnt.BoundText
'    DtcCodGrupoP.BoundText = DtcCodAnt.BoundText
'    DtcSubgrupoP.BoundText = DtcCodAnt.BoundText
'End Sub
'
'Private Sub Dtccodbien_Click(Area As Integer)
'    Dtcdesbien.BoundText = Dtccodbien.BoundText
'    DtcPrecioU.BoundText = Dtccodbien.BoundText
'    DtcPrecioUV.BoundText = Dtccodbien.BoundText
'    Dtc_UniMed.BoundText = Dtccodbien.BoundText
'    DtcdesAnt.BoundText = Dtccodbien.BoundText
'    DtcCodAnt.BoundText = Dtccodbien.BoundText
'    DtcCodUniv.BoundText = Dtccodbien.BoundText
'    DtcPrecioC.BoundText = Dtccodbien.BoundText
'    DtcCodGrupoP.BoundText = Dtccodbien.BoundText
'    DtcSubgrupoP.BoundText = Dtccodbien.BoundText
'End Sub
'
'Private Sub Dtccodbien_LostFocus()
'  'If swnuevo = 1 Then
'    Txtrazon_s.Text = Dtcdesbien.Text
'    TxtPrecioU.Text = DtcPrecioU.Text
'    TxtPrecioC.Text = DtcPrecioC.Text
'  'End If
'  DtcCodGrupo.Text = DtcCodGrupoP.Text
'  DtcSubgrupo.Text = DtcSubgrupoP.Text
'End Sub
'
'Private Sub DtcCodGrupo_Click(Area As Integer)
'    DtcGrupo.BoundText = DtcCodGrupo.BoundText
'End Sub
'
'Private Sub DtcCodGrupoP_Click(Area As Integer)
'    Dtcdesbien.BoundText = DtcCodGrupoP.BoundText
'    DtcPrecioU.BoundText = DtcCodGrupoP.BoundText
'    DtcPrecioUV.BoundText = DtcCodGrupoP.BoundText
'    Dtc_UniMed.BoundText = DtcCodGrupoP.BoundText
'    DtcdesAnt.BoundText = DtcCodGrupoP.BoundText
'    DtcCodAnt.BoundText = DtcCodGrupoP.BoundText
'    DtcCodUniv.BoundText = DtcCodGrupoP.BoundText
'    DtcPrecioC.BoundText = DtcCodGrupoP.BoundText
'    Dtccodbien.BoundText = DtcCodGrupoP.BoundText
'    DtcSubgrupoP.BoundText = DtcCodGrupoP.BoundText
'End Sub
'
'Private Sub DtcCodUniv_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcCodUniv.BoundText
'    Dtcdesbien.BoundText = DtcCodUniv.BoundText
'    DtcPrecioU.BoundText = DtcCodUniv.BoundText
'    DtcPrecioUV.BoundText = DtcCodUniv.BoundText
'    DtcdesAnt.BoundText = DtcCodUniv.BoundText
'    Dtc_UniMed.BoundText = DtcCodUniv.BoundText
'    DtcCodAnt.BoundText = DtcCodUniv.BoundText
'    DtcPrecioC.BoundText = DtcCodUniv.BoundText
'    DtcCodGrupoP.BoundText = DtcCodUniv.BoundText
'    DtcSubgrupoP.BoundText = DtcCodUniv.BoundText
'End Sub
'
'Private Sub DtcdesAnt_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcdesAnt.BoundText
'    Dtcdesbien.BoundText = DtcdesAnt.BoundText
'    Dtc_UniMed.BoundText = DtcdesAnt.BoundText
'    DtcPrecioU.BoundText = DtcdesAnt.BoundText
'    DtcPrecioUV.BoundText = DtcdesAnt.BoundText
'    DtcCodAnt.BoundText = DtcdesAnt.BoundText
'    DtcCodUniv.BoundText = DtcdesAnt.BoundText
'    DtcPrecioC.BoundText = DtcdesAnt.BoundText
'    DtcCodGrupoP.BoundText = DtcdesAnt.BoundText
'    DtcSubgrupoP.BoundText = DtcdesAnt.BoundText
'End Sub
'
'Private Sub Dtcdesbien_Click(Area As Integer)
'    Dtccodbien.BoundText = Dtcdesbien.BoundText
'    DtcPrecioU.BoundText = Dtcdesbien.BoundText
'    DtcPrecioUV.BoundText = Dtcdesbien.BoundText
'    Dtc_UniMed.BoundText = Dtcdesbien.BoundText
'    DtcdesAnt.BoundText = Dtcdesbien.BoundText
'    DtcCodAnt.BoundText = Dtcdesbien.BoundText
'    DtcCodUniv.BoundText = Dtcdesbien.BoundText
'    DtcPrecioC.BoundText = Dtcdesbien.BoundText
'    DtcCodGrupoP.BoundText = Dtcdesbien.BoundText
'    DtcSubgrupoP.BoundText = Dtcdesbien.BoundText
'End Sub
'
'Private Sub Dtcdesbien_LostFocus()
'  'If swnuevo = 1 Then
'    Txtrazon_s.Text = Dtcdesbien.Text
'    TxtPrecioU.Text = DtcPrecioU.Text
'    TxtPrecioC.Text = DtcPrecioC.Text
'  'End If
'  DtcCodGrupo.Text = DtcCodGrupoP.Text
'  DtcSubgrupo.Text = DtcSubgrupoP.Text
'End Sub
'
'Private Sub DtcGrupo_Click(Area As Integer)
'    DtcCodGrupo.BoundText = DtcGrupo.BoundText
''    Call pSubGrupo(DtcCodGrupo.BoundText)
'End Sub
'
'Private Sub dtcmaternosol_Click(Area As Integer)
'    Dtcpaternosol.BoundText = dtcmaternosol.BoundText
'    dtcnombresol.BoundText = dtcmaternosol.BoundText
'    dtccisol.Text = Dtcpaternosol.BoundText
''    lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
''    Call fbuscaunidad
'End Sub
'
'Private Sub dtcnombresol_Click(Area As Integer)
'    Dtcpaternosol.BoundText = dtcnombresol.BoundText
'    dtcmaternosol.BoundText = dtcnombresol.BoundText
'    dtccisol.Text = Dtcpaternosol.BoundText
''    lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
''    Call fbuscaunidad
'End Sub
'
'Private Sub Dtcpaternobe_Click(Area As Integer)
'    Dtccibe.BoundText = Dtcpaternobe.BoundText
''    Dtcmaternobe.BoundText = Dtcpaternobe.BoundText
''    Dtcnombrebe.BoundText = Dtcpaternobe.BoundText
'End Sub
'
''Private Sub Dtcpaternosol_Change()
'''  lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
'''  Call fbuscaunidad
''End Sub
'
'Private Sub Dtcpaternosol_Click(Area As Integer)
'    dtcmaternosol.BoundText = Dtcpaternosol.BoundText
'    dtcnombresol.BoundText = Dtcpaternosol.BoundText
'    dtccisol.BoundText = Dtcpaternosol.BoundText
'
''    dtccisol.Text = Dtcpaternosol.BoundText
''    If Not (IsNull(dtccisol.Text)) And (Trim(dtccisol.Text) <> "") Then
''        If Not (adopuestosol.Recordset.BOF) Then adopuestosol.Recordset.MoveFirst
''        adopuestosol.Recordset.Find "ci = '" & Trim(dtccisol.Text) & "' ", , adSearchForward
''        If Not adopuestosol.Recordset.EOF Then
''            dtcmaternosol.Text = IIf(IsNull(adopuestosol.Recordset("materno")) = True, " ", adopuestosol.Recordset("materno"))
''            dtcnombresol.Text = IIf(IsNull(adopuestosol.Recordset("nombres")) = True, " ", adopuestosol.Recordset("nombres"))
'''            lblUni_codigo = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, "", adopuestosol.Recordset("codigo_unidad"))
'''            Call fbuscaunidad
'''            dtccodpuesto.Text = IIf(IsNull(adopuestosol.Recordset("codigo_puesto")) = True, " ", adopuestosol.Recordset("codigo_puesto"))
'''            dtcdenopuesto.Text = dtccodpuesto.BoundText
'''            dtccoduni.Text = IIf(IsNull(adopuestosol.Recordset("codigo_unidad")) = True, " ", adopuestosol.Recordset("codigo_unidad"))
'''            dtcdescripuni.Text = dtccoduni.BoundText
''        End If
''    End If
'End Sub
'
'Private Sub DtcPOA_Click(Area As Integer)
'    DtcPOADes.BoundText = DtcPOA.BoundText
'End Sub
'
'Private Sub DtcPOADes_Click(Area As Integer)
'    DtcPOA.BoundText = DtcPOADes.BoundText
'End Sub
'
'Private Sub DtcPrecioC_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcPrecioC.BoundText
'    Dtcdesbien.BoundText = DtcPrecioC.BoundText
'    Dtc_UniMed.BoundText = DtcPrecioC.BoundText
'    DtcdesAnt.BoundText = DtcPrecioC.BoundText
'    DtcCodAnt.BoundText = DtcPrecioC.BoundText
'    DtcCodUniv.BoundText = DtcPrecioC.BoundText
'    DtcPrecioU.BoundText = DtcPrecioC.BoundText
'    DtcPrecioUV.BoundText = DtcPrecioC.BoundText
'    DtcCodGrupoP.BoundText = DtcPrecioC.BoundText
'    DtcSubgrupoP.BoundText = DtcPrecioC.BoundText
'End Sub
'
'Private Sub DtcPrecioU_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcPrecioU.BoundText
'    Dtcdesbien.BoundText = DtcPrecioU.BoundText
'    Dtc_UniMed.BoundText = DtcPrecioU.BoundText
'    DtcdesAnt.BoundText = DtcPrecioU.BoundText
'    DtcCodAnt.BoundText = DtcPrecioU.BoundText
'    DtcCodUniv.BoundText = DtcPrecioU.BoundText
'    DtcPrecioC.BoundText = DtcPrecioU.BoundText
'    DtcCodGrupoP.BoundText = DtcPrecioU.BoundText
'    DtcSubgrupoP.BoundText = DtcPrecioU.BoundText
'    DtcPrecioUV.BoundText = DtcPrecioU.BoundText
'End Sub
'
'Private Sub DtcPrecioUV_Click(Area As Integer)
'    Dtccodbien.BoundText = DtcPrecioUV.BoundText
'    Dtcdesbien.BoundText = DtcPrecioUV.BoundText
'    Dtc_UniMed.BoundText = DtcPrecioUV.BoundText
'    DtcdesAnt.BoundText = DtcPrecioUV.BoundText
'    DtcCodAnt.BoundText = DtcPrecioUV.BoundText
'    DtcCodUniv.BoundText = DtcPrecioUV.BoundText
'    DtcPrecioC.BoundText = DtcPrecioUV.BoundText
'    DtcCodGrupoP.BoundText = DtcPrecioUV.BoundText
'    DtcSubgrupoP.BoundText = DtcPrecioUV.BoundText
'    DtcPrecioU.BoundText = DtcPrecioUV.BoundText
'End Sub
'
'Private Sub DtcSubgrupo_Click(Area As Integer)
'    DtcSubgrupoDes.BoundText = DtcSubgrupo.BoundText
'End Sub
'
'Private Sub DtcSubgrupoDes_Click(Area As Integer)
'    DtcSubgrupo.BoundText = DtcSubgrupoDes.BoundText
'    'Call pProducto(DtcSubgrupo.BoundText)
'End Sub
'
'Private Sub DtcSubgrupoP_Click(Area As Integer)
'    Dtcdesbien.BoundText = DtcSubgrupoP.BoundText
'    DtcPrecioU.BoundText = DtcSubgrupoP.BoundText
'    DtcPrecioUV.BoundText = DtcSubgrupoP.BoundText
'    Dtc_UniMed.BoundText = DtcSubgrupoP.BoundText
'    DtcdesAnt.BoundText = DtcSubgrupoP.BoundText
'    DtcCodAnt.BoundText = DtcSubgrupoP.BoundText
'    DtcCodUniv.BoundText = DtcSubgrupoP.BoundText
'    DtcPrecioC.BoundText = DtcSubgrupoP.BoundText
'    Dtccodbien.BoundText = DtcSubgrupoP.BoundText
'    DtcCodGrupoP.BoundText = DtcSubgrupoP.BoundText
'End Sub
'
'Private Sub DtcUnidad_Click(Area As Integer)
'    DtcUnidadDes.BoundText = DtcUnidad.BoundText
'End Sub
'
'Private Sub DtcUnidad_LostFocus()
'    Set rstrc_personalSoli = New adodb.Recordset
'    If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'    rstrc_personalSoli.Open "select * from unidad_responsable WHERE codigo_unidad='" & DtcUnidad.Text & "' ", db, adOpenKeyset, adLockReadOnly
'    Set adopuestosol.Recordset = rstrc_personalSoli
'    adopuestosol.Refresh
'    lblUni_codigo.Caption = DtcUnidad.Text
'
'    Set rstpoaAux = New adodb.Recordset
'    If rstpoaAux.State = 1 Then rstpoaAux.Close
'    'rstpoa.Open "select par_codigo,* from fc_Relacionador_poa_ppto where (codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "') and (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
'    If GlSistema = "Z" Then
'        rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtcUnidad.Text & "') and (nivel  = 5)  ", db, adOpenKeyset, adLockReadOnly
'    Else
'        rstpoaAux.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtcUnidad.Text & "') and (nivel  = 5) and (ARCHIVO = '" & GlSistema & "') ", db, adOpenKeyset, adLockReadOnly
'    End If
'    If rstpoaAux.RecordCount > 0 Then
'        DtcPOA.Text = rstpoaAux!codigo_poa
'        DtcPOADes.Text = rstpoaAux!descripcion_poa
'    End If
''    Set AdoPOA.Recordset = rstpoa
''    AdoPOA.Refresh
'End Sub
'
'Private Sub DtcUnidadDes_Click(Area As Integer)
'    DtcUnidad.BoundText = DtcUnidadDes.BoundText
'End Sub
'
'Private Sub DTPfechasol_Change()
'    txtGes_gestion = CStr(Year(DTPfechasol.Value))
'End Sub
'
'Private Sub Form_Load()
''jqa JUN/2008
'   GlNombFor = "F01"
'   label7.Caption = glusuario
'
'   Set rstFc_unidad_ejecutora = New adodb.Recordset
'   If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'   rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora WHERE UNI_ACTIVO='S' ", db, adOpenKeyset, adLockReadOnly
'   Set AdoUnidad.Recordset = rstFc_unidad_ejecutora
'   AdoUnidad.Refresh
'
'   Set rstrc_personalSoli = New adodb.Recordset
'   If rstrc_personalSoli.State = 1 Then rstrc_personalSoli.Close
'   rstrc_personalSoli.Open "select * from unidad_responsable WHERE status='S' ORDER BY PATERNO ", db, adOpenKeyset, adLockReadOnly
'   Set adopuestosol.Recordset = rstrc_personalSoli
'   adopuestosol.Refresh
'
'   Call OptFilGral1_Click
'
'   If (Not adosolicitud.Recordset.BOF) And (Not adosolicitud.Recordset.EOF) Then
'      dtccisol.Text = IIf(IsNull(adosolicitud.Recordset("ci")) = True, "-", adosolicitud.Recordset("ci"))
'        'Dtcpaternosol.Text = dtccisol.BoundText
'   End If
'
'   Set rs_montador = New adodb.Recordset
'   If rs_montador.State = 1 Then rs_montador.Close
'   rs_montador.Open "select * from Al_Montador order by descripcion ", db, adOpenKeyset, adLockReadOnly
'   Set AdoMontador.Recordset = rs_montador
'   AdoMontador.Refresh
'
'   Set rsgrupo = New adodb.Recordset
'   If rsgrupo.State = 1 Then rsgrupo.Close
'   rsgrupo.Open "select * from ALCLGrupo order by DescGrupo ", db, adOpenKeyset, adLockReadOnly
'   Set AdoGrupo.Recordset = rsgrupo
'   AdoGrupo.Refresh
'
'    Set rstpoa = New adodb.Recordset
'    If rstpoa.State = 1 Then rstpoa.Close
'    'rstpoa.Open "select * from fc_Relacionador_poa_ppto where (codigo_unidad = '" & DtcUnidad.Text & "') and (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
'    rstpoa.Open "select * from fc_Relacionador_poa_ppto where (nivel  = 5) ORDER BY codigo_poa", db, adOpenKeyset, adLockReadOnly
'    Set AdoPOA.Recordset = rstpoa
'    AdoPOA.Refresh
'
'    'modi alb
'    'FrmApertura.Visible = False
'    ''''Lbltipo_bien_Cta_doc.Caption = ""
'    Set rscc_parametros = New adodb.Recordset
'    If rscc_parametros.State = 1 Then rscc_parametros.Close
'    rscc_parametros.Open " select * from cc_parametros where valor2 = 'F1A' order by valor1 ", db, adOpenKeyset, adLockReadOnly
'    Set Adocc_parametros.Recordset = rscc_parametros
'    Adocc_parametros.Refresh
'    '
'    Set rstrc_personalCargo = New adodb.Recordset
'    If rstrc_personalCargo.State = 1 Then rstrc_personalCargo.Close
'    rstrc_personalCargo.Open "select * from fc_beneficiario where tipo_beneficiario>0 ORDER BY denominacion_beneficiario", db, adOpenKeyset, adLockReadOnly
'    Set adopuestobe.Recordset = rstrc_personalCargo
'    adopuestobe.Refresh
'
''    Set rstao_solicitud_lista = New ADODB.Recordset
''    If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
''    rstao_solicitud_lista.Open "select * from ao_solicitud_lista order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
''    'rstao_solicitud_lista.Open "select * from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
''    Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
''    adoao_solicitud_lista.Refresh
'
'   parametro = "ges_gestion" + " <> " + "'2011'"
'   parametro2 = "cod_marca" + " <> " + "'0'"
'   'parametro = "ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
'   Call ABRE_SOL_LISTA
'
'   FrmEditaDet.Enabled = False
'   swnuevo = 0
'   SSTab1.Tab = 0
'   SSTab1.TabEnabled(0) = True
'   SSTab1.TabEnabled(1) = False
'	Call SeguridadSet(Me)
End Sub
'
'Private Sub ABRE_SOL_LISTA()
'   Set rstao_solicitud_lista = New adodb.Recordset
'   If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
'   rstao_solicitud_lista.Open "select * from ao_solicitud_lista where " & parametro & " order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
'   'rstao_solicitud_lista.Open "select * from ao_solicitud_lista where ges_gestion = '" & adosolicitud.Recordset!ges_gestion & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud & " and codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' order by CodGrupo, COD_MONTADOR, profesion", db, adOpenKeyset, adLockOptimistic
'   Set adoao_solicitud_lista.Recordset = rstao_solicitud_lista
'   adoao_solicitud_lista.Refresh
'   If adoao_solicitud_lista.Recordset.RecordCount > 0 Then
'        SOLISTA = "A"
'   Else
'        SOLISTA = "B"
'   End If
'   Set rs_Bienes = New adodb.Recordset
'   If rs_Bienes.State = 1 Then rs_Bienes.Close
'   'rs_Bienes.Open "select * from AlClDetalle where Cod_marca = '" & DtcMarca.Text & "' order by DescDetalle ", db, adOpenKeyset, adLockReadOnly
'   'rs_Bienes.Open "select * from AlClDetalle where " & parametro2 & " order by DescDetalle ", DB, adOpenKeyset, adLockReadOnly
'   rs_Bienes.Open "select * from AlClDetalle order by DescDetalle ", db, adOpenKeyset, adLockReadOnly
'   Set ado_bienes.Recordset = rs_Bienes
'   ado_bienes.Refresh
'
''     Set rsformulacion = New ADODB.Recordset       'Abrir POA
''    If rsformulacion.State = 1 Then rsformulacion.Close
''    queryinicial = "Select * from POA where " & parametro & " "
''    rsformulacion.Open queryinicial, db, adOpenKeyset, adLockOptimistic
''    rsformulacion.Sort = "codigo_estrategico, codigo_especifico, codigo_componente, codigo_actividad, codigo_tarea, codigo_insumo "
'End Sub
'
'Private Sub grabar()
''JQA JUN/2008
'    db.BeginTrans
'    If swgrabar = 1 Then
'       'rstdestino.Open "select * from Ao_solicitud where formulario = '0'", db, adOpenDynamic, adLockOptimistic
'       Dim correlsolic As Integer
'       Dim rstao_solicitud_correl As New adodb.Recordset
'       Set rstao_solicitud_correl = New adodb.Recordset
'       If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
'       rstao_solicitud_correl.Open "select * from ao_solicitud_correl where codigo_unidad = '" & Trim(DtcUnidad.Text) & "' ", db, adOpenDynamic, adLockOptimistic
'        '
'        'rstao_solicitud_correl.Find "formulario = 'F03'", , adSearchForward
'       If rstao_solicitud_correl.RecordCount > 0 Then
'          If Not (rstao_solicitud_correl.BOF) Then rstao_solicitud_correl.MoveFirst
'          rstao_solicitud_correl("correl_solicitud") = rstao_solicitud_correl("correl_solicitud") + 1
'          correlsolic = rstao_solicitud_correl("correl_solicitud")
'          rstao_solicitud_correl.Update
'       Else
'          rstao_solicitud_correl.AddNew
'          rstao_solicitud_correl("codigo_unidad") = Trim(lblUni_codigo.Caption)
'          rstao_solicitud_correl("correl_solicitud") = 1
'          correlsolic = rstao_solicitud_correl("correl_solicitud")
'          rstao_solicitud_correl.Update
'       End If
'       If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
'       adosolicitud.Recordset("ges_gestion") = glGestion        'CStr(Year(DTPfechasol.Value))
'       adosolicitud.Recordset("codigo_solicitud") = correlsolic
'       adosolicitud.Recordset("codigo_unidad") = DtcUnidad.Text
'       adosolicitud.Recordset("Lista_adjunta") = "N"
'     End If
'        adosolicitud.Recordset("fecha_solicitud") = DTPfechasol.Value
'        adosolicitud.Recordset("CI") = dtccisol.Text
'        adosolicitud.Recordset("CI_aprueba") = Dtccibe.Text
'        adosolicitud.Recordset("codigo_poa") = DtcPOA.Text
'        adosolicitud.Recordset("codigo_bien") = Dtccodbien.Text
'        adosolicitud.Recordset("caracteristicas") = Txtcaracteristicas.Text
'        adosolicitud.Recordset("justificacion_solicitud") = Txtcaracteristicas.Text     'txtjustifica.Text
'        adosolicitud.Recordset("observaciones") = Txtobservaciones.Text
'        adosolicitud.Recordset("tr_adjuntos") = IIf(IsNull(txtterref.Text), "N", txtterref.Text)
'        adosolicitud.Recordset("TipoF1") = "CC"      'DtCvalor1.BoundText  'jqa jun/2008 Cargo de Cuenta
'
'        adosolicitud.Recordset("subcta2") = "02"     'JQA JUN/2008 Cargo de Cuenta Otros
'        If Val(Txt_porcentaje.Text) > 0 Then
'            adosolicitud.Recordset("por_tiempo") = Val(Txt_porcentaje.Text)     '/ 100
'        Else
'            adosolicitud.Recordset("por_tiempo") = 0
'        End If
'        adosolicitud.Recordset("formulario") = "F01"
'        adosolicitud.Recordset("tipo_bien_Cta_doc") = "A"
'        If adosolicitud.Recordset("Lista_adjunta") = "S" Then
'            adosolicitud.Recordset("Lista_adjunta") = "S"
'        Else
'            adosolicitud.Recordset("Lista_adjunta") = "N"
'        End If
'        adosolicitud.Recordset("codigo_bien") = Dtccodbien.Text
'        adosolicitud.Recordset("nro_pagos") = 1     'IIf(IsNull(TxtCantPedi), 1, TxtCantPedi)
'        adosolicitud.Recordset("usr_usuario") = glusuario '"xxx"
'        adosolicitud.Recordset("fecha_registro") = Format(Date, "dd/mm/yyyy")
'        adosolicitud.Recordset("hora_registro") = Format(Time, "hh:mm:ss") '"16:00:00"
'        adosolicitud.Recordset("usuario_aprueba") = ""
'        adosolicitud.Recordset("hora_aprueba") = ""
''        adosolicitud.Recordset("AUnidad") = "-"
''        adosolicitud.Recordset("APlanilla") = 0
''        adosolicitud.Recordset("Planilla_depto") = "-"
''        adosolicitud.Recordset("Bco_codigo") = "-"
''        adosolicitud.Recordset("Ges_Gestion_ant") = "-"
''        adosolicitud.Recordset("APlanilla_existe") = "N"
'        adosolicitud.Recordset("estado_enviado") = "N"
'        adosolicitud.Recordset("estado_aprobado") = "N"
'      adosolicitud.Recordset.Update
'
'    db.CommitTrans
'    If adosolicitud.Recordset.RecordCount > 0 Then
'       marca1 = adosolicitud.Recordset.Bookmark
'       Call OptFilGral1_Click
'       'adosolicitud.Refresh
'       adosolicitud.Recordset.Move marca1 - 1
''       If swgrabar = 1 Then
''           adosolicitud.Refresh
''           adosolicitud.Recordset.MoveLast
''       End If
'    End If
''JQA JUN 2008
''    Dim rstdestino As New ADODB.Recordset
''    Set rstdestino = New ADODB.Recordset
''    If rstdestino.State = 1 Then rstdestino.Close
''    db.BeginTrans
''    If swgrabar = 1 Then
''      rstdestino.Open "select * from Ao_solicitud where formulario = '0'", db, adOpenDynamic, adLockOptimistic
''        Dim correlsolic As Integer
''        Dim rstao_solicitud_correl As New ADODB.Recordset
''        Set rstao_solicitud_correl = New ADODB.Recordset
''        If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
''        rstao_solicitud_correl.Open "select * from ao_solicitud_correl where codigo_unidad = '" & Trim(lblUni_codigo.Caption) & "' ", db, adOpenDynamic, adLockOptimistic
''        If rstao_solicitud_correl.RecordCount > 0 Then
''          If Not (rstao_solicitud_correl.BOF) Then rstao_solicitud_correl.MoveFirst
''          rstao_solicitud_correl("correl_solicitud") = rstao_solicitud_correl("correl_solicitud") + 1
''          correlsolic = rstao_solicitud_correl("correl_solicitud")
''          rstao_solicitud_correl.Update
''        Else
''            rstao_solicitud_correl.AddNew
''            'rstao_solicitud_correl("formulario") = "F03"
''            rstao_solicitud_correl("codigo_unidad") = Trim(lblUni_codigo.Caption)
''            rstao_solicitud_correl("correl_solicitud") = 1
''            correlsolic = rstao_solicitud_correl("correl_solicitud")
''            rstao_solicitud_correl.Update
''        End If
''        If rstao_solicitud_correl.State = 1 Then rstao_solicitud_correl.Close
''            rstdestino.AddNew
''            rstdestino("codigo_solicitud") = correlsolic
''        Else
''            rstdestino.Open "select * from Ao_solicitud where codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & lblUni_codigo & "'", db, adOpenDynamic, adLockOptimistic
''            correlsolic = adosolicitud.Recordset("codigo_solicitud")
''            '
''            'ges_gestion1 = adosolicitud.Recordset!ges_gestion_ant 'CStr(Year(DTPfechasol.Value))
''            'codigo_solicitud1 = adosolicitud.Recordset!codigo_solicitud_ant
''            'codigo_unidad1 = adosolicitud.Recordset!codigo_unidad_ant   'dtccoduni.Text
''        End If
'    'ALB MODI
''    If rstdestino.RecordCount > 0 Then
''      rstdestino("ges_gestion") = GlGestion         'CStr(Year(DTPfechasol.Value))
''      rstdestino("codigo_solicitud") = correlsolic
''      rstdestino("codigo_unidad") = Trim(lblUni_codigo)   'dtccoduni.Text
''      rstdestino("justificacion_solicitud") = txtjustifica.Text
''      rstdestino("CI") = dtccisol.Text
''      rstdestino("CI_aprueba") = Dtccibe.Text
''      rstdestino("fecha_solicitud") = DTPfechasol.Value
''      rstdestino("tr_adjuntos") = txtterref.Text
''      rstdestino("usr_usuario") = GlUsuario
''      rstdestino("fecha_registro") = Date
''      rstdestino!TipoF1 = "CC"      'DtCvalor1.BoundText  'jqa jun/2008 Cargo de Cuenta
''
''      rstdestino!subcta2 = "02"     'JQA JUN/2008 Cargo de Cuenta Otros
''        If Val(Txt_porcentaje.Text) > 0 Then
''            adosolicitud.Recordset("por_tiempo") = Val(Txt_porcentaje.Text)     '/ 100
''        Else
''            adosolicitud.Recordset("por_tiempo") = 0
''        End If
'''    Select Case cmbSubCta2.Text
'''      Case "Regulares" 'Cargos de Cuenta Regulares
'''        rstdestino!subcta2 = "01"
'''      Case "Otros" 'Cargos de Cuenta Otros
'''        rstdestino!subcta2 = "02"
'''      Case "PASE" 'Cargos de Cuenta PASE
'''        rstdestino!subcta2 = "03"
'''      Case Else
'''        rstdestino!subcta2 = "-"
'''    End Select
''
''    rstdestino!tipo_bien_Cta_doc = "A"      'JQA JUN/2008 Apertura de Cargo de Cuenta
''
'''    Select Case Lbltipo_bien_Cta_doc.Caption
'''      Case "APERTURA" 'Fondo Rotatorio Apertura
'''        rstdestino!tipo_bien_Cta_doc = "A"
'''        'ges_gestion1 = "-"
'''        'codigo_unidad1 = "-"
'''        'codigo_solicitud1 = 0
'''      Case "RENDICION" 'Fondo Rotatorio Rendición
'''        If Me.adosolicitud.Recordset!tipo_bien_Cta_doc = "A" Then
'''          rstdestino!tipo_bien_Cta_doc = "R"
'''        End If
'''      Case "CIERRE" 'Fondo Rotatorio Cierre
'''        rstdestino!tipo_bien_Cta_doc = "C"
'''      Case "CIERRE B.A." 'Fondo Rotatorio Cierre
'''        rstdestino!tipo_bien_Cta_doc = "CC"
'''    End Select
''
''          rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
''          rstdestino("usuario_aprueba") = ""
''          rstdestino("hora_aprueba") = ""
''          rstdestino("formulario") = "F01"
''          'ALB
''          'rstdestino("") = "-"
''          rstdestino("CONSULTOR_EMPRESA") = "-"
''          rstdestino("NACIONAL_EXTRANJERO") = "-"
''          rstdestino("FUNCION_ACTIVIDAD") = "-"
''          rstdestino("duracion_estimada_numero") = 0
''          rstdestino("duracion_estimada_tiempo") = "-"
''          rstdestino("esparaRH") = "-"
''          rstdestino("impuestos") = "-"
''          rstdestino("fecha_estimada_inicio") = Format(Date, "dd/mm/yyyy")
''          rstdestino("observaciones") = "-"
''          rstdestino("codigo_bien") = "-"
''          rstdestino("caracteristicas") = "-"
''          rstdestino("APROBADO") = 0
''          rstdestino("estatus") = "N"
''          rstdestino("pas_viat") = "-"
''          rstdestino("TRAMO1") = "-"
''          rstdestino("TRAMO2") = "-"
''          rstdestino("TRAMO3") = "-"
''          rstdestino("TRAMO4") = "-"
''          rstdestino("TRAMO5") = "-"
''          rstdestino("ESTADO_APROBACION") = "N"
''          rstdestino("APLANILLA") = 0
''          rstdestino("NRO_PAGOS") = 0
''          rstdestino("PLANILLA_DEPTO") = "-"
''          rstdestino("BCO_CODIGO") = "-"
''          rstdestino("AUnidad") = "-"
''          rstdestino("APlanilla_existe") = "N"
''          rstdestino("fecha_registro") = Format(Date, "dd/mm/yyyy")
''          rstdestino("fecha_reCEPCIÓN") = Format(Date, "dd/mm/yyyy")
''          rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
''          rstdestino("lista_adjunta") = "N"
''          rstdestino("Observaciones") = ""
''          rstdestino!ges_gestion_ant = "-" 'ges_gestion1 'CStr(Year(DTPfechasol.Value))
''          rstdestino!codigo_solicitud_ant = 0 ' codigo_solicitud1
''          rstdestino!codigo_unidad_ant = "-" 'codigo_unidad1
''          rstdestino("es_planilla") = "-"
''      rstdestino.Update
''    End If
''    db.CommitTrans
''    If rstdestino.State = 1 Then rstdestino.Close
'''        marca1 = IIf(adosolicitud.Recordset.RecordCount > 0, adosolicitud.Recordset.BookMark, 2)
'''        adosolicitud.Refresh
'''        adosolicitud.Recordset.Move marca1 - 1
''adosolicitud.Refresh
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'End Sub
'
'Private Sub Text1_LostFocus()
'    If CDbl(Text1.Text) > 0 Then
'        Txt_porcentaje.Text = CDbl(Text1.Text) / 100
'    Else
'        Txt_porcentaje.Text = 0
'    End If
'End Sub
'
'Private Sub txtjustifica_KeyPress(KeyAscii As Integer)
'   'convertir a mayusculas
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    'solo numeros a numero
''      If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
''
''      Else
''        KeyAscii = Asc(UCase(Chr(0)))
''      End If
'
'End Sub
'
'Private Sub TxtMonto_bolivianos_contra_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'
'End Sub
'
''Private Sub TxtMonto_bolivianos_contra_KeyUp(KeyCode As Integer, Shift As Integer)
''  If Len(TxtTipo_cambio.Text) > 0 Then
''    If (Len(Trim(TxtMonto_bolivianos_contra.Text)) > 0) Then
''       Txtmonto_dolares_contra.Text = IIf(TxtMonto_bolivianos_contra.Text > 0, TxtMonto_bolivianos_contra.Text / TxtTipo_cambio, 0)
''    Else
''       Txtmonto_dolares_contra.Text = 0
''    End If
''  End If
''
''End Sub
''
''Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'''solo numeros y , .
''    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
''
''    Else
''      KeyAscii = Asc(UCase(Chr(0)))
''    End If
''End Sub
''
''Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
''  If Len(TxtTipo_cambio.Text) > 0 Then
''    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
''       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, TxtMonto_bolivianos.Text / TxtTipo_cambio, 0)
''    Else
''       Txtmonto_dolares.Text = 0
''    End If
''  End If
''
''End Sub
''
''Private Sub Txtmonto_dolares_contra_KeyPress(KeyAscii As Integer)
''    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
''    Else
''      KeyAscii = Asc(UCase(Chr(0)))
''    End If
''
''End Sub
''
''Private Sub Txtmonto_dolares_contra_KeyUp(KeyCode As Integer, Shift As Integer)
''  If Len(TxtTipo_cambio.Text) > 0 Then
''    If Len(Trim(Txtmonto_dolares_contra.Text)) > 0 Then
''      TxtMonto_bolivianos_contra.Text = IIf(Txtmonto_dolares_contra.Text > 0, Txtmonto_dolares_contra * TxtTipo_cambio, 0)
''    Else
''      TxtMonto_bolivianos_contra.Text = 0
''    End If
''  End If
''
''End Sub
''
''Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
''    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
''    Else
''      KeyAscii = Asc(UCase(Chr(0)))
''    End If
''
''End Sub
''
''Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
''  If Len(TxtTipo_cambio.Text) > 0 Then
''    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
''      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Txtmonto_dolares * TxtTipo_cambio, 0)
''    Else
''      TxtMonto_bolivianos.Text = 0
''    End If
''  End If
''
''End Sub
'
'Private Sub txtterref_KeyPress(KeyAscii As Integer)
'    If KeyAscii < 58 And KeyAscii > 47 Then
'        KeyAscii = Asc(UCase(Chr(0)))
'    Else
'        If UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "N" Or KeyAscii = 8 Then
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Else
'            KeyAscii = Asc(UCase(Chr(0)))
'            MsgBox "Debe escribir solo 'N' o 'S'", vbOKOnly, "Error..."
'        End If
'    End If
'End Sub
'
'Private Sub TxtTipo_cambio_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'    Else
'      KeyAscii = Asc(UCase(Chr(0)))
'    End If
'
'End Sub
'Private Sub OptFilGral1_Click()
'  '===== Proceso para filtrado general de datos(registros aprobados)
'  Set rstAo_solicitud = New adodb.Recordset
'  'queryinicial = "select * from Ao_solicitud where formulario = 'F01' and estatus <> 'A' and estatus <> 'S' AND usr_usuario = '" & GlUsuario & "' "
'  queryinicial = "select * from Ao_solicitud where formulario = 'F01' and estado_enviado = 'N'"
'  If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'  rstAo_solicitud.Open queryinicial & " order by codigo_unidad , codigo_solicitud", db, adOpenKeyset, adLockOptimistic
'  rstAo_solicitud.Requery
'  Set adosolicitud.Recordset = rstAo_solicitud
'  If rstAo_solicitud.RecordCount > 0 Then
'    Frame10.Enabled = True
'    Frame10.Visible = True
'    CmdImpCabeza.Enabled = True
'    CmdBusCabeza.Enabled = True
'  Else
'    Frame10.Enabled = False
'    Frame10.Visible = False
'    CmdImpCabeza.Enabled = False
'    CmdBusCabeza.Enabled = False
'  End If
'End Sub
'
'Private Sub OptFilGral2_Click()
'  '===== Proceso para filtrado general de datos (todos los registros )Ç
'  'queryinicial = "select * from Ao_solicitud where formulario = 'F01' AND usr_usuario = '" & GlUsuario & "' "
'  queryinicial = "select * from Ao_solicitud where formulario = 'F01' "
'  If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'  rstAo_solicitud.Open queryinicial & " order by codigo_unidad , codigo_solicitud", db, adOpenKeyset, adLockOptimistic
'  rstAo_solicitud.Requery
'  Set adosolicitud.Recordset = rstAo_solicitud
'  If rstAo_solicitud.RecordCount > 0 Then
'    Frame10.Enabled = True
'    Frame10.Visible = True
'    CmdImpCabeza.Enabled = True
'    CmdBusCabeza.Enabled = True
'  Else
'    Frame10.Enabled = False
'    Frame10.Visible = False
'    CmdImpCabeza.Enabled = False
'    CmdBusCabeza.Enabled = False
'  End If
'End Sub
'
'
'Private Sub fbuscaunidad()
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'  'rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  rstFc_unidad_ejecutora.Open "select * from Fc_unidad_ejecutora where uni_codigo = '" & Trim(adopuestosol.Recordset("codigo_unidad")) & "'", db, adOpenKeyset, adLockReadOnly
'  If rstFc_unidad_ejecutora.RecordCount > 0 Then
'    LblUni_descripcion_larga.Caption = rstFc_unidad_ejecutora("Uni_descripcion_larga")
'  Else
'    LblUni_descripcion_larga.Caption = ""
'  End If
'  If rstFc_unidad_ejecutora.State = 1 Then rstFc_unidad_ejecutora.Close
'End Sub
'
'
'
'Private Sub cerea()
'  txtnrosol = ""
'''  Optpasvia1.Value = False
''  Optpasvia2.Value = False
''  dtccodpoa.Text = ""
''  dtcdespoa.Text = dtccodpoa.BoundText
'  dtccisol.Text = ""
'  'Dtcpaternosol.Text = dtccisol.BoundText
'  Dtcpaternosol.Text = ""
'  dtcmaternosol.Text = ""
'  dtcnombresol.Text = ""
'  Dtccibe.Text = ""
'  Dtcpaternobe.Text = ""
''  Dtcpaternobe.Text = Dtccibe.BoundText
''  Dtcmaternobe.Text = ""
''  Dtcnombrebe.Text = ""
''  DtCcodigo_beneficiario = ""
''  DtCdenominacion_beneficiario = DtCcodigo_beneficiario.BoundText
'  txtjustifica.Text = ""
'  txtterref.Text = ""
''  TxtDurac_tiempo.Text = ""
''  DtCDenominacion_moneda.Text = ""
''  TxtTipo_cambio.Text = GlTipoCambioOficial
''  TxtMonto_bolivianos.Text = 0
''  Txtmonto_dolares.Text = 0
''  DtCOrg_descripcion.Text = ""
''  TxtMonto_bolivianos_contra.Text = 0
''  Txtmonto_dolares_contra.Text = 0
'End Sub
'
'
'Private Sub DtCvalor1_LostFocus()
'  If DtCvalor1.BoundText = "CC" Then
'    Label3.Visible = True
'    cmbSubCta2.Visible = True
'  Else
'    Label3.Visible = False
'    cmbSubCta2.Visible = False
'    cmbSubCta2.Text = ""
'  End If
'End Sub
'
'Private Sub APrueba2()
'Dim rsAcum As New adodb.Recordset
'  Dim Acum As Double
'  Dim Acumlimite As Double
'  Set rstdestino = New adodb.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenKeyset, adLockOptimistic
'  If rstdestino.RecordCount < 1 Then
'    MsgBox "No puede aprobar sin Detalle de Registro.", vbCritical + vbOKOnly, "Error al aprobar..."
'    Exit Sub
'  End If
'  Dim swver_monto As Integer
'  swver_monto = 1
'  If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "R" Or Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Then
'    '==== ini convenios ====
'    Dim rstAo_solicitud_detalle_ant As New adodb.Recordset
'    Set rstAo_solicitud_detalle_ant = New adodb.Recordset
'    Dim swconv As Integer
'    swconv = 1
'    conv2 = " "
'    conv1 = " "
'    rstao_solicitud_detalle.MoveFirst
'    While Not rstao_solicitud_detalle.EOF
'      Call fBuscaConvenio(rstao_solicitud_detalle!codigo_poa)
'      If rstAo_solicitud_detalle_ant.State = 1 Then rstAo_solicitud_detalle_ant.Close
'      rstAo_solicitud_detalle_ant.Open "select * from ao_solicitud_detalle where ges_gestion = '" & FrmF01.adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & FrmF01.adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & FrmF01.adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockReadOnly
'      If rstAo_solicitud_detalle_ant.RecordCount > 0 Then
'        While Not rstAo_solicitud_detalle_ant.EOF
'          conv1 = rstAo_solicitud_detalle_ant!codigo_convenio
'          If rstAo_solicitud_detalle_ant!codigo_convenio <> conv2 Then
'            MsgBox "No puede aprobar una RENDICION o un CIERRE " & vbCrLf & "con un convenio diferente al de la APERTURA." & vbCrLf & vbCrLf & _
'                   "       Convenio Apertura  : " & rstAo_solicitud_detalle_ant!codigo_convenio & vbCrLf & _
'                   "       Convenio Solicitud  : " & conv2 & " (poa: " & rstao_solicitud_detalle!codigo_poa & ") " & vbCrLf & _
'            vbCrLf & "          Por favor corrija los codigos POA", vbCritical + vbOKOnly, "Error al aprobar..."
'            swconv = 0
'          End If
'          rstAo_solicitud_detalle_ant.MoveNext
'        Wend
'      Else
'        MsgBox "Error en el convenio de la APERTURA.", vbCritical + vbOKOnly, "Error al aprobar..."
'        Exit Sub
'      End If
'      rstao_solicitud_detalle.MoveNext
'    Wend
'    '==== fin convenios ====
'    If swconv = 0 Then Exit Sub
'    '==== ini CTA ====
''    Dim rstfc_convenio As New ADODB.Recordset
''    Set rstfc_convenio = New ADODB.Recordset
''    If rstfc_convenio.State = 1 Then rstfc_convenio.Close
''    rstfc_convenio.Open "select * from fc_convenioS where codigo_convenio = '" & Conv1 & "' ", db, adOpenKeyset, adLockReadOnly
''    If rstfc_convenio.RecordCount > 0 Then
''      cta1 = rstfc_convenio!cta_codigo
''    Else
''      cta1 = " "
''    End If
''    If rstfc_convenio.State = 1 Then rstfc_convenio.Close
''    If cta1 <> Me.DtCBco_codigo.Text Then
''      MsgBox "La cuenta bancaria de la solicitud " & vbCrLf & "debe pertenecer al Convenio de Apertura" & vbCrLf & vbCrLf & _
''             "     Cuenta de Apertura : " & cta1 & vbCrLf & _
''             "     Cuenta de Solicitud : " & Me.DtCBco_codigo.Text, vbCritical + vbOKOnly, "Error en las Cuentas Bancarias..."
''      Exit Sub
''    End If
'    '==== fin CTA ====
'
'    '==== INI MONTOS ====
'    Set rsAcum = New adodb.Recordset
'    If rsAcum.State = 1 Then rsAcum.Close
'    rsAcum.Open "select sum (monto_bolivianos) + sum (monto_bolivianos_CONTRA)as acum from ao_solicitud_detalle where ges_gestion = '" & FrmF01.adosolicitud.Recordset!ges_gestion & "' and codigo_unidad = '" & FrmF01.adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & FrmF01.adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    If rsAcum.RecordCount > 0 Then
'      Acum = IIf(IsNull(rsAcum!Acum), 0, rsAcum!Acum)
'    End If
'    If rsAcum.State = 1 Then rsAcum.Close
'    Set rsAcum = New adodb.Recordset
'    If rsAcum.State = 1 Then rsAcum.Close
'    rsAcum.Open "select sum (monto_bolivianos) + sum (monto_bolivianos_CONTRA)as acum from ao_solicitud_detalle where ges_gestion = '" & FrmF01.adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & FrmF01.adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & FrmF01.adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockReadOnly
'    If rsAcum.RecordCount > 0 Then
'      Acumlimite = IIf(IsNull(rsAcum!Acum), 0, rsAcum!Acum)
'    End If
'    If rsAcum.State = 1 Then rsAcum.Close
'    If Acum + Me.adosolicitud.Recordset!Nro_pagos > Acumlimite Then
'      MsgBox "No puede aprobar una RENDICION o un CIERRE" & vbCrLf & "con monto mayor al de la APERTURA." & vbCrLf & _
'             "     Monto de Apertura : " & Acumlimite & vbCrLf & _
'             "     Monto Solicitud      : " & Acum + Me.adosolicitud.Recordset!Nro_pagos & vbCrLf & _
'      vbCrLf & vbCrLf & "          Por favor corrija los MONTOS.", vbCritical + vbOKOnly, "Error al aprobar..."
'      Exit Sub
'    End If
'    If Trim(adosolicitud.Recordset!tipo_bien_Cta_doc) = "C" Then
'      If (Acum + Me.adosolicitud.Recordset!Nro_pagos) < Acumlimite Then
'        MsgBox "No puede aprobar un CIERRE con monto menor al de APERTURA." & vbCrLf & _
'        "     Monto APERTURA : " & Acumlimite & vbCrLf & "     Monto CIERRE        : " & Acum + Me.adosolicitud.Recordset!Nro_pagos & vbCrLf & vbCrLf & _
'        "          Por favor corrija los MONTOS.", vbCritical + vbOKOnly, "Error al aprobar..."
'        Exit Sub
'      End If
'    End If
'  End If
''  swver_monto = verifica_montos(Me.adosolicitud.Recordset!codigo_unidad, Me.adosolicitud.Recordset!codigo_solicitud)
'  If swver_monto = 0 Then
'    MsgBox "El registro tiene problemas en los montos, Por favor Verifique e intente aprobarlo luego. Gracias", vbCritical + vbOKOnly, "Error en los montos"
'    Exit Sub
'  End If
''==== MONTOS ====
'  sino = MsgBox("Esta seguro de Aprobar el registro?", vbYesNo, "Confirmando")
'  If sino = vbYes Then
'    db.BeginTrans
'    Set rstdestino = New adodb.Recordset
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from Ao_solicitud where ges_gestion = '" & adosolicitud.Recordset("ges_gestion") & "' and formulario = '" & adosolicitud.Recordset("formulario") & "' and codigo_solicitud = " & adosolicitud.Recordset("codigo_solicitud") & " and codigo_unidad = '" & adosolicitud.Recordset("codigo_unidad") & "'", db, adOpenDynamic, adLockOptimistic
'    If Not rstdestino.BOF Then rstdestino.MoveFirst
'    If Not rstdestino.BOF And Not rstdestino.EOF Then
'      rstdestino("aprobado") = 1
'      rstdestino("estado_aprobacion") = "S"
'      rstdestino.Update
'      If rstdestino!tipo_bien_Cta_doc = "C" Then
'        Set rsAcum = New adodb.Recordset
''        rsAcum.CancelUpdate
'        If rsAcum.State = 1 Then rsAcum.Close
'        rsAcum.Open "select * from ao_solicitud where ges_gestion = '" & FrmF01.adosolicitud.Recordset!ges_gestion_ant & "' and codigo_unidad = '" & FrmF01.adosolicitud.Recordset!codigo_unidad_ant & "' and codigo_solicitud = " & FrmF01.adosolicitud.Recordset!codigo_solicitud_ant, db, adOpenKeyset, adLockOptimistic
'        If rsAcum.RecordCount > 0 Then
'          rsAcum!codigo_unidad_ant = "X"
'          rsAcum.Update
'        End If
'        If rsAcum.State = 1 Then rsAcum.Close
'      End If
'    End If
'    If rstdestino.State = 1 Then rstdestino.Close
'    db.CommitTrans
'    marca1 = adosolicitud.Recordset.Bookmark
'    Set adosolicitud.Recordset = rstAo_solicitud
'    adosolicitud.Refresh
'    adosolicitud.Recordset.Move marca1 - 1
''    rstAo_solicitud.Requery
'  End If
'End Sub
'
'Private Sub fBuscaConvenio(Poa)
'  Dim rst_fc_relacionador_poa_ppto As New adodb.Recordset
'  Set rst_fc_relacionador_poa_ppto = New adodb.Recordset
'  If rst_fc_relacionador_poa_ppto.State = 1 Then rst_fc_relacionador_poa_ppto.Close
'  rst_fc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & Poa & "'", db, adOpenKeyset, adLockReadOnly
'  If rst_fc_relacionador_poa_ppto.RecordCount > 0 Then
'    conv2 = rst_fc_relacionador_poa_ppto!codigo_convenio
'  Else
'    conv2 = "Err"
'  End If
'  If rst_fc_relacionador_poa_ppto.State = 1 Then rst_fc_relacionador_poa_ppto.Close
'End Sub
'
'Private Function fbusmuni(cod)
'  Dim rstfc_unidad_educativa As New adodb.Recordset
'  Set rstfc_unidad_educativa = New adodb.Recordset
'  If rstfc_unidad_educativa.State = 1 Then rstfc_unidad_educativa.Close
'  rstfc_unidad_educativa.Open "select * from fc_unidad_educativa where codigo = '" & cod & "'", db, adOpenKeyset, adLockReadOnly
'  If rstfc_unidad_educativa.RecordCount > 0 Then
'    fbusmuni = rstfc_unidad_educativa!denominacion
'  Else
'    fbusmuni = ""
'  End If
'  If rstfc_unidad_educativa.State = 1 Then rstfc_unidad_educativa.Close
'End Function
'
'Private Sub val_presupF01(adoorigen, GlNombFor)
'
'  'If (GlNombFor <> "F02") And (GlNombFor <> "F06") Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
'  If (GlNombFor = "F01" And Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
'    Set rstao_solicitud_detalle = New adodb.Recordset
'    If rstao_solicitud_detalle.State = 1 Then rstao_solicitud_detalle.Close
'    rstao_solicitud_detalle.Open "select * from ao_solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    If rstao_solicitud_detalle.RecordCount > 0 Then
'      rectot = rstao_solicitud_detalle.RecordCount
'      Fte_contraparte1 = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
'      Org_Contraparte1 = rstao_solicitud_detalle!org_codigo_contra
'      Dim v_EstPoa(50, 14)
'    End If
'    If Not (rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
'    For i = 1 To rstao_solicitud_detalle.RecordCount            ' primer i
'      Set rstfc_relacionador_poa_ppto = New adodb.Recordset
'      If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
'      rstfc_relacionador_poa_ppto.Open "select * from fc_relacionador_poa_ppto where codigo_poa = '" & rstao_solicitud_detalle!codigo_poa & "'", db, adOpenKeyset, adLockReadOnly
'      If rstfc_relacionador_poa_ppto.RecordCount > 0 Then
'        'aqui se puede definir porcentaje
'        v_EstPoa(i, 1) = rstao_solicitud_detalle!codigo_poa
'        v_EstPoa(i, 2) = rstfc_relacionador_poa_ppto!da          'Dirección Administrativa JGCA 10/08/2007
'        v_EstPoa(i, 3) = rstfc_relacionador_poa_ppto!par_codigo             'Par_Codigo1
'        'v_EstPoa(i, 4) = fBuscaFte(rstfc_relacionador_poa_ppto!org_codigo) 'fte_codigo1
'        v_EstPoa(i, 4) = rstfc_relacionador_poa_ppto!fte_codigo             'fte_codigo1    JQA JUL-2005
'        v_EstPoa(i, 5) = rstfc_relacionador_poa_ppto!org_codigo             'Org_Codigo1
'        v_EstPoa(i, 6) = rstfc_relacionador_poa_ppto!pro_programa           'pro_Programa1
'        'v_EstPoa(i, 7) = rstfc_relacionador_poa_ppto!pro_subprograma       'Pro_SubPrograma1
'        v_EstPoa(i, 8) = rstfc_relacionador_poa_ppto!pro_proyecto           'Pro_Proyecto1
'        v_EstPoa(i, 9) = rstfc_relacionador_poa_ppto!pro_actividad          'Pro_Actividad1
'        v_EstPoa(i, 10) = rstfc_relacionador_poa_ppto!codigo_unidad            'uni_codigo1
'        'v_EstPoa(i, 11) = IIf(IsNull(rstfc_relacionador_poa_ppto!Categoria), "xx", rstfc_relacionador_poa_ppto!Categoria) 'codigo_categoria1   JQA JUL-2005
'        v_EstPoa(i, 11) = IIf(IsNull(rstfc_relacionador_poa_ppto!categoria), rstfc_relacionador_poa_ppto!codigo_categoria, rstfc_relacionador_poa_ppto!categoria) 'codigo_categoria1    JQA JUL-2005
'        v_EstPoa(i, 12) = rstfc_relacionador_poa_ppto!codigo_convenio       'codigo_convenio1
'        por_fte_ext1 = rstfc_relacionador_poa_ppto!por_ext
'        por_fte_nal1 = rstfc_relacionador_poa_ppto!por_nal
'        If rstfc_relacionador_poa_ppto!por_ext = 100 Then                   'JQA JUL-2005
'            v_EstPoa(i, 13) = rstfc_relacionador_poa_ppto!fte_codigo        'fte_codigo2   JQA JUL-2005
'            v_EstPoa(i, 14) = rstfc_relacionador_poa_ppto!org_codigo        'Org_Codigo2   JQA JUL-2005
'            cat_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!categoria), rstfc_relacionador_poa_ppto!codigo_categoria, rstfc_relacionador_poa_ppto!categoria) 'codigo_categoria1    JQA JUL-2005
'            conv_nal = rstfc_relacionador_poa_ppto!codigo_convenio          'codigo_convenio1
'            tot_form = 1
'        Else
'            v_EstPoa(i, 13) = rstfc_relacionador_poa_ppto!fte_codigo        'fte_codigo2   JQA NOV-2008
'            v_EstPoa(i, 14) = rstfc_relacionador_poa_ppto!org_codigo        'Org_Codigo2   JQA NOV-2008
'            cat_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal), "FIN_PROPIO", rstfc_relacionador_poa_ppto!CODIGO_Categoria_nal)  'codigo_categoria2    JQA JUL-2005
'            conv_nal = IIf(IsNull(rstfc_relacionador_poa_ppto!codigo_convenio_nal), "FIN_PROPIO", rstfc_relacionador_poa_ppto!codigo_convenio_nal)      'codigo_convenio2
'            tot_form = 2
'        End If
''        If rstao_solicitud_detalle!org_codigo_contra = "" Or rstao_solicitud_detalle!org_codigo_contra = "-" Then
''          v_EstPoa(i, 13) = "10"
''          v_EstPoa(i, 14) = "111"
''        Else
''          v_EstPoa(i, 13) = fBuscaFte(rstao_solicitud_detalle!org_codigo_contra)
''          v_EstPoa(i, 14) = rstao_solicitud_detalle!org_codigo_contra
''        End If
'        If rstfc_relacionador_poa_ppto.State = 1 Then rstfc_relacionador_poa_ppto.Close
'        Dim rstfo_formulacion_gasto As New adodb.Recordset
'        Set rstfo_formulacion_gasto = New adodb.Recordset
'        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
'        'rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & pro_Programa1 & "' and pro_subprograma='" & Pro_SubPrograma1 & "' and pro_proyecto='" & Pro_Proyecto1 & "' and pro_actividad='" & Pro_Actividad1 & "' and par_codigo='" & Par_Codigo1 & "' and org_codigo= '" & Org_Codigo1 & "'", db, adOpenKeyset, adLockOptimistic
'        rstfo_formulacion_gasto.Open "select * from fo_formulacion_gasto where pro_programa='" & v_EstPoa(i, 6) & "' and pro_proyecto='" & v_EstPoa(i, 8) & "' and pro_actividad='" & v_EstPoa(i, 9) & "' and par_codigo='" & v_EstPoa(i, 3) & "' and org_codigo= '" & v_EstPoa(i, 5) & "'", db, adOpenKeyset, adLockOptimistic
'        If Not (rstfo_formulacion_gasto.EOF) Then
'          If (rstfo_formulacion_gasto!FGS_VIGENTE - rstfo_formulacion_gasto!FGS_compromiso < rstao_solicitud_detalle!monto_Bolivianos) Then  'adoorigen         'adoorigen.adosolicitud.Recordset!monto_dolares ) Then
'            'JQA 07/12/01
''            swSubir = "No existe Presup"
''            MsgBox "NO EXISTE Presupuesto para dar curso a la Solicitud ...", vbOKOnly, "ERROR"
''            swpresup = 0
''            Exit Sub
'            'JQA 07/12/01
'            swpresup = 1    'Borrar despues de habilitar JQA
'          Else
'            'JQA 07/12/01
'            'rstfo_formulacion_gasto!0  = rstfo_formulacion_gasto!fgs_precompromiso  + rstao_solicitud_detalle!monto_bolivianos
'            'rstfo_formulacion_gasto.Update
'            'JQA 07/12/01
'            swpresup = 1
'            swSubir = "SI correcto"
'          End If
'          If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
'            swpresup = 1
'          Else
'            'JQA 07/12/01
''          MsgBox "NO EXISTE Estructura presupuestaria...", vbOKOnly, "ERROR ..."
''          swSubir = "NO Error Estruc.Ppto"
''          swpresup = 0
''          Exit Sub
'            'JQA 07/12/01
'            swpresup = 1    'Borrar despues de habilitar JQA
'          End If
'      Else
'        MsgBox "NO Existe POA ... ", vbOKOnly, "ERROR ..."
'        swSubir = "No existe POA"
'        swpresup = 0
'        Exit Sub
'      End If
''          Else
''            swpresup = 1
''          End If
'      rstao_solicitud_detalle.MoveNext
'    Next            'fin del primer i
'
'    If swpresup = 1 Then        'ini swpresup
'      If GlNombFor <> "F02" Or (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
'        If (rstao_solicitud_detalle.RecordCount > 0) And (Not rstao_solicitud_detalle.BOF) Then rstao_solicitud_detalle.MoveFirst
'        Set rstao_solicitud_recibido = New adodb.Recordset
'        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'        rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'        db.BeginTrans
'        'por_fte_ext
'        'por_fte_nal
'        For j = 1 To rstao_solicitud_detalle.RecordCount        ' del j
'          'j = 2
'          v_por_fte(1, 1) = por_fte_ext1
'          v_por_fte(1, 2) = v_EstPoa(j, 4) 'fte_codigo1
'          v_por_fte(1, 3) = v_EstPoa(j, 5) 'Org_Codigo1
'
'          v_por_fte(2, 1) = por_fte_nal1
'          v_por_fte(2, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
'          v_por_fte(2, 3) = v_EstPoa(j, 14) 'Org_Contraparte1
'
'          v_por_fte(3, 1) = por_fte_nal1
'          v_por_fte(3, 2) = v_EstPoa(j, 13) 'Fte_contraparte1
'          v_por_fte(3, 3) = v_EstPoa(j, 14) 'Org_Contraparte1
'
'          Dim SwEsBase As Integer
'          Dim ValEsBase As Double
''          ValEsBase = v_por_fte(1, 1)
''          For I = 1 To tot_form
''            If v_por_fte(I, 1) > ValEsBase Then
''              SwEsBase = I
''              ValEsBase = v_por_fte(I, 1)
''            End If
''          Next
''AQUI UN SOLO FINANCIADOR
''la variable "j" distribuye al Preventivo y al Comprometido cada reg de ao_solicitud_detalle
'        'k = 1
''        While (k <= 2)
''        begin
'         prev_dev = 2       ' 1 p/pagos_espera y 2 p/pagos
'         'prev_dev = 1       ' 1 solo p/pagos_espera
'         For k = 1 To prev_dev
'            '        Print rstpagos!monto_bolivianos
'          For i = 1 To tot_form         'dos (segundo i)
'            If k = 1 Then
'                Set rstpagos = New adodb.Recordset
'                If rstpagos.State = 1 Then rstpagos.Close
'
'                rstpagos.Open "select * from pagos_espera where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
'                'db.Execute " EXEC edGeneraCodigoPagoEspera @org_codigo_ida, @GCODIGO_PAGO OUT "
'                'Organismo1 = v_por_fte(i, 3)
'                'db.Execute "EXEC edGeneraCodigoPagoEspera Organismo1, GCODIGO_PAGO  "
'                'codigo_pago1 = GCODIGO_PAGO
'                Set rscorrelativo = New adodb.Recordset
'                If rscorrelativo.State = 1 Then rscorrelativo.Close
'                rscorrelativo.Open "select * from fc_correlativos_espera", db, adOpenKeyset, adLockOptimistic
'            Else
'                'i = i - 1
'                Set rstpagos = New adodb.Recordset
'                If rstpagos.State = 1 Then rstpagos.Close
'                'org_codigo_ida = v_por_fte(i, 3)
'                rstpagos.Open "select * from pagos where codigo_pago = ''", db, adOpenKeyset, adLockOptimistic
'                'db.Execute " EXEC edGeneraCodigoPago @org_codigo_ida, @GCODIGO_PAGO OUT "
'                'db.Execute " EXEC edGeneraCodigoPago v_por_fte(i, 3), @GCODIGO_PAGO OUT "
'                'codigo_pago1 = GCODIGO_PAGO
'                Set rscorrelativo = New adodb.Recordset
'                If rscorrelativo.State = 1 Then rscorrelativo.Close
'                rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
'            End If
'            rstpagos.AddNew
'            '==== ini generación de correlativo ====
'            'If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then 'Or GlNombFor = "F06"
'            If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
'            'If k = 1 Then
'                If v_por_fte(i, 3) = "111" Then  'TGN
'                  If Not IsNull(rscorrelativo!correl_org111) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                    rscorrelativo!correl_org111 = CDbl(CDbl(rscorrelativo!correl_org111) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "112" Then 'TGNP
'                  If Not IsNull(rscorrelativo!correl_org112) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                    rscorrelativo!correl_org112 = CDbl(CDbl(rscorrelativo!correl_org112) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "113" Then
'                  If Not IsNull(rscorrelativo!Correl_Org113) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!Correl_Org113) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!Correl_Org113) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!Correl_Org113) + 1)
'                    rscorrelativo!Correl_Org113 = CDbl(CDbl(rscorrelativo!Correl_Org113) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "114" Then    'If Org_Codigo1 = "114" Then 'RECON
'                  If Not IsNull(rscorrelativo!correl_org114) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                    rscorrelativo!correl_org114 = CDbl(CDbl(rscorrelativo!correl_org114) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "344" Then 'UNICEF
'    '            codigo_pago1 = 1
'                  If Not IsNull(rscorrelativo!correl_org344) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                    rscorrelativo!correl_org344 = CDbl(CDbl(rscorrelativo!correl_org344) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "381" Then  'FAD
'                  If Not IsNull(rscorrelativo!correl_org381) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org381) + 1)
'                    rscorrelativo!correl_org381 = Val(Val(rscorrelativo!correl_org381) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "411" Then  'BID
'                  If Not IsNull(rscorrelativo!correl_org411) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                    rscorrelativo!correl_org411 = CDbl(CDbl(rscorrelativo!correl_org411) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "415" Then  'IDA
'                  If Not IsNull(rscorrelativo!correl_org415) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                    rscorrelativo!correl_org415 = CDbl(CDbl(rscorrelativo!correl_org415) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "516" Then  'KFW
'                  If Not IsNull(rscorrelativo!correl_org516) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                    rscorrelativo!correl_org516 = CDbl(CDbl(rscorrelativo!correl_org516) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "541" Then  'ALEM
'                  If Not IsNull(rscorrelativo!correl_org541) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                    rscorrelativo!correl_org541 = CDbl(CDbl(rscorrelativo!correl_org541) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "551" Then  'DIN
'                  If Not IsNull(rscorrelativo!correl_org551) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                    rscorrelativo!correl_org551 = CDbl(CDbl(rscorrelativo!correl_org551) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "556" Then  'HOL
'                  If Not IsNull(rscorrelativo!correl_org556) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                    rscorrelativo!correl_org556 = CDbl(CDbl(rscorrelativo!correl_org556) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "565" Then  'SUE
'                  If Not IsNull(rscorrelativo!correl_org565) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                    rscorrelativo!correl_org565 = CDbl(CDbl(rscorrelativo!correl_org565) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "999" Then  'S/N
'                  If Not IsNull(rscorrelativo!correl_org999) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                    rscorrelativo!correl_org999 = CDbl(CDbl(rscorrelativo!correl_org999) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'                'alb aqui 2004
'
'                If v_por_fte(i, 3) = "720" Then
'                  If Not IsNull(rscorrelativo!correl_org720) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org720) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org720) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org720) + 1)
'                    rscorrelativo!correl_org720 = CDbl(CDbl(rscorrelativo!correl_org720) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "000" Then          ' Org. No Clasificado
'                  If Not IsNull(rscorrelativo!correl_org000) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org000) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org000) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org000) + 1)
'                    rscorrelativo!correl_org000 = CDbl(CDbl(rscorrelativo!correl_org000) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "514" Then
'                  If Not IsNull(rscorrelativo!correl_org514) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                    rscorrelativo!correl_org514 = CDbl(CDbl(rscorrelativo!correl_org514) + 1)
'                    rscorrelativo.Update
'                  Else
'                    rscorrelativo!correl_org514 = 0
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "517" Then  'GTZ
'                  If Not IsNull(rscorrelativo!correl_org517) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                    rscorrelativo!correl_org517 = CDbl(CDbl(rscorrelativo!correl_org517) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "528" Then  'AECI
'                  If Not IsNull(rscorrelativo!correl_org528) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                    rscorrelativo!correl_org528 = CDbl(CDbl(rscorrelativo!correl_org528) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "518" Then  'JICA
'                  If Not IsNull(rscorrelativo!correl_org518) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org518) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org518) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org518) + 1)
'                    rscorrelativo!correl_org518 = CDbl(CDbl(rscorrelativo!correl_org518) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "520" Then  'SUECIA
'                  If Not IsNull(rscorrelativo!correl_org520) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org520) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org520) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org520) + 1)
'                    rscorrelativo!correl_org520 = CDbl(CDbl(rscorrelativo!correl_org520) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "210" Then  'MUNICIPIOS
'                  If Not IsNull(rscorrelativo!correl_org210) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org210) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org210) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org210) + 1)
'                    rscorrelativo!correl_org210 = CDbl(CDbl(rscorrelativo!correl_org210) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "729" Then  'Otros Organismos Financiadores Externos
'                  If Not IsNull(rscorrelativo!correl_org729) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org729) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org729) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org729) + 1)
'                    rscorrelativo!correl_org729 = CDbl(CDbl(rscorrelativo!correl_org729) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "345" Then  'UNFPA
'                  If Not IsNull(rscorrelativo!correl_org345) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org345) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org345) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org345) + 1)
'                    rscorrelativo!correl_org345 = CDbl(CDbl(rscorrelativo!correl_org345) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "561" Then  'JAPON
'                  If Not IsNull(rscorrelativo!correl_org561) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org561) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org561) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org561) + 1)
'                    rscorrelativo!correl_org561 = CDbl(CDbl(rscorrelativo!correl_org561) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'
'                If v_por_fte(i, 3) = "639" Then  'CUBA
'                  If Not IsNull(rscorrelativo!correl_org639) Then
'                    rstpagos!codigo_pago = CDbl(CDbl(rscorrelativo!correl_org639) + 1)
'                    rstpagos!nro_comprobante_anterior = CDbl(CDbl(rscorrelativo!correl_org639) + 1)
'                    codigo_pago1 = CDbl(CDbl(rscorrelativo!correl_org639) + 1)
'                    rscorrelativo!correl_org639 = CDbl(CDbl(rscorrelativo!correl_org639) + 1)
'                    rscorrelativo.Update
'                  End If
'                End If
'                rstpagos!org_codigo = v_por_fte(i, 3)
'            Else
'                'Set rscorrelativo = New ADODB.Recordset
'                'If rscorrelativo.State = 1 Then rscorrelativo.Close
'                'rscorrelativo.Open "select * from fc_correlativos", db, adOpenKeyset, adLockOptimistic
'                codigo_pago1 = Cont_Comp
'                'rstpagos!org_codigo = "999"        ' revisar para Cargo de Cuenta!!!!! DIC-2008
'                rstpagos!org_codigo = "111"
'            End If
'            '==== fin generación de correlativo ====
'            'MsgBox "Comprobante : " & codigo_pago1 & vbCrLf & "Organismo :     " & rstpagos!org_codigo, vbInformation + vbOKOnly, " Generando el Comprobante..."
'            rstpagos!codigo_pago = codigo_pago1
'            If i = 1 Then
'              rstpagos!da = v_EstPoa(j, 2)          'Dirección Administrativa JGCA 10/08/2007
'              rstpagos!uni_codigo = v_EstPoa(j, 10) 'v_EstPoa(I, 10) 'uni_codigo1
'              rstpagos!codigo_categoria = v_EstPoa(j, 11) 'v_EstPoa(I, 11) 'codigo_categoria1
'              rstpagos!codigo_convenio = v_EstPoa(j, 12) 'codigo_convenio1
'              CONVE = v_EstPoa(j, 12)
'              CATEG = v_EstPoa(j, 11)
'              rstpagos!es_base = "S"
'            End If
'            If i = 2 Then
'               rstpagos!da = v_EstPoa(i - 1, 2)        'Dirección Administrativa JGCA 10/08/2007
'               rstpagos!uni_codigo = v_EstPoa(i - 1, 10) 'uni_codigo1
''              If rstao_solicitud_detalle!por_fte_nal = 100 Or rstao_solicitud_detalle!por_fte_ext = 100 Then
''                rstpagos!codigo_categoria = v_EstPoa(I - 1, 11)
''                rstpagos!codigo_convenio = v_EstPoa(I - 1, 12)
''              Else
''                v_por_fte(I, 3) = "S/C TGNP"
'                'rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1 JQA JUL-2005
''                Print v_por_fte(j, 3)
'                'rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1  JQA JUL-2005
''              End If
'                rstpagos!codigo_convenio = conv_nal     'codigo_convenio2  JQA JUL-2005
'                rstpagos!codigo_categoria = cat_nal    'codigo_categoria1  JQA JUL-2005
'                rstpagos!es_base = "N"
'            End If
'
'            If i = 3 Then
'              rstpagos!da = v_EstPoa(1, 2)          'Dirección Administrativa JGCA 10/08/2007
'              rstpagos!uni_codigo = v_EstPoa(1, 10) 'uni_codigo1
''              rstpagos!codigo_categoria = "S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
''              rstpagos!codigo_convenio = "S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
'              'rstpagos!codigo_convenio = fbusCatConv(v_por_fte(i, 3), 1)  '"S/C TGN"  'v_EstPoa(i - 1, 12) 'codigo_convenio1
'              'rstpagos!codigo_Categoria = fbusCatConv(v_por_fte(i, 3), 2) '"S/C TGN" 'v_EstPoa(i - 1, 11) 'codigo_categoria1
'              rstpagos!codigo_convenio = conv_nal     'codigo_convenio2  JQA JUL-2005
'              rstpagos!codigo_categoria = cat_nal    'codigo_categoria1  JQA JUL-2005
'              rstpagos!es_base = "N"
'            End If
'            rstpagos!codigo_solicitud = adoorigen!codigo_solicitud
'            rstpagos!codigo_unidad = adoorigen!codigo_unidad
'            rstpagos!fte_codigo = v_por_fte(i, 2)
'            rstpagos!Codigo_orden = adoorigen!codigo_solicitud
'            rstpagos!codigo_documento = "D13"
'            rstpagos!Deducciones = 1
'            rstpagos!justificacion = adoorigen!justificacion_solicitud   'adoorigen.txtjustifica
'            rstpagos!justificacion = adoorigen!caracteristicas
'            rstpagos!tipo_moneda = rstao_solicitud_detalle!tipo_moneda 'adoorigen!tipo_moneda   'DtCDenominacion_moneda.bounttext  '"Bs." 'DtCTipoMoneda.Text
'            If i = 1 Then
'                rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + (rstao_solicitud_detalle!monto_Bolivianos)
'                rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + rstao_solicitud_detalle!monto_dolares
'            End If
'            If i = 2 Then
'              If v_EstPoa(j, 12) <> "FIN_PROPIO" Then
'                ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
'                tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
'                'abel 2004
'                If IsNull(rstpagos!monto_Bolivianos) Then
'                    rstpagos!monto_Bolivianos = 0
'                Else
'                    rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
'                End If
'                If IsNull(rstpagos!monto_dolares) Then
'                    rstpagos!monto_dolares = 0
'                Else
'                    rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'                End If
'              Else
'                rstpagos!monto_Bolivianos = Val(rstao_solicitud_detalle!monto_bolivianos_contra)
'                rstpagos!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra)
'              End If
'              If rstao_solicitud_detalle!monto_Bolivianos > 0 Then
'                rstpagos!es_base = "N"
'              Else
'                rstpagos!es_base = "S"
'              End If
'            End If
'            If i = 3 Then
'              tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
'              rstpagos!monto_Bolivianos = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
'              rstpagos!monto_dolares = IIf(IsNull(rstpagos!monto_dolares), 0, rstpagos!monto_dolares) + ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'              rstpagos!es_base = "I"
'            End If
'
'            If k = 1 Then
'              'rstpago_detalle.Open "select * from pago_detalle_espera where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'              rstpagos!tipo_comp = "DAC"
'            Else
'              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '999' ", db, adOpenKeyset, adLockOptimistic
'              'rstpagos!tipo_comp = "PCE"       ' DIC-2008
'              rstpagos!tipo_comp = "DAC"
'            End If
'            rstpagos!liquido_pagar = IIf(IsNull(rstpagos!monto_Bolivianos), 0, rstpagos!monto_Bolivianos) + ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
'            'rstpagos!tipo_formulario = GlNombFor
'            rstpagos!formulario = GlNombFor
'            'rstpagos!estado_aprobacion  = "X"
'            rstpagos!estado_devengado = ""
'            If GlNombFor = "F04" Then
'              rstpagos!estado_compromiso = "S"
'              rstpagos!es_licitacion = "S"
'            'AQUI ULTIMO
'            Else
'              rstpagos!estado_compromiso = "N"
''              rstpagos!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
'            End If
'            If GlNombFor = "F11" Or GlNombFor = "F01" Then
'              rstpagos!es_licitacion = "D"
''              rstpagos!estado_compromiso = "S"
''              rstpagos!estado_devengado = "S"
'            End If
'            If GlNombFor = "F05" Or GlNombFor = "F10" Then
'              rstpagos!duracion_estimada_tiempo = adoorigen!duracion_estimada_tiempo
'              rstpagos!duracion_estimada_numero = adoorigen!duracion_estimada_numero
'              rstpagos!por_tiempo = adoorigen!por_tiempo
'              rstpagos!estado_compromiso = "S"
'              rstpagos!estado_devengado = ""
'              rstpagos!fecha_estimada_inicio = IIf(IsNull(adoorigen!fecha_estimada_inicio), Date, Format(adoorigen!fecha_estimada_inicio, "dd/mm/yyyy"))
'              rstpagos!Lista_adjunta = adoorigen!Lista_adjunta
'              rstpagos!periodo_de_trabajo = ""
'            End If
'            'rstpagos!estado_devengado  = ""
'            'rstpagos!estado_pagado  = ""
'            If GlNombFor = "F03" Or GlNombFor = "F12" Or GlNombFor = "F11" Then
'              rstpagos!tipo_formulario = "CYD"
'              rstpagos!estado_compromiso = "N"
'              rstpagos!estado_devengado = "N"
'            Else
'              rstpagos!tipo_formulario = "COM"
'            End If
'            'If (GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "R" Or Trim(adoorigen!tipo_bien_Cta_doc) = "C")) Then
'            If (GlNombFor = "F01") Then
'              rstpagos!tipo_formulario = "REG"
'              rstpagos!estado_compromiso = "S"
'              rstpagos!estado_devengado = "S"
'              rstpagos!estado_pagado = "N"
'            End If
'            rstpagos!fecha_egreso = Format(Date, "dd/mm/yyyy") 'CDate(adoorigen!fecha_recepcion)   ', "dd/mm/aaaa
'            rstpagos!ges_gestion = Year(Date)
'            ges_gestion1 = Year(Date)
'            rstpagos!usr_usuario = glusuario
'            rstpagos!fecha_registro = Date  ' Format(Date, "dd/mm/aaaa
'            rstpagos!hora_registro = Format(Time, "hh:mm:ss")
'            rstpagos.Update
'            If rstpagos.State = 1 Then rstpagos.Close
'            '======== fin graba pagos ========
'
'            '======== ini graba pago_detalle ========
'            Set rstpago_detalle = New adodb.Recordset
'            If rstpago_detalle.State = 1 Then rstpago_detalle.Close
'            'If GlNombFor = "F04" Or GlNombFor = "F05" Or GlNombFor = "F10" Or GlNombFor = "F11" Then  'Or GlNombFor = "F06"
'            If k = 1 Then 'Or GlNombFor = "F06"
'              rstpago_detalle.Open "select * from pago_detalle_espera where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'            Else
'              'rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '" & v_por_fte(i, 3) & "' ", db, adOpenKeyset, adLockOptimistic
'              rstpago_detalle.Open "select * from pago_detalle where codigo_pago = '" & codigo_pago1 & "' and org_codigo = '999' ", db, adOpenKeyset, adLockOptimistic
'            End If
'
'            If rstpago_detalle.RecordCount > 0 Then
'              rstpago_detalle.MoveFirst
'            Else
'              rstpago_detalle.AddNew
'            End If
'            rstpago_detalle!codigo_pago = codigo_pago1
'            rstpago_detalle!ges_gestion = ges_gestion1
'            If k = 1 Then
'              rstpago_detalle!org_codigo = v_por_fte(i, 3)
'            Else
'              'rstpago_detalle!org_codigo = "999"           ' VERIFICAR DIC-2008
'              rstpago_detalle!org_codigo = v_por_fte(i, 3)
'            End If
'            rstpago_detalle!codigo_pago_detalle = rstpago_detalle.RecordCount
'
'            rstpago_detalle!par_codigo = v_EstPoa(j, 3) 'Par_Codigo1
'            rstpago_detalle!pro_programa = v_EstPoa(j, 6) 'pro_Programa1
''            rstpago_detalle!pro_subprograma = v_EstPoa(j, 7) 'Pro_SubPrograma1
'            rstpago_detalle!pro_proyecto = v_EstPoa(j, 8) 'Pro_Proyecto1
'            rstpago_detalle!pro_actividad = v_EstPoa(j, 9) 'Pro_Actividad1
'            rstpago_detalle!codigo_beneficiario = adoorigen!CI_aprueba
'
'            '==== ini porcentajes ====
'
'            rstpago_detalle!codigo_poa = rstao_solicitud_detalle!codigo_poa 'adoorigen!codigo_poa
'            If i = 1 Then
'              rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_Bolivianos)  'adoorigen!Monto_bolivianos   '- adoorigen!monto_bolivianos_contra  '* por_fte_ext1   'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_ext1
'              rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares)  'adoorigen!monto_dolares   '- adoorigen!monto_dolares_contra  '* por_fte_ext1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_ext1
'              If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
'                rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_ext)
'                'rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
'              End If
'            End If
'            If i = 2 Then
''              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
''              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'              If v_EstPoa(j, 12) <> "FIN_PROPIO" Then
'                ext1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 1)
'                tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
'
'                'abel 2004
'                If rstao_solicitud_detalle!monto_bolivianos_contra <> 0 Then
'                    rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
'                End If
'                'abel 2004
'                If rstao_solicitud_detalle!monto_dolares_contra <> 0 Then
'                    rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'                End If
'                If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
'                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
'                End If
'              Else
'                rstpago_detalle!monto_total = Val(rstao_solicitud_detalle!monto_bolivianos_contra)
'                rstpago_detalle!monto_dolares = Val(rstao_solicitud_detalle!monto_dolares_contra)
'                If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Or GlNombFor = "F01" Then
'                  rstpago_detalle!Porcentaje = CDbl(rstao_solicitud_detalle!por_fte_nal)
'                  rstpago_detalle!Porcentaje = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 2)
'                End If
'              End If
''''              rstpago_detalle!monto_total = rstao_solicitud_detalle!monto_bolivianos_contra  'adoorigen!monto_bolivianos_contra   'adoorigen!monto_bolivianos  * por_fte_nal1 'adoorigen.adosolicitud.Recordset!monto_bolivianos  * por_fte_nal1
''''              rstpago_detalle!monto_dolares = rstao_solicitud_detalle!monto_dolares_contra 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'            End If
'            If i = 3 Then
'              tgn1 = verporcen(v_EstPoa(j, 3), v_EstPoa(j, 8), v_EstPoa(j, 12), v_EstPoa(j, 11), 3)
'              rstpago_detalle!monto_total = ((rstao_solicitud_detalle!monto_bolivianos_contra * tgn1) / (100 - ext1))
'              rstpago_detalle!monto_dolares = ((rstao_solicitud_detalle!monto_dolares_contra * tgn1) / (100 - ext1)) 'adoorigen!monto_dolares_contra   'adoorigen!monto_dolares  * por_fte_nal1  'adoorigen.adosolicitud.Recordset!monto_dolares  * por_fte_nal1
'              If GlNombFor = "F05" Or GlNombFor = "F04" Or GlNombFor = "F10" Or GlNombFor = "F11" Then
'                rstpago_detalle!Porcentaje = tgn1
'              End If
'            End If
'            '==== fin porcentajes ====
'
'            'rstpago_detalle!Deducciones  = Val(TxtDeducciones.Text)
'            'rstpago_detalle!saldo_bolivianos  = Val(TxtSaldo.Text)
'            rstpago_detalle!tipo_cambio = Val(rstao_solicitud_detalle!tipo_cambio)
'            rstpago_detalle!estado_aprobacion = "N"
'            rstpago_detalle!fecha_pago = Format(Date, "DD/MM/YYYY")  ', "dd/mm/aaaa
'            rstpago_detalle!fecha_registro = Format(Date, "DD/MM/YYYY")
'            rstpago_detalle!usr_usuario = glusuario
'            rstpago_detalle!hora_registro = Format(Time, "hh:mm:ss")
'            rstpago_detalle.Update
'            '======== fin graba pago_detalle
'          Next          'del segundo i      Para nor. comprobantes por cada "pagos_espera" o cada "pagos"
'          'k = 1
'          'k = k + 1
'          If k = 1 Then
'            Call contabPCE(adosolicitud.Recordset, GlNombFor)
'          End If
'         Next           'del k          Para cambiar de "pagos_espera" a "pagos"
'         'End
'         Set rstdestino = New adodb.Recordset
'         If rstdestino.State = 1 Then rstdestino.Close
'         rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'         If rstdestino.RecordCount > 0 Then
'            rstdestino!estado_enviado = "S"
'            rstdestino.Update
'            rstao_solicitud_recibido.AddNew
'            rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
'            rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), CStr(0), CStr(adoorigen!codigo_solicitud))
'            rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
'            rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
'            rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
'            rstao_solicitud_recibido!fecha_solicitud = Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
'            rstao_solicitud_recibido!swSubir = swSubir
'            rstao_solicitud_recibido!usr_usuario = glusuario
'            rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
'            rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
'            rstao_solicitud_recibido.Update
'         End If
'         If rstdestino.State = 1 Then rstdestino.Close
'         rstao_solicitud_detalle.MoveNext
'        Next        'del j
'        db.CommitTrans
'        If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'      End If
'    End If  'fin swpresup
'  End If
'End Sub
'
'Private Sub contabPCE(adoorigen, GlNombFor)
'  '======== tipo de formualrio F01 ========
'  If GlNombFor = "F01" And (Trim(adoorigen!tipo_bien_Cta_doc) = "A") Then
'    tot_reg = 0
'    Dim rstdetalle As New adodb.Recordset
'    Set rstdetalle = New adodb.Recordset
'    If rstdetalle.State = 1 Then rstdetalle.Close
'    rstdetalle.Open "select * from ao_Solicitud_detalle where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockReadOnly
'    If rstdetalle.RecordCount < 1 Then
'      MsgBox "No se puede generar el asiento contable," & vbCrLf & "debido a que el registro no tiene el detalle de montos.", vbOKOnly + vbCritical, "Error al generar el asiento contabl..."
'      If rstdetalle.State = 1 Then rstdetalle.Close
'      Exit Sub
'    Else
'      tot_reg = 0
'      If rstdetalle!monto_Bolivianos > 0 Then tot_reg = tot_reg + 1
'      If rstdetalle!monto_bolivianos_contra > 0 Then tot_reg = tot_reg + 1
'    End If
'
'    Set rstao_solicitud_recibido = New adodb.Recordset
'    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'    rstao_solicitud_recibido.Open "SELECT * FROM ao_solicitud_recibido", db, adOpenKeyset, adLockOptimistic
'    'db.BeginTrans
'    '======== ini registro de co_comprobante_M ========
'    Dim rstCodComp As New adodb.Recordset
'    Set rstdestino = New adodb.Recordset
'    For i = 1 To 1  'tot_reg
'      If rstdetalle!monto_Bolivianos <= 0 And i = 1 Then
'        GoTo etiq
'      End If
'      If rstdetalle!monto_bolivianos_contra <= 0 And i = 2 Then
'        GoTo etiq
'      End If
'      '======== ini GENERA EL CODIGO DE COMPROBANTE ========
'      Set rstCodComp = New adodb.Recordset
'      rstCodComp.CursorLocation = adUseClient
'      If rstCodComp.State = 1 Then rstCodComp.Close
'      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
'      If rstCodComp.RecordCount > 0 Then
'        Cont_Comp = Val(rstCodComp!numero_correlativo)
'        Cont_Comp = Cont_Comp + 1
'        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
'        rstCodComp.Update
'      End If
'      If rstCodComp.State = 1 Then rstCodComp.Close
'      '======== fin TERMINA GENERACION DE COMPROBANTE ========
'
'      '======== ini registro co_comprobantre_m ========
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'      End If
'      rstdestino.AddNew
'      rstdestino!Cod_Comp = Cont_Comp
'      rstdestino!cod_trans = "0"
'      If i = 1 Then
'        rstdestino!org_codigo = "999" 'adoorigen!org_codigo_ext
'      End If
'      If i = 2 Then
'        rstdestino!org_codigo = "999"  'rstdestino!org_codigo = adoorigen!org_codigo_contra
'      End If
'      rstdestino!cod_trans_detalle = 1
'      rstdestino!num_respaldo = adoorigen!codigo_unidad & "/" & Str(adoorigen!codigo_solicitud)
'      rstdestino!codigo_solicitud = (adoorigen!codigo_solicitud) 'adoorigen!codigo_unidad '& "/" & Str(adoorigen!codigo_solicitud)
'      rstdestino!codigo_unidad = (adoorigen!codigo_unidad)
'      rstdestino!fecha_A = Format(Date, "dd/mm/yyyy")         'Format(adoorigen!fecha_solicitud, "dd/mm/yyyy")
'      rstdestino!codigo_beneficiario = adoorigen!CI_aprueba
'      rstdestino!Origen = "1"
'      'aqui fBuscaFteCorta(fte_1)
'      If i = 1 Then
'        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_ext) & ": " & Round((rstdetalle!monto_Bolivianos * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
'      End If
'      If i = 2 Then
'        rstdestino!glosa = Trim(adoorigen!justificacion_solicitud) & " " & fBuscaOrgCorta(rstdetalle!org_codigo_contra) & ": " & Round((rstdetalle!monto_bolivianos_contra * 100 / (rstdetalle!monto_Bolivianos + rstdetalle!monto_bolivianos_contra)), 2) & "%"
'      End If
'      rstdestino!Status = "S"
'      rstdestino!ges_gestion = adoorigen!ges_gestion
'      rstdestino!codigo_documento = "D13"
'      rstdestino!tipo_comp = "PCE" 'IIf(adoorigen!codigo_tipo = "DEV", "CAD", IIf(adoorigen!codigo_tipo = "REC", "CAR", v_Tipo_Comp(i)))
'      '        rstdestino!tipo_moneda = adoorigen!tipo_moneda
'      rstdestino!usr_usuario = glusuario
'      rstdestino!fecha_registro = Date
'      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
'      rstdestino!tipo_moneda = rstdetalle!tipo_moneda
'      rstdestino.Update
'      '======== fin registro co_comprobantre_m ========
'      '======== ini registra CO_diaRIO ========
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'        rstdestino.MoveFirst
'      Else
'        rstdestino.AddNew
'        rstdestino!Cod_Comp = Cont_Comp
'      End If
'
'      rstdestino!tipo_comp = "PCE"
'      rstdestino!d_cuenta = "1127"
'  'y        rstdestino!D_Nombre = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'      rstdestino!d_subcta1 = "02"
'      Select Case adoorigen!subcta2
'        Case "01" '"Regulares" 'Cargos de Cuenta Regulares
'          rstdestino!d_subcta2 = "01"
'          rstdestino!d_Aux3 = "00"
'        Case "02" '"Otros" 'Cargos de Cuenta Otros
'          rstdestino!d_subcta2 = "02"
'          rstdestino!d_Aux3 = "00"
'        Case "03"  '"PASE" 'Cargos de Cuenta PASE
'          rstdestino!d_subcta2 = "03"
'          rstdestino!d_Aux3 = "10"
'      End Select
'      rstdestino!d_Aux1 = "01"
'      rstdestino!d_Aux2 = "09"
''      rstdestino!d_Aux3 = "00"
'      rstdestino!d_cta_larga = adoorigen!CI_aprueba
'      rstdestino!d_des_Larga = "-" ' CAMPO PARA ELIMINAR
'      If i = 1 Then
'        rstdestino!d_montoBs = rstdetalle!monto_Bolivianos
'        rstdestino!d_montoDl = rstdetalle!monto_dolares
'        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_ext   'GABY
'      End If
'      If i = 2 Then
'        rstdestino!d_montoBs = rstdetalle!monto_bolivianos_contra
'        rstdestino!d_montoDl = rstdetalle!monto_dolares_contra
'        rstdestino!d_ctaaux2 = rstdetalle!org_codigo_contra  'GABY
'      End If
'      rstdestino!d_Cambio = rstdetalle!tipo_cambio
'      rstdestino!h_cuenta = "2116"
'  'Y        rstdestino!H_Nombre = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'      rstdestino!h_subcta1 = "02"
'      rstdestino!h_subcta2 = "00"
'      rstdestino!h_Aux1 = "01"
'      rstdestino!h_Aux2 = "09"   'Y
'      rstdestino!h_Aux3 = "00"
'      rstdestino!h_cta_larga = adoorigen!CI_aprueba
'      rstdestino!h_des_Larga = "-"   ' CAMPO PARA ELIMINAR
'      If i = 1 Then
'        rstdestino!h_montoBs = rstdetalle!monto_Bolivianos
'        rstdestino!h_montoDl = rstdetalle!monto_dolares
'        rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
'        rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
'        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
'
'        'rsCo_diario!d_Aux3 = "10"
'        'rsCo_diario!d_ctaaux3 = DtCCodigo.Text
'
'      End If
'      If i = 2 Then
'        rstdestino!h_montoBs = rstdetalle!monto_bolivianos_contra
'        rstdestino!h_montoDl = rstdetalle!monto_dolares_contra
'        rstdestino!h_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
'        rstdestino!d_ctaaux2 = "FIN_PROPIO" 'rstdetalle!codigo_convenio
'        rstdestino!d_CtaAux3 = IIf(IsNull(rstdetalle!aux3), "", rstdetalle!aux3)
'      End If
'      rstdestino!h_Cambio = rstdetalle!tipo_cambio
'      'grabar convenios
'      'en h_ctaaux2 y en d_ctaaux2
''      rstdestino!h_ctaaux2 = rstdetalle!codigo_convenio
''      rstdestino!d_ctaaux2 = rstdetalle!codigo_convenio
'
'      rstdestino!usr_usuario = glusuario
'      rstdestino!fecha_registro = Date
'      rstdestino!hora_registro = Format(Time, "hh:mm:ss")
'      rstdestino.Update
'      If rstdestino.State = 1 Then rstdestino.Close
'      '======== fin registra co_diario ========
'etiq:
'    Next i
''    Set rstdestino = New ADODB.Recordset
''    If rstdestino.State = 1 Then rstdestino.Close
''    rstdestino.Open "select * from ao_solicitud where ges_gestion = '" & adoorigen!ges_gestion & "' and codigo_unidad = '" & adoorigen!codigo_unidad & "' and codigo_solicitud = " & adoorigen!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
''    If rstdestino.RecordCount > 0 Then
''      rstdestino!estado_enviado = "S"
''      rstdestino.Update
''      rstao_solicitud_recibido.AddNew
''      rstao_solicitud_recibido!ges_gestion = IIf(IsNull(adoorigen!ges_gestion), "", adoorigen!ges_gestion)
''      rstao_solicitud_recibido!codigo_solicitud = IIf(IsNull(adoorigen!codigo_solicitud), 0, adoorigen!codigo_solicitud)
''      rstao_solicitud_recibido!formulario = IIf(IsNull(adoorigen!formulario), "", adoorigen!formulario)
''      rstao_solicitud_recibido!codigo_unidad = IIf(IsNull(adoorigen!codigo_unidad), "", adoorigen!codigo_unidad)
''      rstao_solicitud_recibido!justificacion_solicitud = IIf(IsNull(adoorigen!justificacion_solicitud), "", adoorigen!justificacion_solicitud)
''      'rstao_solicitud_recibido!swSubir = swSubir
''      rstao_solicitud_recibido!usr_usuario = GlUsuario
''      rstao_solicitud_recibido!fecha_registro = Format(Date, "dd/mm/yyyy")
''      rstao_solicitud_recibido!hora_registro = Format(Time, "hh:mm:ss")
''      rstao_solicitud_recibido.Update
''    End If
''    If rstdestino.State = 1 Then rstdestino.Close
'    If rstdetalle.State = 1 Then rstdetalle.Close
'    'db.CommitTrans
''    If rstao_solicitud_recibido.State = 1 Then rstao_solicitud_recibido.Close
'  End If
''  '---- fin formulario f01 ----
'End Sub
'
'Private Sub GRABADET()
''   //////////////////////////////////////////
'    Dim rstcodigo_detalle As New adodb.Recordset
'    Set rstcodigo_detalle = New adodb.Recordset
'    If rstcodigo_detalle.State = 1 Then rstcodigo_detalle.Close
'    rstcodigo_detalle.Open "select sum(ao_solicitud_LISTA.monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    'rstcodigo_detalle.Open "select sum(monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud, db, adOpenKeyset, adLockOptimistic
'    If rstcodigo_detalle.RecordCount > 0 Then
'    'db.Execute "select sum(monto_solicitud_dl) as monto_sol_bs from ao_solicitud_LISTA where codigo_unidad = '" & lblcodigo_unidad & "' and codigo_solicitud = " & lblcodigo_solicitud & " "
'    End If
'    Set rsdetalle = New adodb.Recordset
'    If rsdetalle.State = 1 Then rsdetalle.Close
'    db.BeginTrans
'      rsdetalle.Open "select * from ao_solicitud_detalle where codigo_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' and codigo_solicitud = " & adosolicitud.Recordset!codigo_solicitud, db, adOpenKeyset, adLockOptimistic
'      If rsdetalle.RecordCount > 0 Then
'      Else
'        rsdetalle.AddNew
'        rsdetalle!ges_gestion = gestion1        'Year(Date)
'        rsdetalle!codigo_unidad = adosolicitud.Recordset!codigo_unidad
'        rsdetalle!codigo_solicitud = adosolicitud.Recordset!codigo_solicitud
'        rsdetalle!codigo_detalle = rsdetalle.RecordCount + 1
'      End If
'    rsdetalle("codigo_poa") = adosolicitud.Recordset!codigo_poa
'    rsdetalle!por_fte_ext = 100 'CDbl(Txtpor_fte_ext)
'    rsdetalle!por_fte_nal = 0      'CDbl(Val(Txtpor_fte_nal))
'    rsdetalle("Tipo_cambio") = GlTipoCambioOficial
'    rsdetalle("monto_bolivianos") = rstcodigo_detalle!monto_sol_bs      '* 0.87
'    rsdetalle("monto_DOLARES") = (rstcodigo_detalle!monto_sol_bs / GlTipoCambioOficial)
'    rsdetalle("org_codigo_contra") = "111"
'    rsdetalle("org_codigo_EXT") = "111"
'    If rsdetalle!por_fte_nal = 0 Then
'        rsdetalle("monto_bolivianos_contra") = 0
'        rsdetalle("monto_dolares_contra") = 0
'    Else
'        rsdetalle("monto_bolivianos_contra") = rstcodigo_detalle!monto_sol_bs * rsdetalle!por_fte_nal
'        rsdetalle("monto_dolares_contra") = (rstcodigo_detalle!monto_sol_bs * rsdetalle!por_fte_nal) / GlTipoCambioOficial
'    End If
'    rsdetalle("tipo_moneda") = "Bs"       'DtCDenominacion_moneda.BoundText  'dtccisol.Text
'    rsdetalle("codigo_convenio") = "FIN_PROPIO"
'    rsdetalle("aux3") = "FIN_PROPIO"
'    rsdetalle("formulario") = "F01"    'Lblformulario.Caption     '"F11"  O "F01"
'    rsdetalle!usr_usuario = glusuario
'    rsdetalle!fecha_registro = Format(Date, "dd/mm/yyyy")
'    rsdetalle!hora_registro = Format(Time, "hh:mm:ss")
'    rsdetalle.Update
'    db.CommitTrans
'
'  rsdetalle.Requery
'
''////////////////////////////////////////////////
'End Sub
'
''Private Sub Form_Unload(Cancel As Integer)
''    Call GRABADET
''  '  Call SalePantalla
''  sino = MsgBox("Esta Seguro deSalir?", vbQuestion + vbYesNo, "Confirmando...")
''  If sino = vbYes Then
''    Dim rstAo_solicitud As New ADODB.Recordset
''    Set rstAo_solicitud = New ADODB.Recordset
''    If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
''    rstAo_solicitud.Open "select * from ao_solicitud where ges_gestion = '" & Trim(lblges_gestion) & "' and codigo_unidad = '" & Trim(lblcodigo_unidad) & "' and codigo_solicitud = " & lblcodigo_solicitud, db, adOpenKeyset, adLockOptimistic
''    If rstAo_solicitud.RecordCount > 0 Then
''      If rstAo_solicitud.RecordCount > 0 Then
''        rstAo_solicitud!Lista_adjunta = "S"
''      Else
''        rstAo_solicitud!Lista_adjunta = "N"
''      End If
''      rstAo_solicitud.Update
''    End If
''    If rstAo_solicitud.State = 1 Then rstAo_solicitud.Close
'''    If rstAc_departamentos.State = 1 Then rstAc_departamentos.Close
''    If rstao_solicitud_lista.State = 1 Then rstao_solicitud_lista.Close
''    If rstdestino.State = 1 Then rstdestino.Close
''    Unload Me
''  End If
''
''End Sub
'

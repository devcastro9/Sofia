VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmCP 
   Caption         =   "Tesoreria - Pagos Pendientes"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   -255
   ClientWidth     =   15240
   Icon            =   "FrmCP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   60
      Picture         =   "FrmCP.frx":6852
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   120
      Top             =   0
      Width           =   12495
      Begin VB.Label LblUni_descripcion_larga 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3480
         TabIndex        =   126
         Top             =   0
         Visible         =   0   'False
         Width           =   5160
      End
      Begin VB.Label lblUni_codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5160
         TabIndex        =   125
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO A PROVEEDORES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   405
         Left            =   8055
         TabIndex        =   124
         Top             =   120
         Width           =   3750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   240
         Left            =   1560
         TabIndex        =   123
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label6 
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
         Left            =   4200
         TabIndex        =   122
         Top             =   120
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label LblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "LblUsuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2520
         TabIndex        =   121
         Top             =   240
         Width           =   1305
      End
      Begin VB.Image Image3 
         Height          =   840
         Left            =   0
         Picture         =   "FrmCP.frx":B342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15360
      End
   End
   Begin VB.Frame FraBusca 
      BackColor       =   &H80000018&
      Caption         =   "Busca Cmpbte.de Pago:"
      ForeColor       =   &H00C00000&
      Height          =   1725
      Left            =   480
      TabIndex        =   37
      Top             =   3720
      Visible         =   0   'False
      Width           =   2520
      Begin VB.CommandButton CmdCancelarBusqueda 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   1560
         Picture         =   "FrmCP.frx":CEE8
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   840
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   240
         Picture         =   "FrmCP.frx":D0F2
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   840
         Width           =   765
      End
      Begin VB.TextBox TxtCmpte 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   495
         TabIndex        =   40
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox TxtOrg 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1800
         TabIndex        =   39
         Top             =   675
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TxtGes 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3615
         TabIndex        =   38
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label Label19 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   1830
         TabIndex        =   42
         Top             =   375
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label18 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   3900
         TabIndex        =   41
         Top             =   645
         Width           =   795
      End
   End
   Begin MSDataGridLib.DataGrid DtgPago 
      Bindings        =   "FrmCP.frx":D9BC
      Height          =   7200
      Left            =   120
      TabIndex        =   1
      Top             =   735
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   12700
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "PAGOS PENDIENTES"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "codigo_pago"
         Caption         =   "No.Pago"
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
         DataField       =   "org_codigo"
         Caption         =   "Fte.Fin."
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
         DataField       =   "codigo_solicitud"
         Caption         =   "No.Sol."
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
         DataField       =   "tipo_comp"
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
      BeginProperty Column05 
         DataField       =   "estado_pagado"
         Caption         =   "Pagado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "estado_devengado"
         Caption         =   "D"
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
         DataField       =   "estado_pagado"
         Caption         =   "T"
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
         DataField       =   "estado_entregado"
         Caption         =   "E"
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
         DataField       =   "estado_anulado"
         Caption         =   "A"
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
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoPago 
      Height          =   360
      Left            =   120
      Top             =   7920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
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
   Begin VB.Frame FraOpciones 
      Height          =   1035
      Left            =   120
      TabIndex        =   82
      Top             =   6840
      Visible         =   0   'False
      Width           =   3690
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   2640
         Picture         =   "FrmCP.frx":D9D2
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton CmdPagosParciales 
         Caption         =   "Pago Parcial"
         Height          =   720
         Left            =   840
         Picture         =   "FrmCP.frx":DE14
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   240
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   1680
         Picture         =   "FrmCP.frx":E7B6
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   240
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton CmdPagoTotal2 
         Caption         =   "Pagar"
         Height          =   750
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmCP.frx":EBF8
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame FraP 
      BackColor       =   &H00FFC0C0&
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
      Left            =   3870
      TabIndex        =   70
      Top             =   720
      Width           =   8610
      Begin VB.CommandButton CmdPagoIndividual 
         Caption         =   "Pagar"
         Height          =   720
         Left            =   210
         Picture         =   "FrmCP.frx":F59A
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   120
         Width           =   785
      End
      Begin VB.CommandButton CmdImprimirTransfer 
         Caption         =   "Transfer"
         Enabled         =   0   'False
         Height          =   720
         Left            =   2640
         Picture         =   "FrmCP.frx":10264
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cheque"
         Enabled         =   0   'False
         Height          =   720
         Left            =   1800
         Picture         =   "FrmCP.frx":119E6
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton CmdPagoGrupal 
         Caption         =   "Pago Grupo"
         Height          =   720
         Left            =   240
         Picture         =   "FrmCP.frx":13168
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   3495
         Picture         =   "FrmCP.frx":13A32
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdSalirPagos 
         Caption         =   "Salir"
         Height          =   720
         Left            =   7440
         Picture         =   "FrmCP.frx":151B4
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdRestaurarPagos 
         Caption         =   "Refresca"
         Height          =   720
         Left            =   5880
         Picture         =   "FrmCP.frx":153BE
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   4980
         Picture         =   "FrmCP.frx":155C8
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Frame FraPagosParciales 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7710
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   8720
      Begin VB.Frame FraMensajeImportante 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   1605
         TabIndex        =   6
         Top             =   6660
         Visible         =   0   'False
         Width           =   5925
         Begin VB.Label Label10 
            Caption         =   "NO EXISTE SALDO BANCARIO  ! ! !"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   480
            TabIndex        =   7
            Top             =   45
            Width           =   8610
         End
      End
      Begin MSDataGridLib.DataGrid DtGPP 
         Bindings        =   "FrmCP.frx":15E92
         Height          =   1545
         Left            =   0
         TabIndex        =   5
         Top             =   6090
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   2725
         _Version        =   393216
         Enabled         =   0   'False
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
            DataField       =   "codigo_pago_detalle"
            Caption         =   "Correl."
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
            DataField       =   "cta_codigo"
            Caption         =   "Cta. Bancaria"
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
            DataField       =   "numero_cheque_trf"
            Caption         =   "No.Transac"
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
            DataField       =   "cta_codigo_destino"
            Caption         =   "Cta. Destino."
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
            DataField       =   "codigo_beneficiario"
            Caption         =   "Beneficiario"
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
            DataField       =   "monto_bolivianos"
            Caption         =   "Bs."
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
            DataField       =   "monto_dolares"
            Caption         =   "$us"
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
            DataField       =   "tipo_cambio"
            Caption         =   "CAMBIO"
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
            DataField       =   "saldo_bolivianos"
            Caption         =   "SALDO"
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
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FraGrabarCancelar 
         BackColor       =   &H00FFC0C0&
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
         Left            =   30
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   8610
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Cancelar"
            Height          =   675
            Left            =   4200
            Picture         =   "FrmCP.frx":15EAF
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   120
            Width           =   765
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "Grabar"
            Height          =   675
            Left            =   3225
            Picture         =   "FrmCP.frx":160B9
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   120
            Width           =   765
         End
      End
      Begin TabDlg.SSTab SSTTransferencia 
         Height          =   4980
         Left            =   0
         TabIndex        =   8
         Top             =   1080
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   8784
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BackColor       =   14737632
         ForeColor       =   16711680
         TabCaption(0)   =   "DETALLE DEL PAGO"
         TabPicture(0)   =   "FrmCP.frx":162C3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FraPagoDetalle"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "DATOS CARTA TRANFERENCIA"
         TabPicture(1)   =   "FrmCP.frx":162DF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FraDatosCarta"
         Tab(1).ControlCount=   1
         Begin VB.Frame FraDatosCarta 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   4260
            Left            =   -74835
            TabIndex        =   60
            Top             =   405
            Width           =   8430
            Begin VB.Frame FraObservaciones 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Observaciones para carta de Transferencia"
               Enabled         =   0   'False
               ForeColor       =   &H00800000&
               Height          =   2850
               Left            =   165
               TabIndex        =   64
               Top             =   165
               Width           =   8235
               Begin VB.OptionButton OptObs1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Transferencia o giro que deberá realizarse del Banco  según listado."
                  Height          =   345
                  Left            =   180
                  TabIndex        =   67
                  Top             =   225
                  Width           =   7845
               End
               Begin VB.OptionButton OptObs2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "El costo de la comisión bancaria por la transferencia a realizar, debe ser descontado del monto a transferir."
                  Height          =   345
                  Left            =   180
                  TabIndex        =   66
                  Top             =   480
                  Width           =   7830
               End
               Begin VB.TextBox TxtObs 
                  Height          =   1680
                  Left            =   180
                  MaxLength       =   1110
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   65
                  Top             =   1050
                  Width           =   7995
               End
            End
            Begin VB.CheckBox ChkHonorarios 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Pago de Honorarios"
               Enabled         =   0   'False
               Height          =   345
               Left            =   360
               TabIndex        =   63
               Top             =   3045
               Width           =   2640
            End
            Begin VB.CheckBox ChkNombreBeneficiario 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Check1"
               Enabled         =   0   'False
               Height          =   375
               Left            =   375
               TabIndex        =   62
               Top             =   3750
               Width           =   270
            End
            Begin VB.TextBox TxtBeneDest 
               Enabled         =   0   'False
               Height          =   345
               Left            =   705
               TabIndex        =   61
               Top             =   3780
               Width           =   7605
            End
            Begin VB.Label Label22 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nombre de Beneficiario Destino:"
               Height          =   300
               Left            =   360
               TabIndex        =   68
               Top             =   3480
               Width           =   2790
            End
         End
         Begin VB.Frame FraPagoDetalle 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   4485
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   8385
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   1095
               Left            =   0
               TabIndex        =   87
               Top             =   480
               Width           =   8370
               Begin VB.TextBox TxtFechaPago 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  DataSource      =   "AdoPago"
                  ForeColor       =   &H00C00000&
                  Height          =   315
                  Left            =   3960
                  TabIndex        =   119
                  Top             =   255
                  Width           =   1320
               End
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  DataSource      =   "AdoPago"
                  ForeColor       =   &H00C00000&
                  Height          =   240
                  Left            =   6720
                  TabIndex        =   109
                  Text            =   "Bs."
                  Top             =   255
                  Width           =   300
               End
               Begin VB.TextBox TxtCodigoOrden 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  DataField       =   "codigo_pago"
                  DataSource      =   "AdoPago"
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
                  Height          =   195
                  Left            =   1365
                  TabIndex        =   103
                  Top             =   255
                  Width           =   1155
               End
               Begin VB.Frame FraBeneficiario 
                  Height          =   1950
                  Left            =   10500
                  TabIndex        =   92
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   7080
                  Begin VB.TextBox TxtDenominacionBeneficiario 
                     Height          =   285
                     Left            =   1725
                     TabIndex        =   98
                     Top             =   780
                     Width           =   4275
                  End
                  Begin VB.TextBox TxtCodigoBeneficiario 
                     DataField       =   "codigo_beneficiario"
                     DataSource      =   "AdoDetalle"
                     Height          =   285
                     Left            =   90
                     TabIndex        =   97
                     Top             =   0
                     Width           =   1575
                  End
                  Begin VB.ComboBox CmbTipoBeneficiario 
                     Height          =   315
                     ItemData        =   "FrmCP.frx":162FB
                     Left            =   90
                     List            =   "FrmCP.frx":16305
                     TabIndex        =   96
                     Top             =   1425
                     Width           =   2655
                  End
                  Begin VB.CommandButton CmdGrabarBeneficiario 
                     Caption         =   "&Grabar"
                     Height          =   330
                     Left            =   2760
                     TabIndex        =   95
                     Top             =   1440
                     Width           =   1230
                  End
                  Begin VB.CommandButton CmdSalirBeneficiario 
                     Caption         =   "&Salir"
                     Height          =   330
                     Left            =   5475
                     TabIndex        =   94
                     Top             =   1440
                     Width           =   1230
                  End
                  Begin VB.CommandButton CmdCancelarBeneficiario 
                     Caption         =   "&Cancel"
                     Height          =   330
                     Left            =   4117
                     TabIndex        =   93
                     Top             =   1440
                     Width           =   1230
                  End
                  Begin VB.Label Label20 
                     AutoSize        =   -1  'True
                     Caption         =   "Código (RUC/CI)"
                     Height          =   195
                     Left            =   105
                     TabIndex        =   102
                     Top             =   570
                     Width           =   1200
                  End
                  Begin VB.Label Label25 
                     AutoSize        =   -1  'True
                     Caption         =   "Nombre del Beneficiario"
                     Height          =   195
                     Left            =   1770
                     TabIndex        =   101
                     Top             =   555
                     Width           =   1680
                  End
                  Begin VB.Label Label27 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     Caption         =   "BENEFICIARIO"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   -1  'True
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   360
                     Left            =   2550
                     TabIndex        =   100
                     Top             =   135
                     Width           =   2325
                  End
                  Begin VB.Label Label31 
                     AutoSize        =   -1  'True
                     Caption         =   "Tipo"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   99
                     Top             =   1200
                     Width           =   315
                  End
               End
               Begin VB.TextBox TxtCodigoBen 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00C00000&
                  Height          =   300
                  Left            =   1140
                  TabIndex        =   91
                  Top             =   645
                  Width           =   1755
               End
               Begin VB.TextBox TxtNomBen 
                  BackColor       =   &H00E0E0E0&
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
                  ForeColor       =   &H00C00000&
                  Height          =   300
                  Left            =   2940
                  TabIndex        =   90
                  Top             =   645
                  Width           =   5160
               End
               Begin VB.TextBox TxtMontoBolivianos 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2058
                     SubFormatType   =   1
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   195
                  Left            =   7020
                  TabIndex        =   89
                  Text            =   "0"
                  Top             =   255
                  Width           =   1170
               End
               Begin VB.TextBox TxtTipo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   16394
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Left            =   1620
                  TabIndex        =   88
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Fecha Pago:"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   118
                  Top             =   240
                  Width           =   915
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Beneficiario:"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   108
                  Top             =   645
                  Width           =   870
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Nro. de Pago:"
                  Height          =   195
                  Left            =   225
                  TabIndex        =   107
                  Top             =   255
                  Width           =   990
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Monto a Pagar:"
                  Height          =   195
                  Left            =   5505
                  TabIndex        =   106
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Comprobante:"
                  Height          =   195
                  Left            =   225
                  TabIndex        =   105
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1350
               End
               Begin VB.Label LbLAprobado 
                  Caption         =   "Label15"
                  ForeColor       =   &H00000080&
                  Height          =   240
                  Left            =   5760
                  TabIndex        =   104
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   2235
               End
            End
            Begin VB.TextBox TxtNoTransac 
               Appearance      =   0  'Flat
               DataField       =   "numero_cheque_trf"
               Height          =   330
               Left            =   1560
               TabIndex        =   117
               Top             =   3630
               Width           =   2145
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0C0&
               Height          =   510
               Left            =   0
               TabIndex        =   17
               Top             =   1485
               Width           =   8370
               Begin VB.OptionButton OptDeposito 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Depósito Bancario"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   112
                  Top             =   180
                  Value           =   -1  'True
                  Width           =   1755
               End
               Begin VB.OptionButton OptChequeOrigen 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Cheque o Efectivo"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   780
                  TabIndex        =   19
                  Top             =   180
                  Width           =   1755
               End
               Begin VB.OptionButton OptTransferenciaOrigen 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Transferencia"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Left            =   5430
                  TabIndex        =   18
                  Top             =   180
                  Width           =   1665
               End
            End
            Begin VB.TextBox TxtNC 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               DataSource      =   "AdoPago"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   6360
               TabIndex        =   22
               Text            =   "Nro. de Comprobante"
               Top             =   690
               Width           =   1515
            End
            Begin VB.TextBox TxtNoTransaccion 
               Appearance      =   0  'Flat
               DataField       =   "numero_cheque_trf"
               Height          =   330
               Left            =   1560
               TabIndex        =   21
               Top             =   2790
               Width           =   2145
            End
            Begin VB.TextBox TxtDeducciones 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               Enabled         =   0   'False
               ForeColor       =   &H00C0E0FF&
               Height          =   315
               Left            =   1545
               TabIndex        =   20
               Top             =   4050
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.TextBox TxtMonto 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   16394
                  SubFormatType   =   0
               EndProperty
               Height          =   315
               Left            =   5265
               TabIndex        =   16
               Top             =   3630
               Width           =   1410
            End
            Begin VB.Frame FraTotalParcial 
               Height          =   375
               Left            =   5325
               TabIndex        =   13
               Top             =   3990
               Visible         =   0   'False
               Width           =   3015
               Begin VB.CommandButton CmdTotal 
                  Caption         =   "Pago Total"
                  Height          =   240
                  Left            =   840
                  TabIndex        =   15
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.CommandButton CmdPagoParcial 
                  Caption         =   "Pago Parcial"
                  Height          =   360
                  Left            =   1980
                  TabIndex        =   14
                  Top             =   375
                  Width           =   1245
               End
            End
            Begin VB.TextBox TxtMB 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               DataSource      =   "AdoPago"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   6360
               TabIndex        =   12
               Text            =   "MontoBolivianos"
               Top             =   900
               Width           =   1335
            End
            Begin VB.TextBox TxtTipoCambio 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   6360
               TabIndex        =   11
               Top             =   1440
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox CmbNomDep 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H00C0E0FF&
               Height          =   315
               Left            =   1560
               TabIndex        =   10
               Top             =   4050
               Visible         =   0   'False
               Width           =   3765
            End
            Begin MSDataListLib.DataCombo DtCCuentaOrigen 
               Bindings        =   "FrmCP.frx":1631E
               DataField       =   "cta_codigo"
               DataSource      =   "AdoPagoDetalle"
               Height          =   315
               Left            =   1530
               TabIndex        =   23
               Top             =   2055
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ListField       =   "cta_codigo"
               BoundColumn     =   "cta_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
               Bindings        =   "FrmCP.frx":16336
               DataField       =   "cta_codigo"
               DataSource      =   "AdoPagoDetalle"
               Height          =   315
               Left            =   1545
               TabIndex        =   24
               Top             =   2415
               Width           =   6705
               _ExtentX        =   11827
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ListField       =   "Cta_descripcion_larga"
               BoundColumn     =   "cta_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo DtcCtaTGN 
               Bindings        =   "FrmCP.frx":1634E
               DataField       =   "cta_codigo"
               DataSource      =   "AdoPagoDetalle"
               Height          =   315
               Left            =   3720
               TabIndex        =   25
               Top             =   2055
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   12640511
               ListField       =   "Cta_codigo_tgn"
               BoundColumn     =   "cta_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtCuentaDestino 
               Bindings        =   "FrmCP.frx":16366
               DataField       =   "cta_codigo_destino"
               DataSource      =   "AdoPagoDetalle"
               Height          =   315
               Left            =   1560
               TabIndex        =   115
               Top             =   3285
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               ListField       =   "cta_codigo"
               BoundColumn     =   "cta_codigo"
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtBancoDestino 
               Bindings        =   "FrmCP.frx":1637E
               DataField       =   "cta_codigo_destino"
               DataSource      =   "AdoPagoDetalle"
               Height          =   315
               Left            =   3720
               TabIndex        =   116
               Top             =   3285
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483632
               ForeColor       =   12640511
               ListField       =   "Cta_descripcion_larga"
               BoundColumn     =   "cta_codigo"
               Text            =   ""
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "No. Transacción:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   240
               TabIndex        =   114
               Top             =   2800
               Width           =   1230
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H80000010&
               Caption         =   "REGISTRO DE PAGOS"
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
               Height          =   375
               Left            =   2415
               TabIndex        =   29
               Top             =   105
               Width           =   3375
            End
            Begin VB.Label LblTC 
               BackColor       =   &H80000010&
               Caption         =   "CHEQUE/EFECTIVO"
               ForeColor       =   &H000000FF&
               Height          =   330
               Left            =   60
               TabIndex        =   113
               Top             =   1620
               Visible         =   0   'False
               Width           =   1470
            End
            Begin VB.Label LblCtaDestino 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "No. Cta.Destino:"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   285
               TabIndex        =   35
               Top             =   3315
               Width           =   1290
            End
            Begin VB.Label LblTransCheque 
               Caption         =   "CHEQUE"
               ForeColor       =   &H00808000&
               Height          =   300
               Left            =   255
               TabIndex        =   69
               Top             =   720
               Width           =   2970
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "No. Comprobante:"
               Height          =   195
               Left            =   4980
               TabIndex        =   36
               Top             =   690
               Width           =   1290
            End
            Begin VB.Label LblNumeroOrigen 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "No. Transacción:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   225
               TabIndex        =   34
               Top             =   3705
               Width           =   1230
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "No. Cta. Origen:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   315
               TabIndex        =   33
               Top             =   2115
               Width           =   1140
            End
            Begin VB.Label LblDeducciones 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Deducciones:"
               Height          =   195
               Left            =   420
               TabIndex        =   32
               Top             =   3710
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label LblMonto 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "Monto Destino:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   4065
               TabIndex        =   31
               Top             =   3705
               Width           =   1080
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fecha de Pago:"
               Height          =   195
               Left            =   5130
               TabIndex        =   30
               Top             =   1185
               Width           =   1140
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Monto:"
               Height          =   195
               Left            =   5775
               TabIndex        =   28
               Top             =   930
               Width           =   495
            End
            Begin VB.Label LblDepartamento 
               AutoSize        =   -1  'True
               BackColor       =   &H80000010&
               Caption         =   "Departamento:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   420
               TabIndex        =   27
               Top             =   4110
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.Label LblBancoDestino 
               AutoSize        =   -1  'True
               Caption         =   "Banco Destino:"
               Height          =   195
               Left            =   390
               TabIndex        =   26
               Top             =   3330
               Width           =   1095
            End
         End
      End
      Begin MSDataGridLib.DataGrid DtGPP1 
         Bindings        =   "FrmCP.frx":16396
         Height          =   1365
         Left            =   45
         TabIndex        =   3
         Top             =   6255
         Visible         =   0   'False
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2408
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSAdodcLib.Adodc AdoPagoDetalle 
         Height          =   330
         Left            =   45
         Top             =   7320
         Width           =   8535
         _ExtentX        =   15055
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
         Caption         =   "AdoPagoDetalle"
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
   Begin VB.Frame FraDetalle 
      Caption         =   "PAGOS PARCIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2895
      Left            =   3840
      TabIndex        =   0
      Top             =   4440
      Width           =   7950
      Begin MSDataGridLib.DataGrid DtGPagosParciales 
         Height          =   2496
         Left            =   84
         TabIndex        =   4
         Top             =   240
         Width           =   7776
         _ExtentX        =   13705
         _ExtentY        =   4392
         _Version        =   393216
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
            DataField       =   "codigo_pago_detalle"
            Caption         =   "COD."
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
            DataField       =   "cta_codigo"
            Caption         =   "CTA."
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
            DataField       =   "numero_cheque_trf"
            Caption         =   "CHEQUE"
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
            DataField       =   "cta_codigo_destino"
            Caption         =   "CTA. DEST."
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
            DataField       =   "codigo_beneficiario"
            Caption         =   "BENEF."
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
            DataField       =   "monto_bolivianos"
            Caption         =   "Bs."
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
            DataField       =   "monto_dolares"
            Caption         =   "$us"
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
            DataField       =   "tipo_cambio"
            Caption         =   "CAMBIO"
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
            DataField       =   "saldo_bolivianos"
            Caption         =   "SALDO"
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
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
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
         EndProperty
      End
   End
   Begin VB.Frame FraPagoGrupal 
      Caption         =   "Pago de Comprobante Grupal"
      Height          =   2370
      Left            =   3840
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   8040
      Begin VB.TextBox TxtCmpteInicial 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   270
         TabIndex        =   50
         Top             =   885
         Width           =   1515
      End
      Begin VB.TextBox TxtCmpteFinal 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1830
         TabIndex        =   49
         Top             =   870
         Width           =   1515
      End
      Begin VB.CommandButton CmdEjecutarPago 
         Caption         =   "&Pagos Totales"
         Height          =   375
         Left            =   4875
         TabIndex        =   48
         ToolTipText     =   "Los Pagos que se haran son totales y con un solo número de cuenta"
         Top             =   1590
         Width           =   1335
      End
      Begin VB.CommandButton CmdSale 
         Caption         =   "Salir"
         Height          =   360
         Left            =   6390
         TabIndex        =   47
         ToolTipText     =   "Los Pagos que se haran son totales y con un solo número de cuenta"
         Top             =   1605
         Width           =   1335
      End
      Begin VB.TextBox TxtOrganismo 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3405
         TabIndex        =   45
         Top             =   870
         Width           =   1515
      End
      Begin VB.TextBox TxtGestion 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   4995
         TabIndex        =   44
         Top             =   870
         Width           =   1515
      End
      Begin MSComctlLib.ProgressBar PrBPagosTotales 
         Height          =   300
         Left            =   135
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSDataListLib.DataCombo DtCCta 
         Bindings        =   "FrmCP.frx":163B3
         DataField       =   "cta_codigo"
         DataSource      =   "AdoPagoDetalle"
         Height          =   315
         Left            =   270
         TabIndex        =   51
         Top             =   1665
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDescripcion 
         Bindings        =   "FrmCP.frx":163CB
         DataField       =   "cta_codigo"
         DataSource      =   "AdoPagoDetalle"
         Height          =   315
         Left            =   270
         TabIndex        =   52
         Top             =   2055
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCTgn 
         Bindings        =   "FrmCP.frx":163E3
         DataField       =   "cta_codigo"
         DataSource      =   "AdoPagoDetalle"
         Height          =   315
         Left            =   2520
         TabIndex        =   53
         Top             =   1650
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label Label12 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   210
         TabIndex        =   59
         Top             =   570
         Width           =   1260
      End
      Begin VB.Label Label13 
         Caption         =   "Cmpte. Final"
         Height          =   300
         Left            =   1860
         TabIndex        =   58
         Top             =   585
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. Origen:"
         Height          =   195
         Left            =   225
         TabIndex        =   57
         Top             =   1410
         Width           =   1140
      End
      Begin VB.Label Label15 
         Caption         =   "Organismo"
         Height          =   165
         Left            =   3405
         TabIndex        =   56
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label Label16 
         Caption         =   "Cmpte. Inicial"
         Height          =   165
         Left            =   6120
         TabIndex        =   55
         Top             =   390
         Width           =   1260
      End
      Begin VB.Label Label17 
         Caption         =   "Gestión"
         Height          =   165
         Left            =   4980
         TabIndex        =   54
         Top             =   600
         Width           =   1260
      End
   End
   Begin MSAdodcLib.Adodc AdoCuenta 
      Height          =   360
      Left            =   120
      Top             =   8280
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   635
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
      Caption         =   "AdoCuenta"
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
Attribute VB_Name = "FrmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'========================================================================================
' Sistema:                  ADFIN-2002
' Módulo:                   Aprobación de Pagos
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmCP.frm
' Descipción :              Asignacón de montos, numero de cuenta bancaria
'                           si se trata de cheque o transferencia y
'                           datos de carta de transferencia
' Formularios relacionados: Main.frm (Padre)
'                           CryCheque , CryComprobante, CryTransferencia
' Fecha de creación         10/Ene/ 2001
' Fecha última modificación 1/May/ 2001
' Versión:                  2.0
'========================================================================================
Dim rsNada As New ADODB.Recordset
Dim rsPartida As New ADODB.Recordset
Dim rsPAgoDetalle As New ADODB.Recordset
Dim rspago As New ADODB.Recordset
Dim rsControlDet As New ADODB.Recordset
Dim rscuenta As New ADODB.Recordset
Dim rsbeneficiario As New ADODB.Recordset
Dim rsPagoDet As New ADODB.Recordset
Dim rsCtrlPago As New ADODB.Recordset
Dim rsCuentaBancaria As New ADODB.Recordset
Dim recSetAuxcomp As New ADODB.Recordset
Dim rsTransferencia As New ADODB.Recordset
Dim rsBusca As New ADODB.Recordset

Dim rsPD As New ADODB.Recordset
Dim rsPagoAux As New ADODB.Recordset
Dim rsCta As New ADODB.Recordset

Dim Deduccion As Integer
'Dim SumaTotal As Long
Dim SumaTotal As Currency
Dim MontoAuxiliar As Long
Dim CtaAnterior  As String
Dim NumReg  As Long
Public MontoAnterior As Long
Dim swPagoTotal As Integer
Dim swPagoParcial As Integer
Public swModifica As Integer
Dim BUSCA As String
Dim v_UNIDAD As String
Dim V_SOLICITUD As Integer
Dim comCuentasAcumuladas As New ADODB.Command


Private Sub AdoPagoDetalle_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason <> 10 Then
  If Not AdoPagoDetalle.Recordset.EOF And Not AdoPagoDetalle.Recordset.BOF Then
    If Not IsNull(AdoPagoDetalle.Recordset("cta_codigo")) Then DtCCuentaOrigen.Text = AdoPagoDetalle.Recordset("cta_codigo") Else DtCCuentaOrigen.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("cta_codigo_destino")) Then TxtCuentaDestino.Text = AdoPagoDetalle.Recordset("cta_codigo_destino") Else TxtCuentaDestino.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("numero_cheque_trf")) Then TxtNoTransaccion.Text = AdoPagoDetalle.Recordset("numero_cheque_trf") Else TxtNoTransaccion.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("monto_bolivianos")) Then TxtMonto.Text = AdoPagoDetalle.Recordset("monto_bolivianos") Else TxtMonto.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("tipo_cambio")) Then TxtTipoCambio.Text = AdoPagoDetalle.Recordset("tipo_cambio") Else TxtTipoCambio.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("fecha_pago")) Then TxtFechaPago.Text = AdoPagoDetalle.Recordset("fecha_pago") Else TxtFechaPago.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("departamento")) Then CmbNomDep.Text = AdoPagoDetalle.Recordset("departamento") Else CmbNomDep.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("banco_destino")) Then TxtBancoDestino.Text = AdoPagoDetalle.Recordset("banco_destino") Else TxtBancoDestino.Text = ""
    If Not IsNull(AdoPagoDetalle.Recordset("observacion")) Then TxtObs.Text = AdoPagoDetalle.Recordset("observacion") Else TxtObs.Text = ""
    
    '      Select Case Mid(AdoPagoDetalle.Recordset("observacion"), 1, 1)
    '        Case "T"
    '               OptObs1.Value = True
    '               TxtObs.Visible = False
    '        Case "E"
    '               OptObs2.Value = True
    '               TxtObs.Visible = False
    '        Case Else
    '               OptObs3.Value = True
    '               TxtObs.Visible = True
    '      End Select
     If Not IsNull(AdoPagoDetalle.Recordset("beneficiario_destino")) Then
          TxtBeneDest.Text = AdoPagoDetalle.Recordset("beneficiario_destino")
          ChkNombreBeneficiario.Value = 1
     Else
           TxtBeneDest.Text = TxtNomBen.Text
           ChkNombreBeneficiario.Value = 0
     End If
     If Not IsNull(AdoPagoDetalle.Recordset("honorarios")) Then
        If AdoPagoDetalle.Recordset("honorarios") = "H" Then
              ChkHonorarios.Value = 1
        End If
        If AdoPagoDetalle.Recordset("honorarios") = "S" Then
              ChkHonorarios.Value = 0
        End If

     End If
      
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "T" Then
        OptTransferenciaOrigen.Value = True
        TxtCuentaDestino.Visible = True
        LblCtaDestino.Visible = True
        TxtBancoDestino.Visible = True
        LblBancoDestino.Visible = True
        LblDepartamento.Visible = True
        CmbNomDep.Visible = True
        FraObservaciones.Visible = True
        OptTransferenciaOrigen.Value = True
        SSTTransferencia.TabVisible(0) = True
        SSTTransferencia.TabVisible(1) = True
        LblTransCheque.Caption = "TRANSFERENCIA"
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "D" Then
        OptDeposito.Value = True
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
        TxtBancoDestino.Visible = False
        LblBancoDestino.Visible = False
        LblDepartamento.Visible = False
        CmbNomDep.Visible = False
        FraObservaciones.Visible = False
        
        OptDeposito.Value = True
        SSTTransferencia.TabVisible(0) = True
        SSTTransferencia.TabVisible(1) = False
        LblTransCheque.Caption = "DEPOSITO"
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "C" Then
        OptChequeOrigen.Value = True
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
        TxtBancoDestino.Visible = False
        LblBancoDestino.Visible = False
        LblDepartamento.Visible = False
        CmbNomDep.Visible = False
        FraObservaciones.Visible = False
        
        OptChequeOrigen.Value = True
        SSTTransferencia.TabVisible(0) = True
        SSTTransferencia.TabVisible(1) = False
        LblTransCheque.Caption = "CHEQUE"
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "" Then
        OptChequeOrigen.Value = True
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
        TxtBancoDestino.Visible = False
        LblBancoDestino.Visible = False
        LblDepartamento.Visible = False
        CmbNomDep.Visible = False
        FraObservaciones.Visible = False
    
        OptTransferenciaOrigen.Value = False
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
        LblNumeroOrigen.Caption = "Nro.Transacción"
        SSTTransferencia.TabVisible(1) = False
        TxtBancoDestino.Visible = False
        LblBancoDestino.Visible = False
        'Departamento
        LblDepartamento.Visible = False
        CmbNomDep.Visible = False
        SSTTransferencia.TabVisible(0) = True
        SSTTransferencia.TabVisible(1) = False
        LblTransCheque.Caption = "CHEQUE"
        OptChequeOrigen_Click
     End If
  End If
'End If
End Sub

Private Sub AdoPago_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'If BUSCA = 1 Then Exit Sub
   If Not AdoPago.Recordset.EOF And Not AdoPago.Recordset.BOF Then
         If Not IsNull(AdoPago.Recordset("codigo_pago")) Then TxtCodigoOrden.Text = AdoPago.Recordset("codigo_pago") Else TxtCodigoOrden.Text = ""
         'If Not IsNull(AdoPago.Recordset("monto_Bolivianos")) Then TxtMontoBolivianos.Text = AdoPago.Recordset("monto_Bolivianos")
         If Not IsNull(AdoPago.Recordset("liquido_pagar")) Then TxtMontoBolivianos.Text = CDbl(AdoPago.Recordset("liquido_pagar")) Else TxtMontoBolivianos.Text = ""
         If Not IsNull(AdoPago.Recordset("tipo_comp")) Then TxtTipo.Text = AdoPago.Recordset("tipo_comp") Else TxtTipo.Text = ""
         
         'Datos del Control de Datos
         Set rsControlDet = New ADODB.Recordset
         rsControlDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
         If rsControlDet.RecordCount > 0 Then
           If Not IsNull(rsControlDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsControlDet("codigo_beneficiario") Else TxtCodigoBen.Text = ""
           'If Not IsNull(rsControlDet("deducciones")) Then TxtDeducciones.Text = rsControlDet("deducciones")
            If Not IsNull(rsControlDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsControlDet("cta_codigo") Else DtCCuentaOrigen.Text = ""
           If Not IsNull(rsControlDet("fecha_pago")) Then TxtFechaPago.Text = rsControlDet("fecha_pago") Else TxtFechaPago.Text = ""
           If Not IsNull(rsControlDet("monto_bolivianos")) Then
              LbLAprobado.Caption = "APROBADO"
           Else
              LbLAprobado.Caption = ""
           End If
           If Not IsNull(rsControlDet("cheque_o_trf")) Then
              If rsControlDet("cheque_o_trf") = "C" Then
                LblTC.Caption = "CHEQUE"
              End If
              If rsControlDet("cheque_o_trf") = "T" Then
                LblTC.Caption = "TRANSFERENCIA"
              End If
           Else
                LblTC.Caption = "POR PAGAR"
           End If
           Set AdoPagoDetalle.Recordset = rsControlDet
           Set DtGPagosParciales.DataSource = AdoPagoDetalle
           Set DtGPP.DataSource = AdoPagoDetalle
           
           DtGPagosParciales.Refresh
           AdoPagoDetalle.Refresh
           rsControlDet.MoveLast
         Else
           Set DtGPagosParciales.DataSource = rsNada
           Set DtGPP.DataSource = rsNada
           DtGPagosParciales.ReBind
           DtGPagosParciales.Refresh
         End If
         
         Set rsbeneficiario = New ADODB.Recordset
         rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario='" & TxtCodigoBen.Text & "'", db, adOpenKeyset, adLockOptimistic
         If rsbeneficiario.RecordCount > 0 Then
         TxtNomBen.Text = rsbeneficiario("denominacion_beneficiario")
         Else
         TxtNomBen.Text = ""
         
         End If
         rsbeneficiario.Close
    End If
End Sub


Private Sub CmdBorrar_Click()
    sino = MsgBox("Está seguro de eliminar este registro?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then
        'Opcional
         MsgBox "No se puede eliminar ningun registro", vbInformation
         'AdoPagoDetalle.Recordset.Delete            '        Set rsControlDet = New ADODB.Recordset
         '        rsControlDet.Open "select * from pago_detalle where ges_gestion='" & rsPago("ges_gestion") & "' and org_codigo='" & rsPago("org_codigo") & "'and codigo_pago='" & rsPago("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
         '        If rsControlDet.RecordCount > 0 Then
         '          While Not rsControlDet.EOF
         '             rsControlDet.Delete
         '             rsControlDet.MoveNext
         '          Wend
         '        End If
         '
         '        Set AdoPagoDetalle.Recordset = rsControlDet
         '        Set DtGPagosParciales.DataSource = AdoPagoDetalle
         '        Set DtGPP.DataSource = AdoPagoDetalle
         '        DtGPagosParciales.Refresh
    End If
End Sub

Private Sub CmdBuscar_Click()
Dim condicion As String

                    If TxtCmpte.Text = "" Then
                        MsgBox "Necesita números de comprobante"
                        Exit Sub
                    Else
                        condicion = "codigo_pago=" + "'" + TxtCmpte.Text + "'"
                    End If
                    
                    If TxtOrg.Text <> "" Then
                        condicion = condicion + " and org_codigo=" + "'" + txtorganismo.Text + "'"
                    End If
                    
                    If TxtGes.Text <> "" Then
                        condicion = condicion + " and ges_gestion=" + "'" + TxtGestion.Text + "'"
                    End If
                    BUSCA = 1
                    X1 = Second(Time())
                    Set rspago = New ADODB.Recordset
                    If rspago.State = 1 Then rspago.Close
                    
                    rspago.Open "select * from pagos  where  " & condicion & "  and  (estado_contabilidad='P' or estado_devengado='S' ) and  estado_aprobacion <>'A' and (estado_pagado<>'S' or estado_pagado is null) order by codigo_pago", db, adOpenKeyset, adLockOptimistic
                    'rspago.Open "select * from pagos  where  " & condicion & "  and  (estado_contabilidad='P' or estado_devengado='S' ) and  estado_aprobacion <>'A' order by codigo_pago", db, adOpenKeyset, adLockOptimistic
                    If rspago.RecordCount > 0 Then
                    
                    
                        kj = rspago.RecordCount
                        While kj <> 0
                        kj = kj - 1
                        Wend
        
                    
                        Set AdoPago.Recordset = rspago
                        Set DtgPago.DataSource = AdoPago
                        AdoPago.Refresh
                        AdoPago.Recordset.MoveNext
                        DtgPago.Refresh
                    Else
                        MsgBox "No existe registro"
                    End If
                        X2 = Second(Time())
                        
                    BUSCA = 0
'                    FraBusca.Visible = False

                    
                    
End Sub

Private Sub CmdBusqueda_Click()
    FraBusca.Visible = True
End Sub

Private Sub CmdCancelar_Click()
On Error GoTo error_cancelar:
    
                  Set rsControlDet = New ADODB.Recordset
                  rsControlDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
                  If rsControlDet.RecordCount > 0 Then
                     If Not IsNull(rsControlDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsControlDet("codigo_beneficiario")
                     'If Not IsNull(rsControlDet("deducciones")) Then TxtDeducciones.Text = rsControlDet("deducciones")
                     If Not IsNull(rsControlDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsControlDet("cta_codigo")
                     If Not IsNull(rsControlDet("fecha_pago")) Then TxtFechaPago.Text = rsControlDet("fecha_pago")
                     Set AdoPagoDetalle.Recordset = rsControlDet
                     Set DtGPagosParciales.DataSource = AdoPagoDetalle
                     Set DtGPP.DataSource = AdoPagoDetalle
                     DtGPagosParciales.Enabled = True
                     DtgPago.Enabled = True
                     DtGPagosParciales.Refresh
                     rsControlDet.MoveLast
                  End If
         
    FraGrabarCancelar.Visible = False
'    FraOpciones.Visible = True
    FraPagoDetalle.Enabled = False
    FraDatosCarta.Enabled = False
    FraP.Visible = True
Exit Sub
error_cancelar:
    MsgBox Err.Number & " " & Err.Description
End Sub


Private Sub CmdCancelarBusqueda_Click()
    FraBusca.Visible = False
End Sub

Private Sub CmdContabiliza_Click()
On Error GoTo Asiento_Err
db.BeginTrans


'***************************************************copiar
'*********************************
Set recSetGenera = New ADODB.Recordset
recSetGenera.CursorLocation = adUseClient

If recSetGenera.State = 1 Then recSetGenera.Close
recSetGenera.Open "select Cod_Comp from Co_Comprobante_M order by Cod_Comp asc ", db, adOpenDynamic, adLockOptimistic, adCmdText

    If (recSetGenera.EOF) Then
        SW = True
        Cont_Comp = 1
    Else
        SW = False
        recSetGenera.MoveLast
        Cont_Comp = (recSetGenera!Cod_Comp) + 1
    End If
   'TxtComprobante.Text = Str(Cont_Comp)
    'Txt_ges.Text = Year(Now)
    recSetGenera.Close
    'Flag_Actualizacion = "B"
    'LblTitulo.Caption = "Contra Cuenta"

Set recsetaux = New ADODB.Recordset
recsetaux.CursorLocation = adUseClient

If recsetaux.State = 1 Then recsetaux.Close
recsetaux.Open " SELECT  distinct Co_Comprobante_M.Cod_Comp,Co_Comprobante_M.Tipo_Comp,cO_comprobante_M.Num_Respaldo," & _
" Co_Comprobante_M.codigo_beneficiario,Co_Comprobante_M.codigo_Documento,Co_Comprobante_M.Fecha_A,Co_Comprobante_M.ges_gestion," & _
" Co_Comprobante_M.Glosa,status,CO_Diario.D_Aux1,CO_Diario.D_Aux2, CO_Diario.D_Aux3,Co_Diario.d_Cta_Larga,Co_Diario.D_Des_Larga,Co_Diario.cod_Comp_C, " & _
" CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2, CO_Diario.D_Nombre,CO_Diario.D_MontoBs,D_Cambio,H_Cambio," & _
" CO_Diario.D_MontoDl,CO_Diario.H_SubCta1, CO_Diario.H_SubCta2,CO_Diario.H_Aux1, CO_Diario.H_Aux2,Co_Diario.H_Cta_Larga,Co_Diario.H_Des_Larga," & _
" CO_Diario.H_Aux3,CO_Diario.H_Nombre, CO_Diario.H_MontoBs,CO_Diario.H_Montodl,CO_Diario.H_Cuenta " & _
" From CO_Diario,CO_Comprobante_M WHERE CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp AND co_Comprobante_M.Cod_Comp = val('" & TxtCodigoOrden.Text & "') " & _
" and co_Comprobante_M.Tipo_Comp='" & TxtTipo.Text & "' and CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp and status='S'", db, adOpenDynamic, adLockOptimistic, adCmdText

If recsetaux.RecordCount > 0 Then
        Set recSetAuxActualizar1 = New ADODB.Recordset
        If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
         recSetAuxActualizar1.Open " select distinct fc_Cuenta_Bancaria.fte_codigo,tipo_comp,Pagos.ORg_Codigo,Pago_Detalle.ges_Gestion,pago_Detalle.cta_Codigo," & _
        " Tipo_Comp,Pago_Detalle.fecha_Pago from Pagos,Pago_Detalle,fc_Cuenta_Bancaria where Pagos.Ges_Gestion=Pago_Detalle.Ges_Gestion and Pagos.Org_Codigo=Pago_Detalle.Org_Codigo and  Pagos.Codigo_Pago=Pago_Detalle.Codigo_Pago   " & _
        " and Pagos.Tipo_Comp='" & TxtTipo.Text & "' and Pagos.Codigo_Pago = val('" & TxtCodigoOrden.Text & "') and fc_Cuenta_Bancaria.Cta_Codigo=Pago_Detalle.Cta_Codigo ", db, adOpenDynamic, adLockOptimistic, adCmdText
      '  If recSetAuxActualizar1.RecordCount > 0 Then

'        Else
'        MsgBox "No existen cuentas asociadas................"
'        End If
            Set recsetAdicion = New ADODB.Recordset
            If recsetAdicion.State = 1 Then recsetAdicion.Close
            recsetAdicion.Open " select * from Co_Comprobante_M  ", db, adOpenDynamic, adLockOptimistic, adCmdText
             
             recsetAdicion.AddNew
                    
                recsetAdicion!Cod_Comp = Cont_Comp
                recsetAdicion!tipo_comp = recsetaux!tipo_comp
                recsetAdicion!ges_gestion = recSetAuxActualizar1!ges_gestion
                recsetAdicion!fecha_A = recSetAuxActualizar1!fecha_pago
                
             Select Case recsetaux!tipo_comp
                Case "PCE"
                    recsetAdicion!codigo_beneficiario = recsetaux!codigo_beneficiario
                Case "PCO"
             End Select
                
                recsetAdicion!glosa = recsetaux!glosa
                recsetAdicion!codigo_documento = recsetaux!codigo_documento
                recsetAdicion!num_respaldo = recsetaux!num_respaldo
                recsetAdicion!Status = recsetaux!Status
                recsetAdicion.Update
        
        
        '********* adicion Debitos creditos
            Set recSetAuxActualizar = New ADODB.Recordset
            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
            recSetAuxActualizar.Open " select * from Co_Diario  ", db, adOpenDynamic, adLockOptimistic, adCmdText
            recSetAuxActualizar.AddNew
        
                recSetAuxActualizar!Cod_Comp = Cont_Comp
                recSetAuxActualizar!tipo_comp = recsetaux!tipo_comp
                recSetAuxActualizar!Cod_Comp_C = recsetaux!Cod_Comp
                recSetAuxActualizar!d_cuenta = recsetaux!h_cuenta
                recSetAuxActualizar!d_subcta1 = recsetaux!h_subcta1
                recSetAuxActualizar!d_subcta2 = recsetaux!h_subcta2
        
                recSetAuxActualizar!d_Aux1 = recsetaux!h_Aux1
                recSetAuxActualizar!d_Aux2 = recsetaux!h_Aux2
                recSetAuxActualizar!d_Aux3 = recsetaux!h_Aux3
        
                recSetAuxActualizar!d_cta_larga = recsetaux!h_cta_larga
                recSetAuxActualizar!d_des_Larga = recsetaux!h_des_Larga
        
                recSetAuxActualizar!d_montoBs = recsetaux!h_montoBs
                recSetAuxActualizar!d_montoDl = recsetaux!h_montoDl
                recSetAuxActualizar!d_Cambio = recsetaux!h_Cambio
              Select Case recSetAuxActualizar1!FTE_codigo
               Case "10"
                recSetAuxActualizar!h_cuenta = "1111"
                recSetAuxActualizar!h_subcta1 = "02"
                recSetAuxActualizar!h_subcta2 = "01"
        
                recSetAuxActualizar!h_Aux1 = "02"
                recSetAuxActualizar!h_Aux2 = "00"
                recSetAuxActualizar!h_Aux3 = "00"
               
               Case "70"
                recSetAuxActualizar!h_cuenta = "1111"
                recSetAuxActualizar!h_subcta1 = "02"
                recSetAuxActualizar!h_subcta2 = "02"
        
                recSetAuxActualizar!h_Aux1 = "02"
                recSetAuxActualizar!h_Aux2 = "00"
                recSetAuxActualizar!h_Aux3 = "00"
               
               Case "80"
                recSetAuxActualizar!h_cuenta = "1111"
                recSetAuxActualizar!h_subcta1 = "02"
                recSetAuxActualizar!h_subcta2 = "03"
        
                recSetAuxActualizar!h_Aux1 = "02"
                recSetAuxActualizar!h_Aux2 = "00"
                recSetAuxActualizar!h_Aux3 = "00"
             End Select
                recSetAuxActualizar!h_cta_larga = recsetaux!h_cta_larga
                recSetAuxActualizar!h_des_Larga = recsetaux!h_des_Larga
        
                recSetAuxActualizar!h_montoBs = recsetaux!h_montoBs
                recSetAuxActualizar!h_montoDl = recsetaux!h_montoDl
                recSetAuxActualizar!h_Cambio = recsetaux!h_Cambio
        
                recSetAuxActualizar.Update
'         Else
'         MsgBox "No existen cuentas asociadas................"
'         End If


Else
MsgBox "No existen registros"

End If
db.CommitTrans

Exit Sub
Asiento_Err:
    MsgBox "Error al generar contra cuenta"
    db.RollbackTrans
    'CmdAgregarDetalle.Enabled = True
    'Cmd_Modificar.Enabled = True
    'Cmd_Aprobar.Enabled = True
    'CmdSalir.Enabled = True
    'Cmd_GrabaM.Enabled = True
    'Cmd_Cancelar.Enabled = True
    'Cmd_Copiar.Enabled = True

End Sub

Private Sub CmdEjecutarPago_Click()
Dim CMPTE As Long
On Error GoTo error_ejecutar:

    If TxtCmpteInicial.Text = "" Or TxtCmpteInicial.Text = "" Or DtCCta.Text = "" Then
        MsgBox "Necesita números del comprobante inicial, final y la cuenta origen"
        Exit Sub
    End If
    
    If txtorganismo.Text = "" Then
        MsgBox "Necesita organismo"
        Exit Sub
    End If
    
    If TxtGestion.Text = "" Then
        MsgBox "Necesita gestión"
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    PrBPagosTotales.Visible = True
    For CMPTE = Val(TxtCmpteInicial.Text) To Val(TxtCmpteFinal.Text) Step 1
            PrBPagosTotales.Value = CMPTE
            Set rsPD = New ADODB.Recordset
            If rsPD.State = 1 Then rsPD.Close
            rsPD.Open "select * from pago_detalle where codigo_pago= '" & CMPTE & "' and org_codigo= '" & txtorganismo.Text & "' and ges_gestion= '" & TxtGestion.Text & "'", db, adOpenKeyset, adLockOptimistic
         
            'Modificando los datos enviados de Contabilidad o Devengado
            rsPD("cheque_o_trf") = "C"
            rsPD("numero_cheque_trf") = ""
            If DtCCta.Text <> "" Then
                rsPD("cta_codigo") = DtCCta.Text
            Else
                MsgBox "Introducir Cuenta Origen", vbCritical + vbInformation, "Validación de datos"
                Exit Sub
            End If
            If rsPagoAux.State = 1 Then rsPagoAux.Close
            rsPagoAux.Open "select * from pagos where codigo_pago='" & CMPTE & "'", db, adOpenKeyset, adLockOptimistic
            If rsPagoAux.RecordCount > 0 Then
                'rsPago.Find "codigo_pago='" & Str(CMPTE) & "'", , adSearchForward
                'MsgBox rsPago("liquido_pagar")
            End If
            If Not IsNull(rsPagoAux("liquido_pagar")) Then
                rsPD("monto_bolivianos") = rsPagoAux("liquido_pagar")
                rsPD("monto_dolares") = rsPagoAux("liquido_pagar") / rsPD("tipo_cambio")
            Else
                MsgBox "Introducir Monto total", vbCritical + vbInformation, "Validación de datos"
                Exit Sub
            End If
            If rsPD("monto_bolivianos") <> "" Then
                rsPD("literal") = Literal(CStr(rsPD("monto_bolivianos"))) + " BOLIVIANOS"
            End If
            'Datos de seguimiento
            rsPD("usr_usuario") = LblUsuario.Caption
            rsPD("fecha_registro") = Date
            rsPD("hora_registro") = Format(Time, "hh:mm:ss")
            rsPD.Update
     Next CMPTE
    
    ' Actualizacion de Cuenta Saldo Actual
    If rsCuentaBancaria.State = 1 Then rsCuentaBancaria.Close
    Set rsCuentaBancaria = New ADODB.Recordset
    rsCuentaBancaria.Open "select * from fc_cuenta_bancaria where Cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rsCuentaBancaria.RecordCount > 0 Then
       rsCuentaBancaria("Cta_saldo_actual") = rsCuentaBancaria("Cta_saldo_actual") - Val(TxtMonto.Text)
       rsCuentaBancaria.Update
    Else
       MsgBox "No existe Nro. de Cuenta", vbInformation, "Validación"
    End If
    
    Me.MousePointer = vbDefault
    PrBPagosTotales.Visible = False
    MsgBox "Se terminó de procesar"
Exit Sub
error_ejecutar:
    MsgBox Err.Number & " " & Err.Description
End Sub
'Private Sub Cmd_Contabiliza(P_codigo_pago As String)
'On Error GoTo Asiento_Err
'db.BeginTrans
'
'''***************************************************copiar
'''*********************************
''Set recSetGenera = New ADODB.Recordset
''recSetGenera.CursorLocation = adUseClient
''
''If recSetGenera.State = 1 Then recSetGenera.Close
''recSetGenera.Open "select Cod_Comp from Co_Comprobante_M order by Cod_Comp asc ", db, adOpenDynamic, adLockOptimistic, adCmdText
''
''    If (recSetGenera.EOF) Then
''        sw = True
''        Cont_Comp = 1
''    Else
''        sw = False
''        recSetGenera.MoveLast
''        Cont_Comp = (recSetGenera!Cod_Comp) + 1
''    End If
''   'TxtComprobante.Text = Str(Cont_Comp)
''    'Txt_ges.Text = Year(Now)
''    recSetGenera.Close
'    'Flag_Actualizacion = "B"
'   'LblTitulo.Caption = "Contra Cuenta"
'
'Set recSetAux = New ADODB.Recordset
'recSetAux.CursorLocation = adUseClient
'
'If recSetAux.State = 1 Then recSetAux.Close
'recSetAux.Open " SELECT  distinct Co_Comprobante_M.Cod_Comp,Co_Comprobante_M.Tipo_Comp,cO_comprobante_M.Num_Respaldo," & _
'" Co_Comprobante_M.codigo_beneficiario,Co_Comprobante_M.codigo_Documento,Co_Comprobante_M.Fecha_A,Co_Comprobante_M.ges_gestion," & _
'" Co_Comprobante_M.Glosa,status,CO_Diario.D_Aux1,CO_Diario.D_Aux2, CO_Diario.D_Aux3,Co_Diario.d_Cta_Larga,Co_Diario.D_Des_Larga,Co_Comprobante_M.cod_Comp, " & _
'" CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2, CO_Diario.D_Nombre,CO_Diario.D_MontoBs,D_Cambio,H_Cambio," & _
'" CO_Diario.D_MontoDl,CO_Diario.H_SubCta1, CO_Diario.H_SubCta2,CO_Diario.H_Aux1, CO_Diario.H_Aux2,Co_Diario.H_Cta_Larga,Co_Diario.H_Des_Larga," & _
'" CO_Diario.H_Aux3,CO_Diario.H_Nombre, CO_Diario.H_MontoBs,CO_Diario.H_Montodl,CO_Diario.H_Cuenta " & _
'" From CO_Diario,CO_Comprobante_M WHERE CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp AND co_Comprobante_M.Cod_Comp = val('" & P_codigo_pago & "') " & _
'" and co_Comprobante_M.Tipo_Comp='PCE' and CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp and status='S' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'If recSetAux.RecordCount > 0 Then
' 'MsgBox recSetAux!Cod_Comp
'       Set recSetAuxActualizar1 = New ADODB.Recordset
'       recSetAuxActualizar1.CursorLocation = adUseClient
'        If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'         recSetAuxActualizar1.Open " select distinct fc_Cuenta_Bancaria.fte_codigo,tipo_comp,Pagos.ORg_Codigo, " & _
'        " Pago_Detalle.ges_Gestion,pago_Detalle.cta_Codigo,fc_Cuenta_Bancaria.cta_codigo_tgn,fc_Cuenta_Bancaria.cta_descripcion_larga," & _
'        " Pago_Detalle.fecha_Pago,Pago_Detalle.GES_GESTION from Pagos,Pago_Detalle,fc_Cuenta_Bancaria where " & _
'        " Pagos.Ges_Gestion = Pago_Detalle.Ges_Gestion and Pagos.Org_Codigo=Pago_Detalle.Org_Codigo and  Pagos.Codigo_Pago=Pago_Detalle.Codigo_Pago " & _
'        " and Pagos.Tipo_Comp= 'PCE' and Pagos.Codigo_Pago = '" & P_codigo_pago & "' " & _
'        " and fc_Cuenta_Bancaria.Cta_Codigo=Pago_Detalle.Cta_Codigo ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
' 'MsgBox recSetAuxActualizar1!Ges_Gestion
'
'
'      '  If recSetAuxActualizar1.RecordCount > 0 Then
'
''        Else
''        MsgBox "No existen cuentas asociadas................"
''        End If
'
'            If recsetAdicion.State = 1 Then recsetAdicion.Close
'            recsetAdicion.Open " select * from Co_Comprobante_M  ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'             recsetAdicion.AddNew
'
'              '  recsetAdicion!Cod_Comp = Cont_Comp
'                recsetAdicion!Tipo_Comp = recSetAux!Tipo_Comp
'                recsetAdicion!Ges_Gestion = recSetAuxActualizar1!Ges_Gestion
'                recsetAdicion!Fecha_A = recSetAuxActualizar1!fecha_pago
'
'             Select Case recSetAux!Tipo_Comp
'                Case "PCE"
'                    recsetAdicion!codigo_beneficiario = recSetAux!codigo_beneficiario
'                Case "PCO"
'
'             End Select
'
'                recsetAdicion!Glosa = recSetAux!Glosa
'                recsetAdicion!codigo_documento = recSetAux!codigo_documento
'                recsetAdicion!num_respaldo = recSetAux!num_respaldo
'                recsetAdicion!Status = recSetAux!Status
''                recsetAdicion.Update
'
'
'        '********* adicion Debitos creditos
'            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'            recSetAuxActualizar.Open " select * from Co_Diario  ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            recSetAuxActualizar.AddNew
'
'                'recSetAuxActualizar!Cod_Comp = Cont_Comp
'                recSetAuxActualizar!Tipo_Comp = recSetAux!Tipo_Comp
'                recSetAuxActualizar!Cod_Comp_C = recSetAux!Cod_Comp
'                recSetAuxActualizar!D_Cuenta = recSetAux!H_Cuenta
'                recSetAuxActualizar!D_SubCta1 = recSetAux!H_SubCta1
'                recSetAuxActualizar!D_SubCta2 = recSetAux!H_SubCta2
'
'                recSetAuxActualizar!d_Aux1 = recSetAux!h_Aux1
'                recSetAuxActualizar!d_Aux2 = recSetAux!h_Aux2
'                recSetAuxActualizar!d_Aux3 = recSetAux!h_Aux3
'
'                recSetAuxActualizar!D_cta_Larga = recSetAux!H_cta_Larga
'                recSetAuxActualizar!d_des_Larga = recSetAux!H_des_Larga
'
'                recSetAuxActualizar!D_MontoBs = recSetAux!h_MontoBs
'                recSetAuxActualizar!D_MontoDl = recSetAux!h_MontoDl
'                recSetAuxActualizar!D_Cambio = recSetAux!h_Cambio
'              Select Case recSetAuxActualizar1!Fte_Codigo
'               Case "10", "41"
'                recSetAuxActualizar!H_Cuenta = "1111"
'                recSetAuxActualizar!H_SubCta1 = "02"
'                recSetAuxActualizar!H_SubCta2 = "01"
'
'                recSetAuxActualizar!h_Aux1 = "02"
'                recSetAuxActualizar!h_Aux2 = "00"
'                recSetAuxActualizar!h_Aux3 = "00"
'
'               Case "70", "43"
'                recSetAuxActualizar!H_Cuenta = "1111"
'                recSetAuxActualizar!H_SubCta1 = "02"
'                recSetAuxActualizar!H_SubCta2 = "02"
'
'                recSetAuxActualizar!h_Aux1 = "02"
'                recSetAuxActualizar!h_Aux2 = "00"
'                recSetAuxActualizar!h_Aux3 = "00"
'
'              Case "80"
'                recSetAuxActualizar!H_Cuenta = "1111"
'                recSetAuxActualizar!H_SubCta1 = "02"
'                recSetAuxActualizar!H_SubCta2 = "03"
'
'                recSetAuxActualizar!h_Aux1 = "02"
'                recSetAuxActualizar!h_Aux2 = "00"
'                recSetAuxActualizar!h_Aux3 = "00"
'             End Select
'
'                recSetAuxActualizar!H_cta_Larga = recSetAuxActualizar1!cta_Codigo
'                recSetAuxActualizar!H_des_Larga = recSetAuxActualizar1!cta_Descripcion_larga
'
'                recSetAuxActualizar!h_MontoBs = recSetAux!h_MontoBs
'                recSetAuxActualizar!h_MontoDl = recSetAux!h_MontoDl
'                recSetAuxActualizar!h_Cambio = recSetAux!h_Cambio
'''************ GENERA EL CODIGO DE COMPROBANTE**********
'
'            Set recSetGenera = New ADODB.Recordset
'            recSetGenera.CursorLocation = adUseClient
'            If recSetGenera.State = 1 Then recSetGenera.Close
'            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'            If recSetGenera.RecordCount > 0 Then
'             Cont_Comp = Val(recSetGenera!Numero_correlativo)
'             Cont_Comp = Cont_Comp + 1
'             recSetGenera!Numero_correlativo = Trim(Str(Cont_Comp))
'
'
'
'''************TERMINA GENERACION DE COMPROBANTE********
'                recsetAdicion!Cod_Comp = Cont_Comp
'                recSetAuxActualizar!Cod_Comp = Cont_Comp
'                recsetAdicion.Update
'                recSetAuxActualizar.Update
'                recSetGenera.Update
'                'LblTitulo.Caption = "Contra cuenta completada"
'             End If
'
''         Else
''         MsgBox "No existen cuentas asociadas................"
''         End If
'
'
'Else
'MsgBox "No existen registros"
'
'End If
'db.CommitTrans
'MsgBox "Contabilizacion exitosa..............."
'
'
'
'
'Exit Sub
'Asiento_Err:
'    MsgBox "Error al generar contra cuenta"
'    db.RollbackTrans
'    'CmdAgregarDetalle.Enabled = True
'    'Cmd_Modificar.Enabled = True
'    'Cmd_Aprobar.Enabled = True
'    'CmdSalir.Enabled = True
'    'Cmd_GrabaM.Enabled = True
'    'Cmd_Cancelar.Enabled = True
'    'Cmd_Copiar.Enabled = True
'
'End Sub


'''Private Sub Cmd_contabiliza(P_codigo_pago As String, P_org_codigo As String, P_ges_gestion As String)
''''On Error GoTo Asiento_Err
'''db.BeginTrans
'''
'''MsgBox "contabilizando"
'''
'''Set recsetaux = New ADODB.Recordset
'''recsetaux.CursorLocation = adUseClient
'''
'''If recsetaux.State = 1 Then recsetaux.Close
'''recsetaux.Open " SELECT  distinct Co_Comprobante_M.Cod_Comp,Co_Comprobante_M.Tipo_Comp,cO_comprobante_M.Num_Respaldo," & _
'''" Co_Comprobante_M.codigo_beneficiario,Co_Comprobante_M.codigo_Documento,Co_Comprobante_M.Fecha_A,Co_Comprobante_M.ges_gestion," & _
'''" Co_Comprobante_M.Glosa,status,CO_Diario.D_Aux1,CO_Diario.D_Aux2, CO_Diario.D_Aux3,Co_Diario.d_Cta_Larga,Co_Diario.D_Des_Larga,Co_Comprobante_M.cod_Comp, " & _
'''" CO_Diario.D_Cuenta, CO_Diario.D_Subcta1,CO_Diario.D_SubCta2, CO_Diario.D_Nombre,CO_Diario.D_MontoBs,D_Cambio,H_Cambio," & _
'''" CO_Diario.D_MontoDl,CO_Diario.H_SubCta1, CO_Diario.H_SubCta2,CO_Diario.H_Aux1, CO_Diario.H_Aux2,Co_Diario.H_Cta_Larga,Co_Diario.H_Des_Larga," & _
'''" CO_Diario.H_Aux3,CO_Diario.H_Nombre, CO_Diario.H_MontoBs,CO_Diario.H_Montodl,CO_Diario.H_Cuenta " & _
'''" From CO_Diario,CO_Comprobante_M WHERE CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp AND co_Comprobante_M.Cod_Comp=" & Trim(P_codigo_pago) & _
'''" and co_Comprobante_M.Tipo_Comp='PCE' and CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp and status='S' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
''''If recSetAux.RecordCount > 0 Then
''' 'MsgBox recSetAux!Cod_Comp
'''  Set recSetAuxActualizar1 = New ADODB.Recordset
'''  recSetAuxActualizar1.CursorLocation = adUseClient
'''  If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'''  recSetAuxActualizar1.Open " select distinct fc_Cuenta_Bancaria.fte_codigo,tipo_comp,Pagos.ORg_Codigo, " & _
'''  " Pago_Detalle.ges_Gestion,pago_Detalle.cta_Codigo,fc_Cuenta_Bancaria.cta_codigo_tgn,fc_Cuenta_Bancaria.cta_descripcion_larga," & _
'''  " Pago_Detalle.fecha_Pago from Pagos,Pago_Detalle,fc_Cuenta_Bancaria where " & _
'''  " Pagos.Ges_Gestion = Pago_Detalle.Ges_Gestion and Pagos.Org_Codigo=Pago_Detalle.Org_Codigo and  Pagos.Codigo_Pago=Pago_Detalle.Codigo_Pago " & _
'''  " and Pagos.Tipo_Comp= 'PCE' and Pagos.Codigo_Pago = '" & P_codigo_pago & "' and  Pagos.Org_Codigo='999' and " & _
'''  " fc_Cuenta_Bancaria.Cta_Codigo=Pago_Detalle.Cta_Codigo ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
'''If recSetAuxActualizar1.RecordCount > 0 Then recSetAuxActualizar1.MoveFirst
'''While Not (recSetAuxActualizar1.EOF)
'''v_Fte = recSetAuxActualizar1!fte_codigo
'''
'''If recsetAdicion.State = 1 Then recsetAdicion.Close
'''recsetAdicion.Open " select * from Co_Comprobante_M  where cod_Trans=" & P_codigo_pago & " and Org_Codigo='999' and tipo_Comp='PCE' and Ges_Gestion='" & P_ges_gestion & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
''''MsgBox recsetAdicion!Cod_Comp
'''
''''recsetAdicion.cod_Comp
''''recsetAdicion.tipo_Comp
''''recsetAdicion.Org_Codigo
'''If Not recsetAdicion.BOF Then recsetAdicion.MoveFirst
'''If (recsetAdicion.BOF) And (recsetAdicion.EOF) Then
'''
''''************* GENERA EL CODIGO DE COMPROBANTE**********
'''            Set recSetGenera = New ADODB.Recordset
'''            recSetGenera.CursorLocation = adUseClient
'''            If recSetGenera.State = 1 Then recSetGenera.Close
'''            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'''            If recSetGenera.RecordCount > 0 Then
'''                Cont_Comp = Val(recSetGenera!numero_correlativo)
'''                Cont_Comp = Cont_Comp + 1
'''                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'''                recSetGenera.Update
'''            End If
'''            If recSetGenera.State = 1 Then recSetGenera.Close
''''************TERMINA GENERACION DE COMPROBANTE********
'''
'''  recsetAdicion.AddNew
'''
'''
''' ' recsetAdicion!usr_Usuario = usuario2
'''  recsetAdicion!fecha_registro = Date
'''  recsetAdicion!hora_registro = Format(Time, "hh:mm:ss")
'''
'''
'''  recsetAdicion!Cod_Comp = Cont_Comp
'''  recsetAdicion!Cod_trans = recsetaux!Cod_Comp
'''  recsetAdicion!Cod_Trans_Detalle = "1"
'''  recsetAdicion!org_codigo = P_org_codigo
'''  recsetAdicion!tipo_comp = "PCC" 'recsetaux!tipo_comp
'''  recsetAdicion!Ges_gestion = recSetAuxActualizar1!Ges_gestion
'''  recsetAdicion!fecha_A = CDate(recSetAuxActualizar1!fecha_pago)
'''  Select Case recsetaux!tipo_comp
'''      Case "PCE"
'''         recsetAdicion!Codigo_beneficiario = recsetaux!Codigo_beneficiario
'''      Case "PCO"
'''
'''  End Select
'''  recsetAdicion!glosa = recsetaux!glosa
'''  recsetAdicion!codigo_documento = recsetaux!codigo_documento
'''  recsetAdicion!num_respaldo = recsetaux!num_respaldo
'''  recsetAdicion!Status = recsetaux!Status
'''  recsetAdicion!usr_usuario = GlUsuario
'''  recsetAdicion!fecha_registro = Format(Date, "dd/mm/yyyy")
'''  recsetAdicion!hora_registro = Format(Time, "hh:mm:ss")
'''  recsetAdicion.Update
'''  If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''        '********* adicion Debitos creditos
''' Set recSetAuxActualizar = New ADODB.Recordset
''' If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
''' recSetAuxActualizar.Open " select * from Co_Diario where  cod_Comp_c=" & recsetaux!Cod_Comp, db, adOpenDynamic, adLockOptimistic, adCmdText
''' If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
'''    recSetAuxActualizar.AddNew
'''    recSetAuxActualizar!usr_usuario = GlUsuario
'''    recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'''    recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
'''    'recsetAdicion!Cod_Comp = Cont_Comp
'''    recSetAuxActualizar!Cod_Comp = Cont_Comp
'''    recSetAuxActualizar!tipo_comp = "PCC" 'recsetaux!tipo_comp
'''    recSetAuxActualizar!Cod_Comp_C = recsetaux!Cod_Comp
'''    recSetAuxActualizar!d_cuenta = recsetaux!h_cuenta
'''    recSetAuxActualizar!d_subcta1 = recsetaux!h_subcta1
'''    recSetAuxActualizar!d_subcta2 = recsetaux!h_subcta2
'''    recSetAuxActualizar!d_Aux1 = recsetaux!h_Aux1
'''    recSetAuxActualizar!d_Aux2 = recsetaux!h_Aux2
'''    recSetAuxActualizar!d_Aux3 = recsetaux!h_Aux3
'''    recSetAuxActualizar!d_cta_larga = recsetaux!h_cta_larga
'''    recSetAuxActualizar!d_des_Larga = IIf(IsNull(recsetaux!h_des_Larga), " ", Trim(recsetaux!h_des_Larga))
'''    recSetAuxActualizar!d_montobs = recsetaux!h_montoBs
'''    recSetAuxActualizar!d_montoDl = recsetaux!h_montoDl
'''    recSetAuxActualizar!d_Cambio = recsetaux!h_Cambio
'''    Select Case v_Fte
'''        Case "10", "41"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "01"
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''        Case "70", "43"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "02"
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''        Case "80"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "03"
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''    End Select
'''    recSetAuxActualizar!h_cta_larga = recSetAuxActualizar1!cta_codigo
'''    recSetAuxActualizar!h_des_Larga = IIf(IsNull(recSetAuxActualizar1!cta_descripcion_larga), " ", recSetAuxActualizar1!cta_descripcion_larga)
'''    recSetAuxActualizar!h_montoBs = recsetaux!h_montoBs
'''    recSetAuxActualizar!h_montoDl = recsetaux!h_montoDl
'''    recSetAuxActualizar!h_Cambio = recsetaux!h_Cambio
'''    recSetAuxActualizar!usr_usuario = GlUsuario
'''    recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'''    recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
'''
'''    recSetAuxActualizar.Update
'''    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''
'''
''''''************ GENERA EL CODIGO DE COMPROBANTE**********
''''
''''            Set recSetGenera = New ADODB.Recordset
''''            recSetGenera.CursorLocation = adUseClient
''''            If recSetGenera.State = 1 Then recSetGenera.Close
''''            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
''''            If recSetGenera.RecordCount > 0 Then
''''             Cont_Comp = Val(recSetGenera!Numero_correlativo)
''''             Cont_Comp = Cont_Comp + 1
''''             recSetGenera!Numero_correlativo = Trim(Str(Cont_Comp))
''''
''''
''''
''''''************TERMINA GENERACION DE COMPROBANTE********
'''
''''                recSetAuxActualizar!Cod_Comp = Cont_Comp
''''                recsetAdicion.Update
''''                recSetAuxActualizar.Update
''''               recSetGenera.Update
'''                'LblTitulo.Caption = "Contra cuenta completada"
'''         End If 'Adicion del diario
'''       Else
'''         MsgBox "Ya fue contabilizado anteriormente"
'''' ******Modifica registro existente
'''
'''        'recsetAdicion!usr_Usuario = usuario2
'''        recsetAdicion!fecha_registro = Date
'''        recsetAdicion!hora_registro = Format(Time, "hh:mm:ss")
'''
'''        Cont_Comp = recsetAdicion!Cod_Comp
'''        recsetAdicion!Cod_Comp = Cont_Comp
'''        recsetAdicion!Cod_trans = recsetaux!Cod_Comp
'''        recsetAdicion!Cod_Trans_Detalle = "1"
'''        recsetAdicion!org_codigo = recSetAuxActualizar1!org_codigo
'''        recsetAdicion!tipo_comp = recsetaux!tipo_comp
'''        recsetAdicion!Ges_gestion = recSetAuxActualizar1!Ges_gestion
'''        recsetAdicion!fecha_A = CDate(recSetAuxActualizar1!fecha_pago)
'''        Select Case recsetaux!tipo_comp
'''            Case "PCE"
'''               recsetAdicion!Codigo_beneficiario = recsetaux!Codigo_beneficiario
'''            Case "PCO"
'''
'''        End Select
'''
'''        recsetAdicion!glosa = recsetaux!glosa
'''        recsetAdicion!codigo_documento = recsetaux!codigo_documento
'''        recsetAdicion!num_respaldo = recsetaux!num_respaldo
'''        recsetAdicion!Status = recsetaux!Status
'''        recsetAdicion!usr_usuario = GlUsuario
'''        recsetAdicion!fecha_registro = Format(Date, "dd/mm/yyyy")
'''        recsetAdicion!hora_registro = Format(Time, "hh:mm:ss")
'''        recsetAdicion.Update
'''        If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''
''''******Termina de Modificar
'''
''''******Modifica el Diario
'''Set recSetAuxActualizar = New ADODB.Recordset
'''        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''        recSetAuxActualizar.Open " select * from Co_Diario where  cod_Comp=" & Cont_Comp, db, adOpenDynamic, adLockOptimistic, adCmdText
'''        If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
'''          recSetAuxActualizar.AddNew
'''          recSetAuxActualizar!Cod_Comp = Cont_Comp
'''          recSetAuxActualizar!tipo_comp = recsetaux!tipo_comp
'''        Else
'''          If (Not recSetAuxActualizar.BOF) Then recSetAuxActualizar.MoveFirst
'''        End If
'''
'''    recSetAuxActualizar!usr_usuario = GlUsuario
'''    recSetAuxActualizar!fecha_registro = Format(Date, "dd/mm/yyyy")
'''    recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
'''
'''    'recsetAdicion!Cod_Comp = Cont_Comp
'''
''''    recSetAuxActualizar!Cod_Comp = Cont_Comp
''''    recSetAuxActualizar!Tipo_comp = recSetAux!Tipo_comp
'''    recSetAuxActualizar!Cod_Comp_C = recsetaux!Cod_Comp
'''    recSetAuxActualizar!d_cuenta = recsetaux!h_cuenta
'''    recSetAuxActualizar!d_subcta1 = recsetaux!h_subcta1
'''    recSetAuxActualizar!d_subcta2 = recsetaux!h_subcta2
'''
'''    recSetAuxActualizar!d_Aux1 = recsetaux!h_Aux1
'''    recSetAuxActualizar!d_Aux2 = recsetaux!h_Aux2
'''    recSetAuxActualizar!d_Aux3 = recsetaux!h_Aux3
'''
'''    recSetAuxActualizar!d_cta_larga = recsetaux!h_cta_larga
'''    recSetAuxActualizar!d_des_Larga = IIf(IsNull(recsetaux!h_des_Larga), " ", recsetaux!h_des_Larga)
'''    recSetAuxActualizar!d_montobs = recsetaux!h_montoBs
'''    recSetAuxActualizar!d_montoDl = recsetaux!h_montoDl
'''    recSetAuxActualizar!d_Cambio = recsetaux!h_Cambio
'''
'''    Select Case v_Fte
'''
'''               Case "10", "41"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "01"
'''
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''
'''               Case "70", "43"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "02"
'''
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''
'''              Case "80"
'''                recSetAuxActualizar!h_cuenta = "1111"
'''                recSetAuxActualizar!h_subcta1 = "02"
'''                recSetAuxActualizar!h_subcta2 = "03"
'''
'''                recSetAuxActualizar!h_Aux1 = "02"
'''                recSetAuxActualizar!h_Aux2 = "00"
'''                recSetAuxActualizar!h_Aux3 = "00"
'''             End Select
'''
'''                recSetAuxActualizar!h_cta_larga = recSetAuxActualizar1!cta_codigo
'''                recSetAuxActualizar!h_des_Larga = IIf(IsNull(recSetAuxActualizar1!cta_descripcion_larga), "", recSetAuxActualizar1!cta_descripcion_larga)
'''
'''                recSetAuxActualizar!h_montoBs = recsetaux!h_montoBs
'''                recSetAuxActualizar!h_montoDl = recsetaux!h_montoDl
'''                recSetAuxActualizar!h_Cambio = recsetaux!h_Cambio
'''             recSetAuxActualizar.Update
'''             If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''
'''       End If '*****Existe comprobante modificaion
'''
''''******Termina de Modificar el diario
'''
''''         Else
''''         MsgBox "No existen cuentas asociadas................"
''''         End If
'''recSetAuxActualizar1.MoveNext
'''
'''Wend
'''
''''Else
''''MsgBox "No existen registros"
''''
''''End If
'''db.CommitTrans
'''MsgBox "Contabilizacion exitosa...............", vbInformation + vbDefaultButton1, "CONTABILIZACION"
'''
'''
'''
'''
'''Exit Sub
'''Asiento_Err:
'''    MsgBox "Error al generar contra cuenta"
'''    db.RollbackTrans
'''    'CmdAgregarDetalle.Enabled = True
'''    'Cmd_Modificar.Enabled = True
'''    'Cmd_Aprobar.Enabled = True
'''    'CmdSalir.Enabled = True
'''    'Cmd_GrabaM.Enabled = True
'''    'Cmd_Cancelar.Enabled = True
'''    'Cmd_Copiar.Enabled = True
'''
'''End Sub


'*************Comentario anterior************************

'Private Sub Cmd_Pagado(P_codigo_pago As String, P_codigo_pago_detalle As String, P_org_codigo As String, P_ges_gestion As String)
'Dim sw As Boolean
'Dim Sw_Fuente As Boolean
'Dim Cont_Comp As Long
'Dim aux_T As String
'
'On Error GoTo errorPag
'
'db.BeginTrans
''While Not (recSetAuxcomp1.EOF)
'
'MsgBox "Contabilizando............"
'        Set recSetAuxcomp = New ADODB.Recordset
'        recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'
'        If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'        recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
'        " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
'        " pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and   " & _
'        " pago_Detalle.Org_codigo= '" & P_org_codigo & "' and  pago_detalle.Ges_Gestion='" & P_ges_gestion & "' and pago_detalle.codigo_Pago='" & P_codigo_pago & "' and " & _
'        " pago_detalle.Ges_Gestion = pagos.Ges_Gestion and pagos.estado_pagado ='S'  and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        'pagos.Org_codigo='" & rsCheque!cod_org & "' and
'        'pago_detalle.estado_aprobacion ='A' and pago_detalle.Ges_Gestion='" & rsCheque!Ges_Gestion & "' and pago_detalle.codigo_Pago='" & rsCheque!Numero_comprobante & "'
'        'and  pagos.estado_Pagado='S'  AND Pagos.Tipo_comp='PAC'
'        'AND pago_detalle.estado_aprobacion = 'A'
'        'MsgBox pago_detalle.estado_aprobacion
'        If recSetAuxcomp.RecordCount > 0 Then
'        recSetAuxcomp.MoveFirst
'        End If
'
'While Not (recSetAuxcomp.EOF)
'        '************Abrimos un record set para adicionar datos*********************'
'        Set recSetAuxActualizar = New ADODB.Recordset
'        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'        recSetAuxActualizar.Open " select * from CO_Comprobante_M ", db, adOpenDynamic, adLockOptimistic, adCmdText
'
'        Set recSetAuxActualizar1 = New ADODB.Recordset
'        If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'        recSetAuxActualizar1.Open " select * from CO_Diario ", db, adOpenDynamic, adLockOptimistic, adCmdText
'        Dim Aux_Cont As String
'
'        aux_T = "select * from Co_comprobante_M"
'
'        'While Not (recSetAuxcomp.EOF)
'
'        If Not Buscar(aux_T, recSetAuxcomp!Codigo_Pago, recSetAuxcomp!Org_Codigo, recSetAuxcomp!Ges_gestion, "PAC", recSetAuxcomp!codigo_Pago_detalle) Then
'
'            Select Case recSetAuxcomp!Fte_Codigo
'
'            Case "10", "41"
'
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst= 'PAG' and CC_Cuenta_H.Inst= 'PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=1 AND " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'           Case "70", "43"
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=2 AND " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'            Case "80"
'            Set recSetPartida = New ADODB.Recordset
'            recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'            If recSetPartida.State = 1 Then recSetPartida.Close
'            recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H, CC_Cuentas_D" & _
'            " WHERE   CC_Cuenta_H.Par_I = CC_Cuentas_D.Par_I AND CC_Cuenta_H.Par_F = CC_Cuentas_D.Par_F AND CC_Cuentas_D.Inst='PAG' and CC_Cuenta_H.Inst='PAG' and " & _
'            " CC_Cuentas_D.O_C=CC_Cuenta_H.O_C and CC_Cuenta_H.O_C=3 and  " & _
'            " cc_Cuenta_H.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'            Sw_Fuente = True
'
'            Case Else
'            'MsgBox "No esta asociado a ninguna fuente ... partida no relacionada "
'            Sw_Fuente = False
'
'            End Select
'          If Sw_Fuente Then
'
'            recSetAuxActualizar.AddNew
'            recSetAuxActualizar1.AddNew
'            'recSetAuxActualizar!Cod_Comp = Cont_Comp
'            recSetAuxActualizar!Cod_Trans = recSetAuxcomp!Codigo_Pago
'            recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_Pago_detalle
'            recSetAuxActualizar!Org_Codigo = recSetAuxcomp!Org_Codigo
'            recSetAuxActualizar!Codigo_Beneficiario = recSetAuxcomp!Codigo_Beneficiario
'            recSetAuxActualizar!Ges_gestion = recSetAuxcomp!Ges_gestion
'            recSetAuxActualizar!Num_respaldo = recSetAuxcomp!codigo_orden
'            recSetAuxActualizar!Codigo_documento = recSetAuxcomp!Codigo_documento
'
'            recSetAuxActualizar!Fecha_A = recSetAuxcomp!fecha_pago
'            recSetAuxActualizar!Glosa = recSetAuxcomp!Justificacion
'            'recSetAuxActualizar!codigo_solicitud = recSetAuxcomp!codigo_solicitud
'            recSetAuxActualizar!tipo_Comp = "PAC"
'
'            recSetAuxActualizar!Status = "S"
'            recSetAuxActualizar1!tipo_Comp = "PAC"
'            recSetAuxActualizar1!D_Cuenta = recSetPartida!cuenta
'            recSetAuxActualizar1!D_Nombre = recSetPartida!NombreCta
'            recSetAuxActualizar1!D_SubCta1 = recSetPartida!Subcta1
'            recSetAuxActualizar1!D_SubCta2 = recSetPartida!Subcta2
'            recSetAuxActualizar1!d_Aux1 = recSetPartida!Aux1
'            recSetAuxActualizar1!d_Aux2 = recSetPartida!Aux2
'            recSetAuxActualizar1!d_Aux3 = recSetPartida!Aux3
'
'        '************* CONTABILIZA AUXILIAARES DEBITO
'            Select Case recSetPartida!Aux1
'            Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_Beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_Larga = recsetAdicion!Codigo_Beneficiario
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'
'            Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!d_cta_Larga = recsetAdicion!cta_Codigo
'                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!Cta_descripcion_larga
'
'            Case Else
'            End Select
'        ''****************** finaliza sesion de auxiliares
'
'
'            recSetAuxActualizar1!h_Aux1 = recSetPartida!h_Aux1
'            recSetAuxActualizar1!h_Aux2 = recSetPartida!h_Aux2
'            recSetAuxActualizar1!h_Aux3 = recSetPartida!h_Aux3
'
'        '************* CONTABILIZA AUXILIAARES DEBITO
'
'            Select Case recSetPartida!h_Aux1
'            Case "01"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_Beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_Larga = recsetAdicion!Codigo_Beneficiario
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'
'            Case "02"
'                    Set recsetAdicion = New ADODB.Recordset
'                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'
'                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    'recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    recSetAuxActualizar1!h_cta_Larga = recsetAdicion!cta_Codigo
'                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!Cta_descripcion_larga
'
'            Case Else
'            End Select
'        ''****************** finaliza sesion de auxiliares
'
'            recSetAuxActualizar1!H_Cuenta = recSetPartida!H_Cuenta
'            recSetAuxActualizar1!H_Nombre = recSetPartida!H_NombCta
'            recSetAuxActualizar1!H_SubCta1 = recSetPartida!H_SubCta1
'            recSetAuxActualizar1!H_SubCta2 = recSetPartida!H_SubCta2
'            recSetAuxActualizar1!D_MontoBs = recSetAuxcomp!Monto_Bolivianos
'            recSetAuxActualizar1!D_MontoDl = recSetAuxcomp!Monto_Dolares
'            recSetAuxActualizar1!D_MontoDl = recSetAuxcomp!Monto_Dolares
'            recSetAuxActualizar1!D_Cambio = recSetAuxcomp!tipo_cambio
'
'            recSetAuxActualizar1!h_MontoBs = recSetAuxcomp!Monto_Bolivianos
'            recSetAuxActualizar1!h_MontoDl = recSetAuxcomp!Monto_Dolares
'            recSetAuxActualizar1!h_MontoDl = recSetAuxcomp!Monto_Dolares
'            recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'        '************* GENERA EL CODIGO DE COMPROBANTE**********
'
'                    Set recSetGenera = New ADODB.Recordset
'                    recSetGenera.CursorLocation = adUseClient
'                    If recSetGenera.State = 1 Then recSetGenera.Close
'                    recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'                    If recSetGenera.RecordCount > 0 Then
'                     Cont_Comp = Val(recSetGenera!Numero_correlativo)
'                     Cont_Comp = Cont_Comp + 1
'                     recSetGenera!Numero_correlativo = Trim(Str(Cont_Comp))
'
'        '************TERMINA GENERACION DE COMPROBANTE********
'                     recSetAuxActualizar!Cod_Comp = Cont_Comp
'                     recSetAuxActualizar1!Cod_Comp = Cont_Comp
'                     recSetAuxActualizar1.Update
'                     recSetAuxActualizar.Update
'                     recSetGenera.Update
'
'                    End If
'
'           Else
'                MsgBox "No esta asociado a ninguna fuente ...  "
'
'           End If
'        Else
'            MsgBox "Existe registro....."
'        End If
'            'Cont_Comp = Cont_Comp + 1
'            recSetAuxcomp.MoveNext
'Wend
'''Unload Me
''recSetGenera!Numero_correlativo = Cont_Comp
''recSetGenera.Update
'db.CommitTrans
''MsgBox "Contabilizacion exitosa...... "
''Cerrar record
'    Set recSetAuxcomp = New ADODB.Recordset
'    recSetAuxcomp.CursorLocation = adUseClient
'    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'
'    Set recSetAuxActualizar = New ADODB.Recordset
'    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'
'    Set recSetAuxActualizar1 = New ADODB.Recordset
'    If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'
'    Set recSetPartida = New ADODB.Recordset
'    recSetPartida.CursorLocation = adUseClient
'    If recSetPartida.State = 1 Then recSetPartida.Close
'
'
'Exit Sub
'errorPag:
'db.RollbackTrans
'MsgBox "No se contabilizó ... "
'
'End Sub

'*******************fin de comentario

'''Private Sub Cmd_Pagado(P_codigo_pago As String, P_codigo_pago_detalle As String, P_org_codigo As String, P_ges_gestion As String)
'''Dim sw As Boolean
'''
'''Dim Sw_Fuente As Boolean
'''Dim Cont_Comp As Long
'''Dim aux_T As String
'''
'''Dim v_Cuenta As String
'''Dim v_SubCta1 As String
'''Dim v_SubCta2 As String
'''Dim v_NombreCta As String
'''Dim v_H_Cuenta As String
'''Dim v_H_SubCta1 As String
'''Dim v_H_SubCta2 As String
'''Dim v_H_NombCta As String
'''Dim v_Aux1 As String
'''Dim v_Aux2 As String
'''Dim v_Aux3 As String
'''Dim v_H_Aux1 As String
'''Dim v_H_Aux2 As String
'''Dim v_H_Aux3 As String
'''Dim Aux_Cont As String
'''Dim rstipopy As ADODB.Recordset
'''Set rstipopy = New ADODB.Recordset
'''
''''On Error GoTo errorPag
'''
'''db.BeginTrans
'''        MsgBox "Contabilizando............", vbInformation + vbOKOnly, "Contabilización"
'''        Set recSetAuxcomp = New ADODB.Recordset
'''        recSetAuxcomp.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'''
'''    If Me.DtCCuentaOrigen.Text = "" Then
'''            MsgBox "ERROR, NO SE CONTABILIZO", vbCritical + vbDefaultButton1 + vbOKOnly, "Contabilicación"
'''            Exit Sub
'''    End If
'''        If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'''        recSetAuxcomp.Open "SELECT distinct pago_detalle.codigo_Pago,pagos.codigo_solicitud,pago_detalle.codigo_Pago_detalle,Pagos.Fte_Codigo,pagos.Ges_Gestion,Estado_Pagado,Pago_Detalle.Cta_Codigo,Pago_Detalle.tipo_cambio," & _
'''        " Pago_Detalle.Codigo_Beneficiario,pagos.Justificacion,pago_detalle.fecha_pago,pago_detalle.par_codigo,pago_detalle.Monto_Bolivianos,estado_Devengado,Pagos.Org_Codigo,Pagos.Codigo_Orden,Pagos.Codigo_Documento," & _
'''        " pago_detalle.pro_programa, pago_detalle.pro_subprograma, pago_detalle.pro_proyecto, pago_detalle.pro_actividad, " & _
'''        " pago_detalle.Monto_Dolares,pago_detalle.estado_aprobacion From pago_detalle,pagos Where pago_detalle.codigo_Pago = pagos.codigo_Pago and pago_detalle.Org_Codigo = pagos.Org_codigo and   " & _
'''        " pago_Detalle.Org_codigo= '" & P_org_codigo & "' and  pago_detalle.Ges_Gestion='" & P_ges_gestion & "' and pago_detalle.codigo_Pago='" & P_codigo_pago & "' and " & _
'''        " pago_detalle.Ges_Gestion = pagos.Ges_Gestion  and pago_detalle.codigo_pago_detalle='" & P_codigo_pago_detalle & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        If recSetAuxcomp.RecordCount > 0 Then
'''            recSetAuxcomp.MoveFirst
'''        Else
'''            MsgBox "ERROR EN LA CONTABILIZACION", vbCritical + vbDefaultButton1, "Contabilización"
'''            Exit Sub
'''        End If
'''While Not (recSetAuxcomp.EOF)
''''VERIFICA FUENTE
'''    If rstipopy.State = 1 Then rstipopy.Close
'''    Dim sqlpy  As String
'''    Dim tipopy As String
'''    rstipopy.Open "select tipo_proyecto from fc_estructura_programatica where Pro_programa='" & recSetAuxcomp!Pro_programa & "' and  Pro_subprograma='" & recSetAuxcomp!Pro_subprograma & "' and Pro_proyecto='" & recSetAuxcomp!Pro_proyecto & "' and Pro_actividad='" & recSetAuxcomp!Pro_actividad & "'", db, adOpenKeyset, adLockReadOnly
'''    If rstipopy.RecordCount <> 0 Then
'''        tipopy = rstipopy!tipo_proyecto
'''    Else
'''        MsgBox "Error en la contabilización. No se encontró la Categoria Programática Asociada", vbExclamation + vbDefaultButton1, "Contabilización"
'''        Exit Sub
'''    End If
''''VERIFICA FUENTE
'''    Select Case recSetAuxcomp!fte_codigo
'''    Case "10", "41"
'''        Set recSetPartida = New ADODB.Recordset
'''        recSetPartida.CursorLocation = adUseClient
'''
'''        Select Case tipopy
'''            Case "S"
'''                    If recSetPartida.State = 1 Then recSetPartida.Close
'''                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                            " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PSP' and CC_Cuenta_H1.Inst= 'PSP' and " & _
'''                            " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
'''                            " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
'''            Case "F"
'''                    If recSetPartida.State = 1 Then recSetPartida.Close
'''                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                            " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PFP' and CC_Cuenta_H1.Inst= 'PFP' and " & _
'''                            " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
'''                            " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
'''            Case "N"
'''                    If recSetPartida.State = 1 Then recSetPartida.Close
'''                    recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                            " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst= 'PAG' and CC_Cuenta_H1.Inst= 'PAG' and " & _
'''                            " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=1 AND " & _
'''                            " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        End Select
'''        If recSetPartida.RecordCount > 0 Then
'''            Sw_Fuente = True
'''        Else
'''            Sw_Fuente = False
'''        End If
'''
'''    Case "70", "43"
'''        Set recSetPartida = New ADODB.Recordset
'''        recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'''
'''        Select Case tipopy
'''            Case "S"
'''                If recSetPartida.State = 1 Then recSetPartida.Close
'''                recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PSP' and CC_Cuenta_H1.Inst='PSP' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''            Case "F"
'''                If recSetPartida.State = 1 Then recSetPartida.Close
'''                recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PFP' and CC_Cuenta_H1.Inst='PFP' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
'''            Case "N"
'''                If recSetPartida.State = 1 Then recSetPartida.Close
'''                recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3 From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=2 AND " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        End Select
'''        If recSetPartida.RecordCount > 0 Then
'''                     Sw_Fuente = True
'''        Else
'''                     Sw_Fuente = False
'''        End If
'''    Case "80"
'''        Set recSetPartida = New ADODB.Recordset
'''          recSetPartida.CursorLocation = adUseClient  ' Use client cursor to enable AbsolutePosition property.
'''        Select Case tipopy
'''        Case "S"
'''              If recSetPartida.State = 1 Then recSetPartida.Close
'''              recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PSP' and CC_Cuenta_H1.Inst='PSP' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''
'''        Case "F"
'''              If recSetPartida.State = 1 Then recSetPartida.Close
'''              recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PFP' and CC_Cuenta_H1.Inst='PFP' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        Case "N"
'''              If recSetPartida.State = 1 Then recSetPartida.Close
'''              recSetPartida.Open "SELECT Distinct Cuenta,SubCta1,SubCta2,NombreCta,H_Cuenta,H_SubCta1,H_SubCta2,H_NombCta,Aux1,Aux2,Aux3,H_Aux1,H_Aux2,H_Aux3  From CC_Cuenta_H1, CC_Cuentas_D1" & _
'''                    " WHERE   CC_Cuenta_H1.Par_I = CC_Cuentas_D1.Par_I AND CC_Cuenta_H1.Par_F = CC_Cuentas_D1.Par_F AND CC_Cuentas_D1.Inst='PAG' and CC_Cuenta_H1.Inst='PAG' and " & _
'''                    " CC_Cuentas_D1.O_C=CC_Cuenta_H1.O_C and CC_Cuenta_H1.O_C=3 and  " & _
'''                    " cc_Cuenta_H1.Par_I<='" & recSetAuxcomp!par_codigo & "' and  cc_Cuenta_H1.Par_F>='" & recSetAuxcomp!par_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        End Select
'''        If recSetPartida.RecordCount > 0 Then
'''                                          Sw_Fuente = True
'''                                          Else
'''                                          Sw_Fuente = False
'''        End If
'''
'''
'''    Case Else
'''        Sw_Fuente = False
'''        MsgBox "No esta asociado a ninguna fuente ... partida no relacionada ", vbExclamation + vbDefaultButton1, "Contabilización"
'''        MsgBox recSetAuxcomp!codigo_pago
'''        MsgBox recSetAuxcomp!org_codigo
'''
'''    End Select
'''
'''    If Sw_Fuente Then
'''    'Asignacion a variables
'''        v_Cuenta = recSetPartida!cuenta
'''        v_SubCta1 = recSetPartida!subcta1
'''        v_SubCta2 = recSetPartida!subcta2
'''        v_NombreCta = recSetPartida!NombreCta
'''        v_H_Cuenta = recSetPartida!h_cuenta
'''        v_H_SubCta1 = recSetPartida!h_subcta1
'''        v_H_SubCta2 = recSetPartida!h_subcta2
'''        v_H_NombCta = recSetPartida!H_NombCta
'''
'''        v_Aux1 = recSetPartida!aux1
'''        v_Aux2 = recSetPartida!aux2
'''        v_Aux3 = recSetPartida!aux3
'''
'''        v_H_Aux1 = recSetPartida!h_Aux1
'''        v_H_Aux2 = recSetPartida!h_Aux2
'''        v_H_Aux3 = recSetPartida!h_Aux3
'''
'''        If recSetPartida.State = 1 Then recSetPartida.Close
'''
''''************Abrimos un record set para adicionar datos*********************'
'''        Set recSetAuxActualizar = New ADODB.Recordset
'''        If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''        recSetAuxActualizar.Open " select * from CO_Comprobante_M  where Cod_Trans='" & P_codigo_pago & "' and Org_Codigo='" & P_org_codigo & "' " & _
'''        " and Ges_Gestion='" & P_ges_gestion & "' and Tipo_comp='PAC' and Cod_Trans_Detalle='" & P_codigo_pago_detalle & "'", db, adOpenDynamic, adLockOptimistic, adCmdText
'''        If Not recSetAuxActualizar.BOF Then recSetAuxActualizar.MoveFirst
'''        If (recSetAuxActualizar.BOF) And (recSetAuxActualizar.EOF) Then
''''************* GENERA EL CODIGO DE COMPROBANTE**********
'''            Set recSetGenera = New ADODB.Recordset
'''            recSetGenera.CursorLocation = adUseClient
'''            If recSetGenera.State = 1 Then recSetGenera.Close
'''            recSetGenera.Open "select * from fc_Correl  where tipo_tramite='cmbte'", db, adOpenDynamic, adLockOptimistic, adCmdText
'''            If recSetGenera.RecordCount > 0 Then
'''                Cont_Comp = Val(recSetGenera!numero_correlativo)
'''                Cont_Comp = Cont_Comp + 1
'''                recSetGenera!numero_correlativo = Trim(Str(Cont_Comp))
'''                recSetGenera.Update
'''            End If
'''            If recSetGenera.State = 1 Then recSetGenera.Close
''''************TERMINA GENERACION DE COMPROBANTE********
'''' Datos Para co_Comprobante
'''
'''            recSetAuxActualizar.AddNew
'''            recSetAuxActualizar!Cod_Comp = Val(Cont_Comp)
'''            recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'''            recSetAuxActualizar!Cod_Trans_Detalle = IIf(IsNull(recSetAuxcomp!codigo_pago_detalle), "1", recSetAuxcomp!codigo_pago_detalle)
'''            recSetAuxActualizar!org_codigo = Trim(recSetAuxcomp!org_codigo)
'''            recSetAuxActualizar!Codigo_beneficiario = Trim(recSetAuxcomp!Codigo_beneficiario)
'''            recSetAuxActualizar!Ges_gestion = Trim(recSetAuxcomp!Ges_gestion)
'''            recSetAuxActualizar!num_respaldo = Trim(recSetAuxcomp!codigo_orden)
'''            recSetAuxActualizar!codigo_documento = Trim(recSetAuxcomp!codigo_documento)
'''            recSetAuxActualizar!fecha_A = CDate(recSetAuxcomp!fecha_pago)
'''            recSetAuxActualizar!glosa = Trim(recSetAuxcomp!justificacion)
'''            recSetAuxActualizar!tipo_comp = "PAC"
'''            recSetAuxActualizar!Status = "S"
'''            recSetAuxActualizar!usr_usuario = GlUsuario
'''            recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'''            recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
'''            'recSetAuxActualizar!codigo_solicitud = IIf(IsNull(recSetAuxcomp!codigo_solicitud), "-", recSetAuxcomp!codigo_solicitud)
'''
'''            recSetAuxActualizar.Update
'''            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''
'''' Datos Para co_Diario
'''            Set recSetAuxActualizar1 = New ADODB.Recordset
'''            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'''            recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
'''                recSetAuxActualizar1.AddNew
'''                recSetAuxActualizar1!tipo_comp = "PAC"
'''                recSetAuxActualizar1!d_cuenta = v_Cuenta
'''                recSetAuxActualizar1!D_Nombre = v_NombreCta
'''                recSetAuxActualizar1!d_subcta1 = v_SubCta1
'''                recSetAuxActualizar1!d_subcta2 = v_SubCta2
'''                recSetAuxActualizar1!d_Aux1 = v_Aux1
'''                recSetAuxActualizar1!d_Aux2 = v_Aux2
'''                recSetAuxActualizar1!d_Aux3 = v_Aux3
''''************* CONTABILIZA AUXILIAARES DEBITO
'''                Select Case v_Aux1
'''                Case "01"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    If recsetAdicion.RecordCount <> 0 Then
'''                    'recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'''                    Else
'''                    recSetAuxActualizar1!d_cta_larga = recSetAuxcomp!Codigo_beneficiario
'''                    End If
'''                    'recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'''
'''                Case "02"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'''                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'''                Case Else
'''                End Select
'''''****************** finaliza sesion de auxiliares
'''                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
'''                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
'''                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
''''************* CONTABILIZA AUXILIAARES CREDITO
'''
'''                Select Case v_H_Aux1
'''                Case "01"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'''                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'''
'''                Case "02"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
''''recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'''                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'''
'''                Case Else
'''                End Select
'''''****************** finaliza sesion de auxiliares
'''
'''                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
'''                recSetAuxActualizar1!H_Nombre = v_H_NombCta
'''                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
'''                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
'''                recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_bolivianos
'''                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'''                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'''
'''                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_bolivianos
'''                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'''                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'''                recSetAuxActualizar1!Cod_Comp = Cont_Comp
'''                recSetAuxActualizar1!usr_usuario = GlUsuario
'''                recSetAuxActualizar1!fecha_registro = Format(Date, "dd/mm/yyyy")
'''                recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
'''                recSetAuxActualizar1.Update
'''            End If
'''        Else
'''        MsgBox "Ya fue contabilizado anteriormente...  ", vbInformation + vbOKOnly, "Contabilización...  "
'''
'''
'''' buscar el que ya existe y reemplazar los datos
'''
'''            If (Not recSetAuxActualizar.BOF) Then recSetAuxActualizar.MoveFirst
''''            recSetAuxActualizar!Cod_Comp = Cont_Comp
'''            Cont_Comp = recSetAuxActualizar!Cod_Comp
'''            recSetAuxActualizar!Cod_trans = recSetAuxcomp!codigo_pago
'''            recSetAuxActualizar!Cod_Trans_Detalle = recSetAuxcomp!codigo_pago_detalle
'''            recSetAuxActualizar!org_codigo = recSetAuxcomp!org_codigo
'''            recSetAuxActualizar!Codigo_beneficiario = recSetAuxcomp!Codigo_beneficiario
'''            recSetAuxActualizar!Ges_gestion = recSetAuxcomp!Ges_gestion
'''            recSetAuxActualizar!num_respaldo = recSetAuxcomp!codigo_orden
'''            recSetAuxActualizar!codigo_documento = recSetAuxcomp!codigo_documento
'''            recSetAuxActualizar!fecha_A = CDate(recSetAuxcomp!fecha_pago)
'''            '''''g--
'''            recSetAuxActualizar!glosa = recSetAuxcomp!justificacion
''''            recSetAuxActualizar!Tipo_Comp = "PAC"
'''            recSetAuxActualizar!Status = "S"
'''            recSetAuxActualizar!usr_usuario = GlUsuario
'''            recSetAuxActualizar!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
'''            recSetAuxActualizar!hora_registro = Format(Time, "hh:mm:ss")
'''            recSetAuxActualizar.Update
'''            If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''            ' Datos Para co_Diario
'''            Set recSetAuxActualizar1 = New ADODB.Recordset
'''            If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'''            recSetAuxActualizar1.Open " select * from CO_Diario where  cod_Comp = " & Cont_Comp & " ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''            If (recSetAuxActualizar1.BOF) And (recSetAuxActualizar1.EOF) Then
'''                recSetAuxActualizar1.AddNew
'''                recSetAuxActualizar1!tipo_comp = "PAC"
'''                recSetAuxActualizar1!Cod_Comp = Cont_Comp
'''            Else
'''                If (Not recSetAuxActualizar1.BOF) Then recSetAuxActualizar1.MoveFirst
'''            End If
'''                recSetAuxActualizar1!d_cuenta = v_Cuenta
'''                recSetAuxActualizar1!D_Nombre = v_NombreCta
'''                recSetAuxActualizar1!d_subcta1 = v_SubCta1
'''                recSetAuxActualizar1!d_subcta2 = v_SubCta2
'''                recSetAuxActualizar1!d_Aux1 = v_Aux1
'''                recSetAuxActualizar1!d_Aux2 = v_Aux2
'''                recSetAuxActualizar1!d_Aux3 = v_Aux3
''''************* CONTABILIZA AUXILIAARES DEBITO
'''                Select Case v_Aux1
'''                Case "01"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!Codigo_beneficiario
'''                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!denominacion_beneficiario
'''
'''                Case "02"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!d_cta_larga = recsetAdicion!cta_codigo
'''                    recSetAuxActualizar1!d_des_Larga = recsetAdicion!cta_descripcion_larga
'''                Case Else
'''                End Select
'''''****************** finaliza sesion de auxiliares
'''                recSetAuxActualizar1!h_Aux1 = v_H_Aux1
'''                recSetAuxActualizar1!h_Aux2 = v_H_Aux2
'''                recSetAuxActualizar1!h_Aux3 = v_H_Aux3
''''************* CONTABILIZA AUXILIAARES CREDITO
'''
'''                Select Case v_H_Aux1
'''                Case "01"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''                    recsetAdicion.Open " select * from fc_beneficiario where codigo_Beneficiario='" & recSetAuxcomp!Codigo_beneficiario & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!Codigo_beneficiario
'''                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!denominacion_beneficiario
'''
'''                Case "02"
'''                    Set recsetAdicion = New ADODB.Recordset
'''                    If recsetAdicion.State = 1 Then recsetAdicion.Close
'''
'''                    recsetAdicion.Open " select * from fc_cuenta_Bancaria where cta_Codigo='" & recSetAuxcomp!cta_codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
''''recsetAdicion.Open " select * from fc_cuenta_Bancaria where codigo_Cuenta='" & recSetAuxcomp!cta_Codigo & "' ", db, adOpenDynamic, adLockOptimistic, adCmdText
'''                    recSetAuxActualizar1!h_cta_larga = recsetAdicion!cta_codigo
'''                    recSetAuxActualizar1!h_des_Larga = recsetAdicion!cta_descripcion_larga
'''
'''                Case Else
'''                End Select
'''''****************** finaliza sesion de auxiliares
'''
'''                recSetAuxActualizar1!h_cuenta = v_H_Cuenta
'''                recSetAuxActualizar1!H_Nombre = v_H_NombCta
'''                recSetAuxActualizar1!h_subcta1 = v_H_SubCta1
'''                recSetAuxActualizar1!h_subcta2 = v_H_SubCta2
'''                recSetAuxActualizar1!d_montobs = recSetAuxcomp!monto_bolivianos
'''                recSetAuxActualizar1!d_montoDl = recSetAuxcomp!monto_Dolares
'''                recSetAuxActualizar1!d_Cambio = recSetAuxcomp!tipo_cambio
'''                recSetAuxActualizar1!h_montoBs = recSetAuxcomp!monto_bolivianos
'''                recSetAuxActualizar1!h_montoDl = recSetAuxcomp!monto_Dolares
'''                recSetAuxActualizar1!h_Cambio = recSetAuxcomp!tipo_cambio
'''                recSetAuxActualizar1!usr_usuario = GlUsuario
'''                recSetAuxActualizar1!fecha_registro = Format(Date, "dd/mm/yyyy")
'''                recSetAuxActualizar1!hora_registro = Format(Time, "hh:mm:ss")
'''                recSetAuxActualizar1.Update
'''        End If
'''    Else
'''           MsgBox "No esta asociado a ninguna fuente ...  ", vbInformation + vbOKOnly, "Contabilizacion"
'''    End If
'''    recSetAuxcomp.MoveNext
'''MsgBox "Contabilizacion exitosa...... ", vbInformation + vbOKOnly, "Contabilizacion"
'''Wend
'''db.CommitTrans
'''
'''
'''    Set recSetAuxcomp = New ADODB.Recordset
'''    recSetAuxcomp.CursorLocation = adUseClient
'''    If recSetAuxcomp.State = 1 Then recSetAuxcomp.Close
'''
'''    Set recSetAuxActualizar = New ADODB.Recordset
'''    If recSetAuxActualizar.State = 1 Then recSetAuxActualizar.Close
'''
'''    Set recSetAuxActualizar1 = New ADODB.Recordset
'''    If recSetAuxActualizar1.State = 1 Then recSetAuxActualizar1.Close
'''
'''    Set recSetPartida = New ADODB.Recordset
'''    recSetPartida.CursorLocation = adUseClient
'''    If recSetPartida.State = 1 Then recSetPartida.Close
'''
'''
'''Exit Sub
'''errorPag:
'''db.RollbackTrans
'''MsgBox "No se contabilizó ... ", vbCritical + vbDefaultButton1, "Contabilización"
'''
'''End Sub

Private Sub CmdGrabar_Click()
    Dim v_codigo_Pago As Integer
    Dim V_CTACOD, v_convenio As String
    Dim v_Codigo_Pago_dedtalle As String
    Dim v_Org As String
    Dim v_Gestion As String
    Dim P_fecha_pago As String
    Dim P_Glosa As String
    Dim P_D_SubCta2 As String
    Dim P_H_SubCta2 As String
    Dim P_MontoBs As Double
    Dim P_MontoDl As Double
    Dim P_Cambio As Double
    Dim SaldoBancario_Real As Double
    Dim SaldoBancario_dol As Double
    '****************************************************************************************************************************************************************************
    '*****************   O B S E R V A C I O N E S
    '*****************   Se grabará en fecha_pago cuando se realiza el pago
    '*****************   Se grabará en fecha_impresion_cheque cuando se realiza el pago
    'Determinar que el monto deba ser menor o igual al recuperado
    
    If Val(TxtMonto.Text) > Val(TxtMB.Text) Then
        MsgBox "El monto debe ser menor o igual al recuperado de la base", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
'Determina si se puede o no cancelar
    Set rsCta = New ADODB.Recordset
    rsCta.Open "select * from fc_cuenta_bancaria where cta_codigo='" & DtCCuentaOrigen.Text & "' ", db, adOpenKeyset, adLockOptimistic
    'JQA 24/11/01
    If rsCta.RecordCount > 0 Then
        '''''SaldoBancario_Real = rsCta("Cta_saldo_inicial") - rsCta("Cta_Acumulado") + rsCta("Cta_Saldo_Debe") + rsCta("Cta_Pco_Debe") - rsCta("Cta_Pco_Haber") + rsCta("Cta_Ingresos") + rsCta("Cta_Acum_dev") + rsCta("Cta_Acum_anl")
        SaldoBancario_Real = rsCta(IIf(IsNull("Cta_saldo_actual"), 0, "Cta_saldo_actual"))
        SaldoBancario_dol = rsCta(IIf(IsNull("saldo_actual"), 0, "saldo_actual"))
    Else
        MsgBox "No existe cuenta bancaria", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    If (rsCta("cta_patron") = "F") And (SaldoBancario_Real - Val(TxtMonto.Text) < 0) Then
        MsgBox "No existe saldo en Bs. para realizar el pago", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    If (rsCta("cta_patron") = "E") And (SaldoBancario_dol - (Val(TxtMonto.Text) / AdoPagoDetalle.Recordset("tipo_cambio")) < 0) Then
        MsgBox "No existe saldo en $US para realizar el pago", vbInformation + vbCritical, "Validación de datos"
        Exit Sub
    End If
    'JQA 24/11/01
    Marca = AdoPagoDetalle.Recordset.AbsolutePosition
    AdoPagoDetalle.Recordset("ges_gestion") = AdoPago.Recordset("ges_gestion")
    AdoPagoDetalle.Recordset("org_codigo") = AdoPago.Recordset("org_codigo")
    AdoPagoDetalle.Recordset("codigo_pago") = AdoPago.Recordset("codigo_pago")
    
    If OptDeposito.Value = True Then AdoPagoDetalle.Recordset("cheque_o_trf") = "D"
    If OptChequeOrigen.Value = True Then AdoPagoDetalle.Recordset("cheque_o_trf") = "C"
    If OptTransferenciaOrigen.Value = True Then AdoPagoDetalle.Recordset("cheque_o_trf") = "T"
    AdoPagoDetalle.Recordset("Observacion") = TxtObs
    If ChkHonorarios.Value = 1 Then AdoPagoDetalle.Recordset("honorarios") = "H"
    If ChkHonorarios.Value = 0 Then AdoPagoDetalle.Recordset("honorarios") = "S"
    If ChkNombreBeneficiario.Value = 1 Then
       AdoPagoDetalle.Recordset("beneficiario_destino") = TxtBeneDest.Text
     Else
       AdoPagoDetalle.Recordset("beneficiario_destino") = " "
     End If
    
    If OptTransferenciaOrigen.Value = False And OptChequeOrigen.Value = False And OptDeposito.Value = False Then
       MsgBox "Click en la opción de cheque o transferencia", vbCritical + vbInformation, "Validación de datos"
       Exit Sub
    End If
    AdoPagoDetalle.Recordset("numero_cheque_trf") = TxtNoTransaccion.Text
    AdoPagoDetalle.Recordset("estado_aprobacion") = "N"
    AdoPagoDetalle.Recordset("numero_cheque_trf_destino") = TxtNoTransac.Text
    If DtCCuentaOrigen.Text <> "" Then
      AdoPagoDetalle.Recordset("cta_codigo") = DtCCuentaOrigen.Text
    Else
      MsgBox "Introducir Cuenta Origen", vbCritical + vbInformation, "Validación de datos"
      Exit Sub
    End If
    
    If TxtCuentaDestino.Text <> "" Then
      AdoPagoDetalle.Recordset("cta_codigo_destino") = TxtCuentaDestino.Text
    Else
      AdoPagoDetalle.Recordset("cta_codigo_destino") = ""
    End If
    
    If TxtBancoDestino.Text <> "" Then
      AdoPagoDetalle.Recordset("banco_destino") = TxtBancoDestino
    Else
      AdoPagoDetalle.Recordset("banco_destino") = ""
    End If
    
    If OptTransferenciaOrigen.Value = True Then
        If CmbNomDep.Text <> "" Then
            AdoPagoDetalle.Recordset("departamento") = CmbNomDep.Text
        Else
            MsgBox "Introducir nombre de departamento ", vbCritical + vbInformation, "Validación de datos"
            Exit Sub
        End If
    End If

    'If TxtFechaPago.Text = "" Then
        AdoPagoDetalle.Recordset("fecha_pago") = Date 'TxtFechaPago.Text
        AdoPagoDetalle.Recordset("Fecha_Aprobacion_tesoreria") = Date
    'Else
    '  MsgBox "Introducir fecha de pago", vbCritical + vbInformation, "Validación de datos"
    '  Exit Sub
    'End If
    
    If Val(TxtMonto.Text) <> 0 Then     'Or Val(txtTipoCambio.Text) <> 0
      AdoPagoDetalle.Recordset("monto_bolivianos") = CCur(CDbl(TxtMonto.Text))
      AdoPagoDetalle.Recordset("tipo_cambio") = GlTipoCambioMercado
      AdoPagoDetalle.Recordset("monto_dolares") = Val(TxtMonto.Text) / GlTipoCambioMercado  'Val(txtTipoCambio.Text)
    Else
      MsgBox "Introducir Monto total o tipo de cambio", vbCritical + vbInformation, "Validación de datos"
      Exit Sub
    End If
    AdoPagoDetalle.Recordset("saldo_bolivianos") = 0
    If CStr(AdoPagoDetalle.Recordset("monto_bolivianos")) <> "" Then
        AdoPagoDetalle.Recordset("literal") = Literal(CStr(AdoPagoDetalle.Recordset("monto_bolivianos"))) + " BOLIVIANOS"
    End If
    
    'Datos de seguimiento
    AdoPagoDetalle.Recordset("usr_usuario") = LblUsuario.Caption
    AdoPagoDetalle.Recordset("fecha_registro") = Date
    AdoPagoDetalle.Recordset("hora_registro") = Format(Time, "hh:mm:ss")
    
    AdoPagoDetalle.Recordset.Update
    
    FraPagoDetalle.Enabled = False
'    FraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
        
    Set rsControlDet = New ADODB.Recordset
    rsControlDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "' and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "' order by codigo_pago_detalle ", db, adOpenKeyset, adLockOptimistic
    If rsControlDet.RecordCount > 0 Then
      Set AdoPagoDetalle.Recordset = rsControlDet
      Set DtGPagosParciales.DataSource = AdoPagoDetalle
      Set DtGPP.DataSource = AdoPagoDetalle
      AdoPagoDetalle.Recordset.AbsolutePosition = Marca
        AdoPagoDetalle.Recordset.MoveFirst
        DtGPagosParciales.Refresh
        DtGPP.Refresh
    Else
      Set DtGPagosParciales.DataSource = rsNada
      DtGPagosParciales.ReBind
    End If
    
    AdoPago.Enabled = True
    DtgPago.Enabled = True
    
    swModifica = 0
    swPagoParcial = 0
    TxtMonto.Enabled = True
    CtaAnterior = ""
    
    v_codigo_Pago = AdoPagoDetalle.Recordset("codigo_pago")
    v_Codigo_Pago_dedtalle = AdoPagoDetalle.Recordset("codigo_pago_detalle")
    v_Org = AdoPagoDetalle.Recordset("org_codigo")
    v_Gestion = AdoPagoDetalle.Recordset("ges_gestion")
    V_CTACOD = AdoPagoDetalle.Recordset("cta_codigo")

    'Call ESTADO_APROBADO(AdoPagoDetalle.Recordset("Codigo_Pago"), AdoPagoDetalle.Recordset("Org_Codigo"), AdoPagoDetalle.Recordset("Ges_Gestion"))
       'estado pagado=S
        Set rspago = New ADODB.Recordset
        If rspago.State = 1 Then rspago.Close
        rspago.Open "SELECT * from pagos where codigo_pago= '" & AdoPagoDetalle.Recordset("Codigo_Pago") & "' and ges_gestion= '" & AdoPagoDetalle.Recordset("Ges_Gestion") & "' and org_codigo='" & AdoPagoDetalle.Recordset("Org_Codigo") & "'", db, adOpenKeyset, adLockOptimistic
        If rspago.RecordCount > 0 Then
          AdoPago.Recordset("estado_pagado") = "S"
          v_convenio = AdoPago.Recordset("estado_pagado")
          v_UNIDAD = IIf(IsNull(AdoPago.Recordset("codigo_unidad")), "ALFAPARF", AdoPago.Recordset("codigo_unidad"))
          V_SOLICITUD = AdoPago.Recordset("codigo_solicitud")
          AdoPago.Recordset.Update
        End If
        
        
    'Procedimiento que da por aprobado el pago
    'Call ESTADO_APROBADO(CodigoPago, CodigoOrg, GesGestion)

'MsgBox AdoPagoDetalle.Recordset("estado_aprobacion")
''GROVER.........
'''''If swModifica <> 1 Then
'''''    If AdoPago.Recordset("tipo_comp") = "PCE" And AdoPago.Recordset("Org_Codigo") = "999" Then
'''''        Cmd_Contabiliza v_codigo_Pago
'''''    End If
'''''    If AdoPago.Recordset("tipo_comp") = "DAC" Then
'''''        Cmd_Pagado v_codigo_Pago, v_Codigo_Pago_dedtalle, v_Org, v_Gestion
'''''    End If
'''''End If

'Esto es lo último que copié
    'If swModifica <> 1 Then
           If AdoPago.Recordset("tipo_comp") = "PCE" And AdoPago.Recordset("Org_Codigo") = "999" Then
               Cmd_contabiliza v_codigo_Pago, v_Org, v_Gestion
            End If
            
         '   If AdoPago.Recordset("tipo_comp") = "PCE" And AdoPago.Recordset("Org_Codigo") <> "999" Then
        '        Cmd_Contabiliza v_codigo_Pago
        '    End If
            
            If AdoPago.Recordset("tipo_comp") = "DAC" Then
                Cmd_Pagado v_codigo_Pago, v_Codigo_Pago_dedtalle, v_Org, v_Gestion
            End If
    'End If

'ini - Actualiza saldos por 07-12-00

  Dim cmdts_mf_ActualizaCtaBancaria As ADODB.Command
  Set cmdts_mf_ActualizaCtaBancaria = New ADODB.Command
  With cmdts_mf_ActualizaCtaBancaria
      .ActiveConnection = db
      .CommandType = adCmdStoredProc
      .CommandText = "ts_mf_ActualizaCtaBancariaZ"
      .Parameters("@P_CTA") = Trim(V_CTACOD)
      .Parameters("@P_CONVENIO") = Trim(v_convenio)
      .Execute
  End With


'  Execute ts_mf_ActualizaCtaBancaria
'fin - Actualiza saldos por 07-12-00
FraDatosCarta.Enabled = False
FraP.Visible = True
Exit Sub
error_GRABAR:
    MsgBox Err.Number & " " & Err.Description
End Sub


Private Sub Cmdimprimir_Click()
    ' FrmImprimirComprobantenuevo.s
    ' FraImprimeCmpte.Visible = True
    FrmImprimeComprobanteNuevo.Show 'vbModal
    ' FraImprimeCmpte.Visible = False
End Sub

Private Sub CmdImprimirPagos_Click()
    FrmComprobante.Show
End Sub

Private Sub CmdImprimirTransfer_Click()
    FrmTransferenciasNuevo.Show vbModal
End Sub

Private Sub CmdModificar_Click()

Dim x As Variant
    FraPagoDetalle.Enabled = True
    FraGrabarCancelar.Visible = True
'    FraOpciones.Visible = False
    FraPagoDetalle.Enabled = True
    FraPagosParciales.Enabled = True
    TxtMonto.Enabled = True
    swModifica = 1
    If Not IsNull(AdoPagoDetalle.Recordset("monto_bolivianos")) Then MontoAnterior = AdoPagoDetalle.Recordset("monto_bolivianos")
    If Not IsNull(AdoPagoDetalle.Recordset("cta_codigo")) Then CtaAnterior = AdoPagoDetalle.Recordset("cta_codigo")
         x = AdoPagoDetalle.Recordset("codigo_pago_detalle")
         Set rsPagoDet = New ADODB.Recordset
         rsPagoDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'And codigo_pago_detalle = '" & AdoPagoDetalle.Recordset("codigo_pago_detalle") & "'", db, adOpenKeyset, adLockOptimistic
         'rsPagoDet.Open "select * from pago_detalle where ges_gestion='" & rsPago("ges_gestion") & "' and org_codigo='" & rsPago("org_codigo") & "'and codigo_pago='" & rsPago("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
         If rsPagoDet.RecordCount > 0 Then
            If Not IsNull(rsPagoDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsPagoDet("codigo_beneficiario")
            If Not IsNull(rsPagoDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsPagoDet("cta_codigo")
         End If
    FraDatosCarta.Enabled = True
    OptObs2.Value = False
    OptObs1.Value = False
Exit Sub
End Sub

Private Sub CmdNuevoBeneficiario_Click()
    FraBeneficiario.Visible = True
End Sub

Private Sub CmdPagoGrupal_Click()
    FraPagoGrupal.Visible = True
    DtgPago.Enabled = False
End Sub

Private Sub CmdPagoIndividual_Click()
    'Caso opción cheque
    TxtCuentaDestino.Visible = False
    LblCtaDestino.Visible = False
    TxtBancoDestino.Visible = False
    LblBancoDestino.Visible = False
    CmbNomDep.Visible = False
    LblDeducciones.Visible = False
    
    FraPagosParciales.Visible = True
    
    TxtNC.Text = TxtCodigoOrden.Text
    If IsNull(TxtMontoBolivianos.Text) Then
        TxtMB.Text = 0
    Else
        TxtMB.Text = Round(CDbl(TxtMontoBolivianos.Text), 2)
    End If
    TxtBeneDest.Text = TxtNomBen.Text
    'Datos del Control de Datos
         Set rsPagoDet = New ADODB.Recordset
         rsPagoDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "' order by codigo_pago_detalle", db, adOpenKeyset, adLockOptimistic
         If rsPagoDet.RecordCount > 0 Then
            If Not IsNull(rsPagoDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsPagoDet("codigo_beneficiario")
            If Not IsNull(rsPagoDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsPagoDet("cta_codigo")
                Set AdoPagoDetalle.Recordset = rsPagoDet
                Set DtGPagosParciales.DataSource = AdoPagoDetalle
                Set DtGPP.DataSource = AdoPagoDetalle
                DtGPagosParciales.Refresh
                rsPagoDet.MoveLast
                   
         Else
            Set DtGPagosParciales.DataSource = rsNada
            Set DtGPP.DataSource = rsNada
         End If
'         FraOpciones.Visible = True
         FraP.Visible = True
     
     'Actualizando Cuenta Bancaria
        If AdoPago.Recordset("tipo_comp") = "TRP" Then
                DtCCuentaOrigen.Text = AdoPagoDetalle.Recordset("CTA_CODIGO")
                Set rscuenta = New ADODB.Recordset
                If rscuenta.State = 1 Then rscuenta.Close
                rscuenta.Open "select * from fc_cuenta_bancaria where cta_codigo='" & AdoPagoDetalle.Recordset("CTA_CODIGO") & "'", db, adOpenKeyset, adLockOptimistic
                DtCCuentaOrigen.Text = rscuenta("cta_codigo")
                DtcCtaTGN.Text = rscuenta("cta_codigo_tgn")
                DtCCuentaOrigenDes.Text = rscuenta("cta_descripcion_larga")
                Set AdoCuenta.Recordset = rscuenta
                Exit Sub
         End If
         If AdoPago.Recordset("tipo_comp") = "PCE" Then
            Set rscuenta = New ADODB.Recordset
            If rscuenta.State = 1 Then rscuenta.Close
            rscuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
            Set AdoCuenta.Recordset = rscuenta
         Else
            Set rscuenta = New ADODB.Recordset
            If rscuenta.State = 1 Then rscuenta.Close
            'MODIFICACION POR NUEVA MODALIDAD DE PAGOS DE LA KFW
'            If AdoPago.Recordset("org_codigo") = "516" Then
'              rscuenta.Open "select * from fc_cuenta_bancaria where org_codigo='111'", db, adOpenKeyset, adLockOptimistic
'            Else
              rscuenta.Open "select * from fc_cuenta_bancaria where org_codigo='" & AdoPago.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'            End If
            Set AdoCuenta.Recordset = rscuenta
         End If
    Call CmdPagoTotal_Click
End Sub

Private Sub CmdPagoParcial_Click()
    FraTotalParcial.Visible = False
    lblmonto.Caption = "Monto Parcial"
'    LblDeducciones.Visible = True
'    TxtDeducciones.Visible = True
    OK = ControlSuma
    If OK = 1 Then
        AdoPago.Enabled = False
        DtgPago.Enabled = False
        rsControlPago.AddNew
        OK = 0
    End If
End Sub
Private Sub CmdPagos_Click()
    FraPagosParciales.Visible = True
    TxtNC.Text = TxtCodigoOrden.Text
    
    TxtMB.Text = TxtMontoBolivianos.Text
    'Datos del Control de Datos
         Set rsPagoDet = New ADODB.Recordset
         'rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsPago("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
         rsPagoDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "' order by codigo_pago_detalle", db, adOpenKeyset, adLockOptimistic
         If rsPagoDet.RecordCount > 0 Then
            If Not IsNull(rsPagoDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsPagoDet("codigo_beneficiario")
            'If Not IsNull(rsPagoDet("deducciones")) Then TxtDeducciones.Text = rsPagoDet("deducciones")
            If Not IsNull(rsPagoDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsPagoDet("cta_codigo")
            Set AdoPagoDetalle.Recordset = rsPagoDet
            Set DtGPagosParciales.DataSource = AdoPagoDetalle
            Set DtGPP.DataSource = AdoPagoDetalle
            DtGPagosParciales.Refresh
            rsPagoDet.MoveLast
         Else
            Set DtGPagosParciales.DataSource = rsNada
            Set DtGPP.DataSource = rsNada
         End If
'         FraOpciones.Visible = True
         FraP.Visible = True
         
End Sub

Private Sub CmdPagosParciales_Click()
    If TxtMB.Text = "0" Then
        MsgBox "No existe monto para asignar", vbInformation, "Validación de datos"
        Exit Sub
    End If
    Total
       'If NumReg <> 1 Then
          FraTotalParcial.Visible = False
          lblmonto.Caption = "Monto Parcial"
          
          If (SumaTotal = Val(TxtMontoBolivianos.Text)) Then
              MsgBox "Se canceló todo el monto o no se puede cancelar monto total", vbCritical + vbInformation, "Validación de datos"
              FraGrabarCancelar.Visible = False
'              FraOpciones.Visible = True
              FraPagoDetalle.Enabled = False
              ControlSuma = 0
          
              Set rsCtrlPago = New ADODB.Recordset
              rsCtrlPago.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "' and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
              Set AdoPagoDetalle.Recordset = rsCtrlPago
              Set DtGPagosParciales.DataSource = AdoPagoDetalle
              Set DtGPP.DataSource = AdoPagoDetalle
              swPagoTotal = 0
          End If
          If SumaTotal < Val(TxtMontoBolivianos.Text) Then
              ControlSuma = 1
              FraGrabarCancelar.Visible = True
'              FraOpciones.Visible = False
              FraPagoDetalle.Enabled = True
              FraPagosParciales.Enabled = True
              swPagoParcial = 0
          
          End If
        
          OK = ControlSuma
          swPagoParcial = 1
          If OK = 1 Then
              AdoPago.Enabled = False
              DtgPago.Enabled = False
              AdoPagoDetalle.Recordset.AddNew
              'rsPagoDetalle.AddNew
              OK = 0
          End If
     'Else
     '     FraOpciones.Visible = False
     '     FraGrabarCancelar.Visible = True
     '     FraPagoDetalle.Enabled = True
     '     FraPagosParciales.Enabled = True
     'End If
End Sub

Private Sub CmdPagoTotal_Click()
On Error GoTo error_Pagar:
    If TxtMB.Text = "0" Then
        MsgBox "No existe monto para asignar", vbInformation, "Validación de datos"
        Exit Sub
    End If

    Total
    TxtMonto.Text = TxtMB.Text
    TxtNoTransaccion = ""
    TxtMonto.Enabled = False
'    FraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraPagoDetalle.Enabled = True
    FraP.Visible = False
    LblNumeroOrigen.Caption = "Nro.Cheque/Transac."
    TxtObs.Text = ""
    ChkHonorarios.Value = 0
    ChkNombreBeneficiario.Value = 0
    TxtBancoDestino.Text = ""
    TxtCuentaDestino.Text = ""
    TxtNoTransaccion.Text = ""
    CmbNomDep = "CHUQUISACA"
    OptObs2.Value = True
    OptObs1.Value = True
    SSTTransferencia.TabVisible(0) = True
    SSTTransferencia.TabVisible(1) = False
    'ini - Actualiza saldos por 07-12-00

    Dim cmdts_mf_ActualizaCtaBancaria As ADODB.Command
    Set cmdts_mf_ActualizaCtaBancaria = New ADODB.Command
    With cmdts_mf_ActualizaCtaBancaria
      .ActiveConnection = db
      .CommandType = adCmdStoredProc
      .CommandText = "ts_mf_ActualizaCtaBancariaZ"
      .Parameters("@P_CTA") = "NO_HAY"
      .Parameters("@P_CONVENIO") = AdoPago.Recordset("codigo_convenio")
      .Execute
    End With
    
    '  Execute ts_mf_ActualizaCtaBancaria
    'fin - Actualiza saldos por 07-12-00

Exit Sub
error_Pagar:
    MsgBox Err.Number & " " & Err.Description

End Sub

Private Sub CmdRestaurarPagos_Click()
    Set rspago = New ADODB.Recordset
    rspago.Open "select * from pagos where (estado_contabilidad='P' or estado_devengado='S' ) and  (estado_aprobacion <>'A' or estado_aprobacion IS NULL) and (estado_pagado<>'S' or estado_pagado IS NULL) order by codigo_pago", db, adOpenKeyset, adLockOptimistic
    If rspago.RecordCount > 0 Then
      Set AdoPago.Recordset = rspago
      Set DtgPago.DataSource = AdoPago
      DtgPago.ReBind
    Else
      MsgBox "No existen registros !", vbInformation, "Validación de datos"
      DtgPago.Enabled = False
      Set DtGPagosParciales.DataSource = rsNada
      Exit Sub
    End If
    
End Sub

Private Sub CmdSale_Click()
'For I = 1 To 100
    Me.MousePointer = vbDefault
    
    'PrBPagosTotales.Value = I
    'I = I + 1
'Next I
FraPagoGrupal.Visible = False
DtgPago.Enabled = True
End Sub

Private Sub CmdSalir_Click()
Dim sino As Variant
    sino = MsgBox("Está seguro de salir?", vbYesNo + vbQuestion, "Atenciòn")
    If sino = vbYes Then
        FraPagosParciales.Visible = False
        FraMensajeImportante.Visible = False
    End If
End Sub

Private Sub CmdSalirControl_Click()
    Unload Me
End Sub

Private Sub CmdSalirPagos_Click()
    Unload Me
End Sub

Private Sub CmdTotal_Click()
'On Error GoTo error_Pagar:
    TxtFechaPago.Text = ""
    DtCCuentaOrigen.Text = ""
    DtcCtaTGN.Text = ""
    DtCCuentaOrigenDes.Text = ""
    TxtCuentaDestino.Text = ""
    TxtMonto.Text = ""
'    FraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraTotalParcial.Visible = True
    LblNumeroOrigen.Caption = "Nro.Cheque/Transac."
    TxtNoTransaccion.Text = ""
    FraPagoDetalle.Enabled = True
  
  
    FraTotalParcial.Visible = False
    lblmonto.Caption = "Monto Total"
    TxtMonto.Text = TxtMontoBolivianos.Text
    OK = ControlSuma
    If OK = 1 Then
        AdoPago.Enabled = False
        DtgPago.Enabled = False
        rsPAgoDetalle.AddNew
        OK = 0
    End If
End Sub

Private Sub Command1_Click()
'   FrmChequesCuenta.Show
 FrmChequesNuevo.Show
End Sub

Private Sub DtCCta_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCCta.BoundText
    DtCTgn.BoundText = DtCCta.BoundText
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub dtcNombreRuc_Click(Area As Integer)
    dtcRuc.BoundText = dtcNombreRuc.BoundText
End Sub

Private Sub dtcRuc_Click(Area As Integer)
   dtcNombreRuc.BoundText = dtcRuc.BoundText
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub



Private Sub DtCDescripcion_Click(Area As Integer)
   DtCTgn.BoundText = DtCDescripcion.BoundText
   DtCCta.BoundText = DtCDescripcion.BoundText
End Sub

Private Sub DtCTgn_Click(Area As Integer)
    DtCDescripcion.BoundText = DtCTgn.BoundText
    DtCCta.BoundText = DtCTgn.BoundText
End Sub

Private Sub DtgPago_Click()
   If Not AdoPago.Recordset.EOF And Not AdoPago.Recordset.BOF Then
         If Not IsNull(AdoPago.Recordset("codigo_pago")) Then TxtCodigoOrden.Text = AdoPago.Recordset("codigo_pago")
         'If Not IsNull(AdoPago.Recordset("monto_Bolivianos")) Then TxtMontoBolivianos.Text = AdoPago.Recordset("monto_Bolivianos")
         If Not IsNull(AdoPago.Recordset("liquido_pagar")) Then TxtMontoBolivianos.Text = Round(CDbl(AdoPago.Recordset("liquido_pagar")), 2) Else TxtMontoBolivianos.Text = ""
         If Not IsNull(AdoPago.Recordset("tipo_comp")) Then TxtTipo.Text = AdoPago.Recordset("tipo_comp")
         
         'Datos del Control de Datos
         Set rsControlDet = New ADODB.Recordset
         rsControlDet.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "'and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
         If rsControlDet.RecordCount > 0 Then
           If Not IsNull(rsControlDet("codigo_beneficiario")) Then TxtCodigoBen.Text = rsControlDet("codigo_beneficiario")
           'If Not IsNull(rsControlDet("deducciones")) Then TxtDeducciones.Text = rsControlDet("deducciones")
            If Not IsNull(rsControlDet("cta_codigo")) Then DtCCuentaOrigen.Text = rsControlDet("cta_codigo")
           If Not IsNull(rsControlDet("fecha_pago")) Then TxtFechaPago.Text = rsControlDet("fecha_pago")
           If Not IsNull(rsControlDet("monto_bolivianos")) Then
              LbLAprobado.Caption = "PAGADO"
           Else
              LbLAprobado.Caption = "POR PAGAR"
           End If
           
           Set AdoPagoDetalle.Recordset = rsControlDet
           Set DtGPagosParciales.DataSource = AdoPagoDetalle
           Set DtGPP.DataSource = AdoPagoDetalle
           DtGPagosParciales.Refresh
           rsControlDet.MoveLast
         Else
           Set DtGPagosParciales.DataSource = rsNada
           Set DtGPP.DataSource = rsNada
           DtGPagosParciales.ReBind
         End If
         
         Set rsbeneficiario = New ADODB.Recordset
         rsbeneficiario.Open "select * from fc_beneficiario where codigo_beneficiario='" & TxtCodigoBen.Text & "'", db, adOpenKeyset, adLockOptimistic
         If rsbeneficiario.RecordCount > 0 Then
         TxtNomBen.Text = rsbeneficiario("denominacion_beneficiario")
         End If
         rsbeneficiario.Close
         
            Set rscuenta = New ADODB.Recordset
            If rscuenta.State = 1 Then rscuenta.Close
            'MODIFICACION POR LA NUEVA MODALIDAD DE PAGO DE LA KFW (PAGO CON CUENTAS DEL TGN)
'            If AdoPago.Recordset!org_codigo = "516" Then
'              rscuenta.Open "select * from fc_cuenta_bancaria where org_codigo='111'", db, adOpenKeyset, adLockOptimistic
'            Else
'              rscuenta.Open "select * from fc_cuenta_bancaria where org_codigo='" & AdoPago.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'            End If
            rscuenta.Open "select * from fc_cuenta_bancaria where org_codigo='" & AdoPago.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
            Set AdoCuenta.Recordset = rscuenta
         
End If
End Sub

Private Sub DtGPP_Click()
If adReason <> 10 Then
  If Not AdoPagoDetalle.Recordset.EOF And Not AdoPagoDetalle.Recordset.BOF Then
    If Not IsNull(AdoPagoDetalle.Recordset("cta_codigo")) Then DtCCuentaOrigen.Text = AdoPagoDetalle.Recordset("cta_codigo")
    If Not IsNull(AdoPagoDetalle.Recordset("cta_codigo_destino")) Then TxtCuentaDestino.Text = AdoPagoDetalle.Recordset("cta_codigo_destino")
    If Not IsNull(AdoPagoDetalle.Recordset("numero_cheque_trf")) Then TxtNoTransaccion.Text = AdoPagoDetalle.Recordset("numero_cheque_trf")
    If Not IsNull(AdoPagoDetalle.Recordset("monto_bolivianos")) Then TxtMonto.Text = AdoPagoDetalle.Recordset("monto_bolivianos")
    'If Not IsNull(AdoPagoDetalle.Recordset("deducciones")) Then TxtDeducciones.Text = AdoPagoDetalle.Recordset("deducciones")
    If Not IsNull(AdoPagoDetalle.Recordset("fecha_pago")) Then TxtFechaPago.Text = AdoPagoDetalle.Recordset("fecha_pago")
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "D" Then
        OptDeposito.Value = True
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "C" Then
        OptChequeOrigen.Value = True
        TxtCuentaDestino.Visible = False
        LblCtaDestino.Visible = False
    End If
    If AdoPagoDetalle.Recordset("cheque_o_trf") = "T" Then
        OptTransferenciaOrigen.Value = True
        TxtCuentaDestino.Visible = True
        LblCtaDestino.Visible = True
    End If
  End If
End If
End Sub

Private Sub Form_Activate()
'Procedimiento Almacenado para determinar acumulados de los saldos
    'saldos_actuales
    BUSCA = 0
    'Colocando el nombre del usuario
    LblUsuario = GlUsuario
    Set rsPAgoDetalle = New ADODB.Recordset
    rsPAgoDetalle.Open "select * from Pago_detalle ", db, adOpenKeyset, adLockOptimistic
    Set AdoPagoDetalle.Recordset = rsPAgoDetalle
    Set DtGPagosParciales.DataSource = AdoPagoDetalle
    Set DtGPP.DataSource = AdoPagoDetalle

    Set rspago = New ADODB.Recordset
    rspago.Open "select * from pagos where (estado_contabilidad='P' or estado_devengado='S') and  (estado_aprobacion <>'A' or estado_aprobacion IS NULL) and (estado_pagado<>'S' or estado_pagado IS NULL) order by codigo_pago", db, adOpenKeyset, adLockOptimistic
    If rspago.RecordCount > 0 Then
      Set AdoPago.Recordset = rspago
      Set DtgPago.DataSource = AdoPago
      DtgPago.ReBind
    Else
      MsgBox "No existen registros !", vbInformation, "Validación de datos"
      DtgPago.Enabled = False
      Set DtGPagosParciales.DataSource = rsNada
      'Unload Me 'Exit Sub
    End If
    
    If rspago.RecordCount > 0 Then
        AdoPago.Recordset.MoveFirst
    End If
    
    Set rscuenta = New ADODB.Recordset
    rscuenta.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
    Set AdoCuenta.Recordset = rscuenta
    CtaAnterior = ""
    
    'Departamentos
    CmbNomDep.AddItem "LA PAZ"
    CmbNomDep.AddItem "ORURO"
    CmbNomDep.AddItem "POTOSI"
    CmbNomDep.AddItem "COCHABAMBA"
    CmbNomDep.AddItem "CHUQUISACA"
    CmbNomDep.AddItem "TARIJA"
    CmbNomDep.AddItem "PANDO"
    CmbNomDep.AddItem "BENI"
    CmbNomDep.AddItem "SANTA CRUZ"
    FraDatosCarta.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsNada.State = 1 Then rsNada.Close
    If rsPartida.State = 1 Then rsPartida.Close
    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
    'If rsPago.State = 1 Then rsPago.Close
    If rsControlDet.State = 1 Then rsControlDet.Close
    If rscuenta.State = 1 Then rscuenta.Close
    If rsbeneficiario.State = 1 Then rsbeneficiario.Close
    'If rsPagoDet.State = 1 Then rsPagoDet.Close
    If rsCtrlPago.State = 1 Then rsCtrlPago.Close
    If rsCuentaBancaria.State = 1 Then rsCuentaBancaria.Close
End Sub

Private Sub OptChequeOrigen_Click()
    TxtCuentaDestino.Visible = False
    LblCtaDestino.Visible = False
    LblNumeroOrigen.Caption = "Nro.Transacción"
    SSTTransferencia.TabVisible(0) = True
    SSTTransferencia.TabVisible(1) = False
    TxtBancoDestino.Visible = False
    LblBancoDestino.Visible = False
    'Departamento
    LblDepartamento.Visible = False
    CmbNomDep.Visible = False
    LblTransCheque.Caption = "CHEQUE/EFECTIVO"
End Sub

Private Sub OptDeposito_Click()
    TxtCuentaDestino.Visible = False
    LblCtaDestino.Visible = False
    LblNumeroOrigen.Caption = "Nro.Transacción"
    SSTTransferencia.TabVisible(0) = True
    SSTTransferencia.TabVisible(1) = False
    TxtBancoDestino.Visible = False
    LblBancoDestino.Visible = False
    'Departamento
    LblDepartamento.Visible = False
    CmbNomDep.Visible = False
    LblTransCheque.Caption = "DEPOSITO BANCARIO"
End Sub

'Private Sub OptColaImpresion_Click()
'    FrmColaImpresion.Show
'    FraImprimeCmpte.Visible = False
'End Sub

Private Sub OptObs1_Click()
    TxtObs.Text = "Transferencia o giro que deberá realizarse del Banco Unión según listado (registrado en UNI-SUELDO)."
End Sub

Private Sub OptObs2_Click()
    TxtObs.Text = "El costo de la comisión bancaria por la transferencia a realizar, debe ser descontado del monto a transferir."
End Sub

Private Sub OptObs3_Click()
'    LblObs.Visible = True
    TxtObs.Visible = True
End Sub

Private Sub OptSalirCmpte_Click()
    FraImprimeCmpte.Visible = False
End Sub

Private Sub OptSeleccion_Click()
'    FrmImprimeComprobanteNuevo.Show vbModal
'    FraImprimeCmpte.Visible = False
End Sub

Private Sub OptTransferenciaOrigen_Click()
    TxtCuentaDestino.Visible = True
    LblCtaDestino.Visible = True
    LblNumeroOrigen.Caption = "Nro.Transacción"
    SSTTransferencia.TabVisible(1) = True
    SSTTransferencia.TabVisible(0) = True
    'Departamento
    LblDepartamento.Visible = True
    CmbNomDep.Visible = True
    LblBancoDestino.Visible = True
    TxtBancoDestino.Visible = True
    FraObservaciones.Visible = True
    FraDatosCarta.Enabled = True
    LblTransCheque.Caption = "TRANSFERENCIA"
End Sub

Public Function Total()
   'Controlando monto a pagar
    SumaTotal = 0
'    SumaDOL = 0
    Set rsCtrlPago = New ADODB.Recordset
    rsCtrlPago.Open "select * from pago_detalle where ges_gestion='" & AdoPago.Recordset("ges_gestion") & "' and org_codigo='" & AdoPago.Recordset("org_codigo") & "' and codigo_pago='" & AdoPago.Recordset("codigo_pago") & "'", db, adOpenKeyset, adLockOptimistic
    NumReg = rsCtrlPago.RecordCount
    If rsCtrlPago.RecordCount > 0 Then
        While Not rsCtrlPago.EOF
            If Not IsNull(rsCtrlPago("monto_bolivianos")) Then
            SumaTotal = SumaTotal + rsCtrlPago("monto_bolivianos")
'            SumaDOL = SumaDOL + rsCtrlPago("monto_dolares")
            End If
            rsCtrlPago.MoveNext
        Wend
    End If
    rsCtrlPago.Close
    
End Function
Private Sub Cmdadicionar_Click()
On Error GoTo error_Adicionar:

    DtCCuentaOrigen.Text = ""
    DtcCtaTGN.Text = ""
    DtCCuentaOrigenDes.Text = ""
    TxtCuentaDestino.Text = ""
    TxtMontoBolivianos.Text = ""
'    TxtDeducciones.Text = ""
    txtmontoparcial.Text = ""
    FraControlPago.Enabled = True
'    FraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    rsControlPago.AddNew
Exit Sub
error_Adicionar:
    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub correlativo_cheque()

Dim NumeroCuenta As String
                Select Case DtCCuentaOrigen.Text
                    Case "4.41.1.1.1.402.208.11-2"
                          NumeroCuenta = "cta_1"
                    Case "4.41.1.1.1.402.208.12-1"
                          NumeroCuenta = "cta_2"
                    Case "4.41.1.1.1.402.208.14-0"
                          NumeroCuenta = "cta_3"
                    Case "4.41.1.1.1.402.208.16-8"
                          NumeroCuenta = "cta_4"
                    Case "4.41.1.1.1.402.208.18-6"
                          NumeroCuenta = "cta_5"
                    Case "4.41.1.1.1.402.254.01-7"
                          NumeroCuenta = "cta_6"
                    Case "4.41.1.1.1.402.254.02-6"
                          NumeroCuenta = "cta_7"
                    Case "1-297792"
                          NumeroCuenta = "cta_8"
                    Case "1-297809"
                          NumeroCuenta = "cta_9"
                    Case "1-297841"
                          NumeroCuenta = "cta_10"
                    Case "1-297867"
                          NumeroCuenta = "cta_11"
                    Case "1-297875"
                          NumeroCuenta = "cta_12"
                    Case "1-297883"
                          NumeroCuenta = "cta_13"
                    Case "1-297891"
                          NumeroCuenta = "cta_14"
                    Case "1-297916"
                          NumeroCuenta = "cta_15"
                    Case "1-297924"
                          NumeroCuenta = "cta_16"
                    Case "1-297932"
                          NumeroCuenta = "cta_17"
                    Case "1-297940"
                          NumeroCuenta = "cta_18"
                    Case "1-297958"
                          NumeroCuenta = "cta_19"
                    Case "1-301973"
                          NumeroCuenta = "cta_20"
                    Case "1-301999"
                          NumeroCuenta = "cta_21"
                    Case "1-302731"
                          NumeroCuenta = "cta_22"
                    Case "1-303515"
                          NumeroCuenta = "cta_23"
                    Case "1-306379"
                          NumeroCuenta = "cta_24"
                    Case "1-302731"
                          NumeroCuenta = "cta_25"
                 End Select
                          
         'Abriendo correlativo para hallar el numero de cheque
         If rsCorrel.State = 1 Then rsCorrel.Close
         Set rsCorrel = New ADODB.Recordset
         rsCorrel.Open "SELECT * FROM fc_correl WHERE tipo_tramite= '" & NumeroCuenta & "' ", db, adOpenKeyset, adLockOptimistic
         If rsCorrel.RecordCount > 0 Then
            rsCorrel("numero_correlativo") = rsCorrel("numero_correlativo") + 1
            rsCorrel.Update
         Else
            rsCorrel("numero_correlativo") = 0
            rsCorrel.Update
         End If
         'MsgBox "Se imprimirá el Nro. de cheque ....   " & rsCorrel("numero_correlativo"), vbInformation, "Información"
         TxtNoTransaccion.Text = rsCorrel("numero_correlativo")
         
End Sub



Public Sub Cmpte_NroTransferencia()
'Esto en el caso de realizarlo por selección
If rsCheque.State = 1 Then rsCheque.Close
Set rsCheque = New ADODB.Recordset
rsCheque.Open "select * FROM ts_cheque", db, adOpenKeyset, adLockOptimistic
If rsCheque.RecordCount > 0 Then
        While Not rsCheque.EOF
            Set rsPagoDet = New ADODB.Recordset
            rsPagoDet.Open "select * from pago_detalle where codigo_pago='" & rsCheque("numero_comprobante") & "' and estado_aprobacion='N'", db, adOpenKeyset, adLockOptimistic
                Select Case Len(rsCheque("numero_cheque"))
                    Case 1
                        NumeroCheque = "0000" + rsCheque("numero_cheque")
                    Case 2
                        NumeroCheque = "000" + rsCheque("numero_cheque")
                    Case 3
                        NumeroCheque = "00" + rsCheque("numero_cheque")
                    Case 4
                        NumeroCheque = "0" + rsCheque("numero_cheque")
                    Case 5
                        NumeroCheque = rsCheque("numero_cheque")
                End Select
                
                rsPagoDet("numero_cheque_trf") = NumeroCheque
                rsPagoDet.Update

            rsCheque.MoveNext
        Wend
End If

End Sub


Public Sub ESTADO_APROBADO(CodigoPago As String, CodigoOrg As String, GesGestion As String)
Dim SumaMontosParciales As Long
        'Determinando comprobante de pagos en detalle como APROBADOS CHEQUES y en pago
            Set rspago = New ADODB.Recordset
            If rspago.State = 1 Then rspago.Close
            rspago.Open "SELECT * from pagos where codigo_pago= '" & CodigoPago & "' and ges_gestion= '" & GesGestion & "' and org_codigo='" & CodigoOrg & "'", db, adOpenKeyset, adLockOptimistic
            If rspago.RecordCount > 0 Then
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & CodigoPago & "' and ges_gestion= '" & GesGestion & "' and org_codigo='" & CodigoOrg & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                     'rsPagoDetalle("estado_aprobacion") = "A"
                     rsPAgoDetalle.Update
                End If
                Set rsPAgoDetalle = New ADODB.Recordset
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                rsPAgoDetalle.Open "SELECT * from pago_detalle where codigo_pago= '" & CodigoPago & "' and estado_aprobacion<>'A' and ges_gestion= '" & GesGestion & "' and org_codigo='" & CodigoOrg & "'", db, adOpenKeyset, adLockOptimistic
                If rsPAgoDetalle.RecordCount > 0 Then
                    SumaMontosParciales = 0
                    While Not rsPAgoDetalle.EOF
                         SumaMontosParciales = SumaMontosParciales + rsPAgoDetalle("monto_bolivianos")
                         rsPAgoDetalle.MoveNext
                    Wend
                    If rspago("liquido_pagar") = SumaMontosParciales And SumaMontosParciales <> 0 Then
                     'rsPago("estado_aprobacion") = "A"
                     rspago("estado_pagado") = "S" 'Total
                     rspago.Update
                    Else
                     rspago("estado_pagado") = "P" 'Parcial
                     rspago.Update
                    End If
                End If
                If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
           End If
End Sub


Public Sub saldos_actuales()
'Primera forma de llamar procedimientos almacenados
' SaldoIBs = db.Parameters("GastoBs")
' db.gastos Format(Date, "dd/mm/yyyy"), Format(Date, "dd/mm/yyyy")

'Ejemplo de ...
  Dim Tsum_829 As New ADODB.Parameter
  Dim Tsum_2676 As New ADODB.Parameter
  Dim Tsum_0922 As New ADODB.Parameter
  Dim Tsum_0921 As New ADODB.Parameter
  Dim Tsum_0873 As New ADODB.Parameter
  Dim Tsum_0872 As New ADODB.Parameter
  Dim Tsum_0870 As New ADODB.Parameter
  Dim Tsum_0869 As New ADODB.Parameter
  Dim Tsum_1_306479 As New ADODB.Parameter
  Dim Tsum_1_303515 As New ADODB.Parameter
  Dim Tsum_1_302731 As New ADODB.Parameter
  Dim Tsum_1_301999 As New ADODB.Parameter
  Dim Tsum_1_301973 As New ADODB.Parameter
  Dim Tsum_1_297958 As New ADODB.Parameter
  Dim Tsum_1_297940 As New ADODB.Parameter
  Dim Tsum_1_297932 As New ADODB.Parameter
  Dim Tsum_1_297924 As New ADODB.Parameter
  Dim Tsum_1_297916 As New ADODB.Parameter
  Dim Tsum_1_297891 As New ADODB.Parameter
  Dim Tsum_1_297883 As New ADODB.Parameter
  Dim Tsum_1_297875 As New ADODB.Parameter
  Dim Tsum_1_297867 As New ADODB.Parameter
  Dim Tsum_1_297841 As New ADODB.Parameter
  Dim Tsum_1_297809 As New ADODB.Parameter
  'Dim Tsum_1_297792 As New ADODB.Parameter
Set comCuentasAcumuladas = New ADODB.Command
With comCuentasAcumuladas
    .CommandText = "Cuentas_Acumuladas"
    .CommandType = adCmdStoredProc
    
    Set Tsum_829 = .CreateParameter("sum_829", adCurrency, adParamOutput)
    .Parameters.Append Tsum_829
    Set Tsum_2676 = .CreateParameter("sum_2676 ", adCurrency, adParamOutput)
    .Parameters.Append Tsum_2676
    Set Tsum_0922 = .CreateParameter("sum_0922", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0922
    Set Tsum_0921 = .CreateParameter("sum_0921", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0921
    Set Tsum_0873 = .CreateParameter("sum_0873", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0873
    Set Tsum_0870 = .CreateParameter("sum_0870", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0870
    Set Tsum_0869 = .CreateParameter("sum_0869", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0869
    Set Tsum_0870 = .CreateParameter("sum_0870", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0870
    Set Tsum_1_306479 = .CreateParameter("sum_1_306479", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_306479
    Set Tsum_1_303515 = .CreateParameter("sum_1_303515", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_303515
    Set Tsum_1_302731 = .CreateParameter("sum_1_302731", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_302731
    Set Tsum_1_301999 = .CreateParameter("sum_1_301999", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_301999
    Set Tsum_1_301973 = .CreateParameter("sum_1_301973", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_301973
    Set Tsum_1_297958 = .CreateParameter("sum_1_297958", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297958
    Set Tsum_1_297940 = .CreateParameter("sum_1_297940", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297940
    Set Tsum_1_297932 = .CreateParameter("sum_1_297932", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297932
    Set Tsum_1_297924 = .CreateParameter("sum_1_297924", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297924
    Set Tsum_1_297916 = .CreateParameter("sum_1_297916", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297916
    Set Tsum_1_297891 = .CreateParameter("sum_1_297891", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297891
    Set Tsum_1_297883 = .CreateParameter("sum_1_297883", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297883
    Set Tsum_1_297875 = .CreateParameter("sum_1_297875", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297875
    Set Tsum_1_297867 = .CreateParameter("sum_1_297867", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297867
    Set Tsum_1_297841 = .CreateParameter("sum_1_297841", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297841
    Set Tsum_1_297809 = .CreateParameter("sum_1_297809", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297809
    'Set Tsum_1_297792 = .CreateParameter("sum_1_297792", adCurrency, adParamOutput)
    .ActiveConnection = db
    .Execute
    'MsgBox Tsum_829
End With
      

End Sub

Public Sub SaldoReal_CtaBancaria()
'Ejemplo de ...
  Dim Tsum_829 As New ADODB.Parameter
  Dim Tsum_2676 As New ADODB.Parameter
  Dim Tsum_0922 As New ADODB.Parameter
  Dim Tsum_0921 As New ADODB.Parameter
  Dim Tsum_0873 As New ADODB.Parameter
  Dim Tsum_0872 As New ADODB.Parameter
  Dim Tsum_0870 As New ADODB.Parameter
  Dim Tsum_0869 As New ADODB.Parameter
  Dim Tsum_1_306479 As New ADODB.Parameter
  Dim Tsum_1_303515 As New ADODB.Parameter
  Dim Tsum_1_302731 As New ADODB.Parameter
  Dim Tsum_1_301999 As New ADODB.Parameter
  Dim Tsum_1_301973 As New ADODB.Parameter
  Dim Tsum_1_297958 As New ADODB.Parameter
  Dim Tsum_1_297940 As New ADODB.Parameter
  Dim Tsum_1_297932 As New ADODB.Parameter
  Dim Tsum_1_297924 As New ADODB.Parameter
  Dim Tsum_1_297916 As New ADODB.Parameter
  Dim Tsum_1_297891 As New ADODB.Parameter
  Dim Tsum_1_297883 As New ADODB.Parameter
  Dim Tsum_1_297875 As New ADODB.Parameter
  Dim Tsum_1_297867 As New ADODB.Parameter
  Dim Tsum_1_297841 As New ADODB.Parameter
  Dim Tsum_1_297809 As New ADODB.Parameter
  'Dim Tsum_1_297792 As New ADODB.Parameter
Set comCuentasAcumuladas = New ADODB.Command
With comCuentasAcumuladas
    .CommandText = "Cuentas_Acumuladas"
    .CommandType = adCmdStoredProc
    
    Set Tsum_829 = .CreateParameter("sum_829", adCurrency, adParamOutput)
    .Parameters.Append Tsum_829
    Set Tsum_2676 = .CreateParameter("sum_2676 ", adCurrency, adParamOutput)
    .Parameters.Append Tsum_2676
    Set Tsum_0922 = .CreateParameter("sum_0922", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0922
    Set Tsum_0921 = .CreateParameter("sum_0921", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0921
    Set Tsum_0873 = .CreateParameter("sum_0873", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0873
    Set Tsum_0870 = .CreateParameter("sum_0870", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0870
    Set Tsum_0869 = .CreateParameter("sum_0869", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0869
    Set Tsum_0870 = .CreateParameter("sum_0870", adCurrency, adParamOutput)
    .Parameters.Append Tsum_0870
    Set Tsum_1_306479 = .CreateParameter("sum_1_306479", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_306479
    Set Tsum_1_303515 = .CreateParameter("sum_1_303515", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_303515
    Set Tsum_1_302731 = .CreateParameter("sum_1_302731", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_302731
    Set Tsum_1_301999 = .CreateParameter("sum_1_301999", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_301999
    Set Tsum_1_301973 = .CreateParameter("sum_1_301973", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_301973
    Set Tsum_1_297958 = .CreateParameter("sum_1_297958", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297958
    Set Tsum_1_297940 = .CreateParameter("sum_1_297940", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297940
    Set Tsum_1_297932 = .CreateParameter("sum_1_297932", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297932
    Set Tsum_1_297924 = .CreateParameter("sum_1_297924", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297924
    Set Tsum_1_297916 = .CreateParameter("sum_1_297916", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297916
    Set Tsum_1_297891 = .CreateParameter("sum_1_297891", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297891
    Set Tsum_1_297883 = .CreateParameter("sum_1_297883", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297883
    Set Tsum_1_297875 = .CreateParameter("sum_1_297875", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297875
    Set Tsum_1_297867 = .CreateParameter("sum_1_297867", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297867
    Set Tsum_1_297841 = .CreateParameter("sum_1_297841", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297841
    Set Tsum_1_297809 = .CreateParameter("sum_1_297809", adCurrency, adParamOutput)
    .Parameters.Append Tsum_1_297809
    'Set Tsum_1_297792 = .CreateParameter("sum_1_297792", adCurrency, adParamOutput)
    .ActiveConnection = db
    .Execute
End With


'************************
'************************
'Abriendo tablas

Set rsCta = New ADODB.Recordset
If rsCta.State = 1 Then rsCta.Close
rsCta.Open "select * from fc_cuenta_bancaria ", db, adOpenKeyset, adLockOptimistic


'Actualizar
'db.Execute "UPDATE nombre_tabla SET CAMPO=VALUE, CAMPO=VALOR"
db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_829) & " WHERE cta_codigo = '0869'"
db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_2676) & " WHERE cta_codigo = '2676'"
db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_829) & " WHERE cta_codigo = '829'"
If Not IsNull(Tsum_0922) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_0922) & " WHERE cta_codigo = '0922'"
If Not IsNull(Tsum_0921) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_0921) & " WHERE cta_codigo = '0921'"
If Not IsNull(Tsum_0873) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado = " & Val(Tsum_0873) & " WHERE cta_codigo = '0873'"
If Not IsNull(Tsum_0870) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_0870) & " WHERE cta_codigo='0870'"
If Not IsNull(Tsum_0869) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_0869) & " WHERE cta_codigo='0869'"
If Not IsNull(Tsum_1_306479) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_306479) & " WHERE cta_codigo='1-306479'"
If Not IsNull(Tsum_1_303515) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_303515) & " WHERE cta_codigo='1-303515'"
If Not IsNull(Tsum_1_302731) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_302731) & " WHERE cta_codigo='1-302731'"
If Not IsNull(Tsum_1_302731) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_302731) & " WHERE cta_codigo='1-303515'"
If Not IsNull(Tsum_1_301999) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_301999) & " WHERE cta_codigo='1-301999'"
If Not IsNull(Tsum_1_301973) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_301973) & " WHERE cta_codigo='1-301973'"
If Not IsNull(Tsum_1_297958) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297958) & " WHERE cta_codigo='1-297958'"
If Not IsNull(Tsum_1_297940) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297940) & " WHERE cta_codigo='1-297940'"
If Not IsNull(Tsum_1_297932) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297932) & " WHERE cta_codigo='1-297932'"
If Not IsNull(Tsum_1_297924) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297924) & " WHERE cta_codigo='1-297924'"
If Not IsNull(Tsum_1_297916) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297916) & " WHERE cta_codigo='1-297916'"
If Not IsNull(Tsum_1_297891) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297891) & " WHERE cta_codigo='1-297891'"
If Not IsNull(Tsum_1_297883) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297883) & " WHERE cta_codigo='1-297883'"
If Not IsNull(Tsum_1_297875) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297875) & " WHERE cta_codigo='1-297875'"
If Not IsNull(Tsum_1_297867) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297867) & " WHERE cta_codigo='1-297867'"
If Not IsNull(Tsum_1_297883) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297883) & " WHERE cta_codigo='1-297883'"
If Not IsNull(Tsum_1_297875) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297875) & " WHERE cta_codigo='1-297875'"
If Not IsNull(Tsum_1_297867) Then db.Execute "UPDATE fc_cuenta_bancaria SET cta_acumulado=" & Val(Tsum_1_297867) & " WHERE cta_codigo='1-297867'"
Select Case DtCCuentaOrigen.Text
            Case "0869"
              If Tsum_0869 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "0870"
              If Tsum_0870 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
               End If
            Case "0872"
              If Tsum_0872 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "0873"
              If Tsum_0873 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "2676"
             If Tsum_2676 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "0922"
              If Tsum_0922 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "0921"
              If Tsum_0921 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297792"
'              If Tsum_1_297792 - Val(TxtMonto) < 0 Then
'                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
'                 Exit Sub
'              End If
            Case "1-297809"
              If Tsum_1_297809 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297841"
            If Tsum_1_297841 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297867"
            If Tsum_1_297867 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297875"
            If Tsum_1_297875 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297883"
            If Tsum_1_297883 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297891"
            If Tsum_1_297891 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297916"
            If Tsum_1_297916 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297924"
            If Tsum_1_297824 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297932"
            If Tsum_1_297932 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297940"
            If Tsum_1_297940 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-297958"
            If Tsum_1_297958 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-301973"
            If Tsum_1_301973 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-301999"
            If Tsum_1_301999 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-302731"
            If Tsum_1_302731 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-303515"
            If Tsum_1_303515 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-306379"
            If Tsum_1_306379 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
            Case "1-302731"
            If Tsum_1_302731 - Val(TxtMonto) < 0 Then
                 MsgBox "Cuenta con saldo insuficiente ", vbInformation + vbCritical
                 Exit Sub
              End If
         End Select

End Sub
Private Sub TxtCmpte_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
      Else
        KeyAscii = Asc(UCase(Chr(0)))
    End If
End Sub

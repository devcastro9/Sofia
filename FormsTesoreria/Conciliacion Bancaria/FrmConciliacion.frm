VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmConciliacion 
   Caption         =   "Comparaci�n de datos con Banco y QUEIROZ GALVAO"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Icon            =   "FrmConciliacion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   2880
      Left            =   -495
      TabIndex        =   55
      Top             =   10080
      Visible         =   0   'False
      Width           =   2925
      Begin VB.CommandButton CmdFechaGTZ 
         Caption         =   "Conciliar  Fecha GTZ"
         Height          =   720
         Left            =   1965
         MousePointer    =   4  'Icon
         Picture         =   "FrmConciliacion.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1785
         Width           =   945
      End
      Begin VB.CommandButton CmdSeleccion 
         Caption         =   "Selecci�n de registros de GTZ"
         Height          =   750
         Left            =   1050
         MousePointer    =   4  'Icon
         Picture         =   "FrmConciliacion.frx":186C
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   210
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimirTotales 
         Caption         =   "Imprimir"
         Height          =   750
         Left            =   1005
         Picture         =   "FrmConciliacion.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1005
         Width           =   960
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Limpiar"
         Height          =   720
         Left            =   1020
         Picture         =   "FrmConciliacion.frx":2878
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1800
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   1320
      TabIndex        =   8
      Top             =   1065
      Width           =   11985
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   360
         Left            =   435
         TabIndex        =   62
         Top             =   1050
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton OptTraspasos 
         Caption         =   "Traspasos"
         Height          =   300
         Left            =   435
         TabIndex        =   61
         Top             =   795
         Width           =   2115
      End
      Begin VB.OptionButton OptIngresos 
         Caption         =   "Ingresos"
         Height          =   345
         Left            =   435
         TabIndex        =   60
         Top             =   480
         Width           =   2205
      End
      Begin VB.OptionButton OptEgresos 
         Caption         =   "Egresos"
         Height          =   330
         Left            =   420
         TabIndex        =   59
         Top             =   180
         Width           =   2475
      End
      Begin VB.TextBox TxtCodigoBanco 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5145
         TabIndex        =   45
         Top             =   615
         Width           =   1005
      End
      Begin VB.TextBox TxtBanco 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6195
         TabIndex        =   44
         Top             =   615
         Width           =   4440
      End
      Begin MSComCtl2.DTPicker DTPInicio 
         Height          =   300
         Left            =   540
         TabIndex        =   16
         Top             =   1620
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   36670
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4665
         Left            =   285
         TabIndex        =   9
         Top             =   2160
         Width           =   11670
         _ExtentX        =   20585
         _ExtentY        =   8229
         _Version        =   393216
         Tabs            =   4
         TabHeight       =   520
         TabCaption(0)   =   "Cheques"
         TabPicture(0)   =   "FrmConciliacion.frx":2CBA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Saldos de Banco"
         TabPicture(1)   =   "FrmConciliacion.frx":2CD6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Saldos GTZ-UDAORE"
         TabPicture(2)   =   "FrmConciliacion.frx":2CF2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Aanlisis de Saldos"
         TabPicture(3)   =   "FrmConciliacion.frx":2D0E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.Frame Frame4 
            Height          =   4635
            Left            =   -74715
            TabIndex        =   25
            Top             =   1035
            Width           =   10680
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1995
               TabIndex        =   27
               Top             =   1530
               Width           =   2910
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   2010
               TabIndex        =   26
               Top             =   2445
               Width           =   2910
            End
            Begin VB.Label Label15 
               Caption         =   "Saldo Copiado"
               Height          =   315
               Left            =   870
               TabIndex        =   29
               Top             =   1095
               Width           =   2775
            End
            Begin VB.Label Label14 
               Caption         =   "Saldo Calculado"
               Height          =   315
               Left            =   900
               TabIndex        =   28
               Top             =   2115
               Width           =   2775
            End
         End
         Begin VB.Frame Frame2 
            Height          =   4635
            Left            =   -74625
            TabIndex        =   20
            Top             =   1020
            Width           =   10680
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   2010
               TabIndex        =   24
               Top             =   2445
               Width           =   2910
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   1995
               TabIndex        =   22
               Top             =   1530
               Width           =   2910
            End
            Begin VB.Label Label13 
               Caption         =   "Saldo Calculado"
               Height          =   315
               Left            =   900
               TabIndex        =   23
               Top             =   2115
               Width           =   2775
            End
            Begin VB.Label Label12 
               Caption         =   "Saldo Copiado"
               Height          =   315
               Left            =   870
               TabIndex        =   21
               Top             =   1095
               Width           =   2775
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Informe General"
            Height          =   3645
            Left            =   270
            TabIndex        =   10
            Top             =   870
            Width           =   9855
            Begin VB.Frame Frame6 
               Height          =   495
               Left            =   5460
               TabIndex        =   64
               Top             =   240
               Visible         =   0   'False
               Width           =   3630
               Begin VB.OptionButton optNoConciliados 
                  Caption         =   "No Conciliados"
                  Height          =   240
                  Left            =   1875
                  TabIndex        =   47
                  Top             =   165
                  Width           =   1650
               End
               Begin VB.OptionButton OptConciliados 
                  Caption         =   "Conciliados"
                  Height          =   195
                  Left            =   345
                  TabIndex        =   65
                  Top             =   180
                  Width           =   1620
               End
            End
            Begin VB.TextBox TxtTotalRegUDAPRE 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   2475
               TabIndex        =   38
               Top             =   645
               Width           =   1515
            End
            Begin VB.TextBox TxtTotalRegBanco 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   405
               TabIndex        =   37
               Top             =   645
               Width           =   1485
            End
            Begin VB.CommandButton CmdNoConciliados 
               Caption         =   " No Conciliados"
               Height          =   405
               Left            =   6375
               TabIndex        =   36
               Top             =   2475
               Width           =   2640
            End
            Begin VB.CommandButton CmdConciliados 
               Caption         =   "Conciliados"
               Height          =   405
               Left            =   6360
               TabIndex        =   35
               Top             =   1695
               Width           =   2640
            End
            Begin VB.TextBox TxtConciliados 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4950
               TabIndex        =   32
               Top             =   1710
               Width           =   1260
            End
            Begin VB.TextBox TxtNoConciliados 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4950
               TabIndex        =   31
               Top             =   2490
               Width           =   1275
            End
            Begin VB.CommandButton CmdNoCobrados 
               Caption         =   "Detalle No Cobrados"
               Height          =   405
               Left            =   2145
               TabIndex        =   15
               Top             =   1665
               Width           =   2640
            End
            Begin VB.TextBox TxtNoCobrados 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   315
               TabIndex        =   14
               Top             =   1725
               Width           =   1725
            End
            Begin MSAdodcLib.Adodc AdoCuenta 
               Height          =   390
               Left            =   120
               Top             =   2895
               Visible         =   0   'False
               Width           =   4665
               _ExtentX        =   8229
               _ExtentY        =   688
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
               Caption         =   "Cuenta Bancaria"
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
            Begin MSAdodcLib.Adodc AdoPagoDetalle 
               Height          =   390
               Left            =   4920
               Top             =   2895
               Visible         =   0   'False
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   688
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
               Caption         =   "Pago Detalle"
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
            Begin VB.Label Label19 
               Caption         =   " Total  registros UDAPRE"
               Height          =   225
               Left            =   2400
               TabIndex        =   40
               Top             =   405
               Width           =   3120
            End
            Begin VB.Label Label18 
               Caption         =   "Total  registros banco"
               Height          =   225
               Left            =   375
               TabIndex        =   39
               Top             =   405
               Width           =   2700
            End
            Begin VB.Label Label17 
               Caption         =   "Total de registros conciliados"
               Height          =   225
               Left            =   4965
               TabIndex        =   34
               Top             =   1440
               Width           =   2700
            End
            Begin VB.Label Label16 
               Caption         =   "Total de registros NO conciliados"
               Height          =   225
               Left            =   4980
               TabIndex        =   33
               Top             =   2250
               Width           =   3120
            End
            Begin VB.Label Label3 
               Caption         =   "Total  de cheques/Transf. no cobrados"
               Height          =   225
               Left            =   345
               TabIndex        =   13
               Top             =   1425
               Width           =   3120
            End
         End
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   300
         Left            =   2820
         TabIndex        =   18
         Top             =   1620
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50528257
         CurrentDate     =   36670
      End
      Begin MSAdodcLib.Adodc AdoBanco 
         Height          =   390
         Left            =   9075
         Top             =   1305
         Visible         =   0   'False
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   688
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
         Caption         =   "Banco"
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
      Begin MSDataListLib.DataCombo DtCDescripcionBanco 
         Bindings        =   "FrmConciliacion.frx":2D2A
         DataField       =   "Bco_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   6900
         TabIndex        =   42
         Top             =   630
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Bco_descripcion_larga"
         BoundColumn     =   "Bco_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCodigoBanco 
         Bindings        =   "FrmConciliacion.frx":2D41
         DataField       =   "Bco_codigo"
         DataSource      =   "AdoCuenta"
         Height          =   315
         Left            =   5145
         TabIndex        =   43
         Top             =   630
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ListField       =   "Bco_codigo"
         BoundColumn     =   "Bco_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.Animation AVI 
         Height          =   795
         Left            =   10710
         TabIndex        =   48
         Top             =   390
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1402
         _Version        =   393216
         FullWidth       =   59
         FullHeight      =   53
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigen 
         Bindings        =   "FrmConciliacion.frx":2D58
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   5130
         TabIndex        =   50
         Top             =   1275
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCuentaOrigenDes 
         Bindings        =   "FrmConciliacion.frx":2D70
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   5130
         TabIndex        =   51
         Top             =   1620
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtcCtaTGN 
         Bindings        =   "FrmConciliacion.frx":2D88
         DataField       =   "cta_codigo"
         Height          =   315
         Left            =   7320
         TabIndex        =   52
         Top             =   1275
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "Cta_codigo_tgn"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin VB.Label Label20 
         Caption         =   "Banco"
         Height          =   225
         Left            =   5160
         TabIndex        =   46
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Fin"
         Height          =   225
         Left            =   2835
         TabIndex        =   19
         Top             =   1425
         Width           =   1410
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Inicio"
         Height          =   225
         Left            =   525
         TabIndex        =   17
         Top             =   1425
         Width           =   1410
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. "
         Height          =   195
         Left            =   5115
         TabIndex        =   12
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   225
         Left            =   6195
         TabIndex        =   11
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.Frame FraOpciones 
      Height          =   6975
      Left            =   30
      TabIndex        =   5
      Top             =   1065
      Width           =   1230
      Begin VB.CommandButton CmdConciliaCuentas 
         Caption         =   "Conciliaci�n"
         Height          =   750
         Left            =   75
         TabIndex        =   54
         Top             =   1020
         Width           =   1080
      End
      Begin VB.CommandButton CmdActualizaInformacion 
         Caption         =   "Actualiza Informaci�n"
         Height          =   750
         Left            =   75
         TabIndex        =   53
         Top             =   270
         Width           =   1080
      End
      Begin VB.CommandButton CmdFechaBanco 
         Caption         =   "Conciliar  Fecha BANCO"
         Height          =   750
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmConciliacion.frx":2DA0
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1035
         Width           =   975
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   795
         Left            =   75
         Picture         =   "FrmConciliacion.frx":3742
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6090
         Width           =   1095
      End
      Begin VB.CommandButton CmdConciliacionFechaUDAPRE 
         Caption         =   "Conciliar por fecha"
         Height          =   750
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "FrmConciliacion.frx":3B84
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton CmdConciliacionBanco 
         Caption         =   "Conciliar por fecha Banco"
         Height          =   720
         Left            =   135
         MousePointer    =   4  'Icon
         Picture         =   "FrmConciliacion.frx":4526
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   5940
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Conciliacion Bancaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   345
         Left            =   2160
         TabIndex        =   30
         Top             =   195
         Width           =   8265
      End
      Begin VB.Label LblUsuario 
         Caption         =   "LblUsuario"
         Height          =   225
         Left            =   10485
         TabIndex        =   4
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   9210
         TabIndex        =   3
         Top             =   645
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Index           =   0
         Left            =   1245
         TabIndex        =   2
         Top             =   690
         Width           =   2460
      End
      Begin VB.Label Label8 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   675
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FrmConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rsTesoreria As New ADODB.Recordset
Dim rsDatosBanco As New ADODB.Recordset
Dim rsPAgoDetalle As New ADODB.Recordset
Dim rscuenta As New ADODB.Recordset
Dim rsBANCO As New ADODB.Recordset
Dim rsDB As New ADODB.Recordset
Dim rsCta As New ADODB.Recordset
Dim rsBco As New ADODB.Recordset
Dim rsUnion As New ADODB.Recordset
Dim rsDG As New ADODB.Recordset
Dim rsPD As New ADODB.Recordset
Dim rsING As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset
Dim NroRegBanco As Double
Dim NroRegUdapre As Double
Dim NroRegConciliados  As Double
Dim NroRegNoConciliados  As Double
Dim NroRegNoCoincidentes  As Double

 Dim comConciliados As New ADODB.Command

Private Sub CmdActualizaInformacion_Click()
    'Se uniran las tablas Co_MovimientoPCO, pago_detalle, fo_ingresos
    db.Execute "delete from fc_datosGTZ"
    db.movimiento_Cuenta_Bancaria
    MsgBox "fin de proceso"
End Sub

Private Sub CmdAnulados_Click()
    FrmComparacion.LblTitulo.Caption = "A N U L A D O S"
    FrmComparacion.Show
End Sub

Private Sub CmdConciliacion_Click()
End Sub

Private Sub CmdDetalleGlobal_Click()
    FrmComparacion.LblTitulo.Caption = "I N F O R M A C I O N  G E N E R A L"
    FrmComparacion.Show
End Sub

Private Sub CmdConciliacionBanco_Click()
Dim SW As Integer
Dim condicion As String
Dim mes_numeral As Integer

NroRegBanco = 0
NroRegUdapre = 0
NroRegConciliados = 0
NroRegNoConciliados = 0
db.Execute "DELETE FROM to_DatosBanco"
MsgBox "Esperar el mensaje de terminado !!!!"
If swFiltro = "FECHA" Then
        'Abrir la tabla de los datos del banco
        Set rsDatosBanco = New ADODB.Recordset
        rsDatosBanco.Open "select * from fc_DatosBanco where month(Fecha_pago)='" & mes_numeral & "' and cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
        NroRegBanco = rsDatosBanco.RecordCount
        TxtTotalRegBanco.Text = NroRegBanco
                While Not rsDatosBanco.EOF
'                         If rsDatosBanco("fecha_liquidacion") > DTPInicio.Value And rsDatosBanco("fecha_liquidacion") < DTPFin.Value Then
'                         rsTesoreria.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
'                         "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                         Set rsTesoreria = New ADODB.Recordset
                         If rsTesoreria.State = 1 Then rsTesoreria.Close
                         rsTesoreria.Open "SELECT pagos.codigo_pago, pagos.ges_gestion, fc_cuenta_bancaria.Ges_gestion, fc_cuenta_bancaria.Bco_codigo, fc_cuenta_bancaria.Cta_codigo, pago_detalle.numero_cheque_trf, pago_detalle.fecha_pago  " & _
                                          "FROM ((pagos INNER JOIN pago_detalle ON (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.org_codigo = pago_detalle.org_codigo) AND (pagos.ges_gestion = pago_detalle.Ges_gestion)) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario ", db, adOpenKeyset, adLockOptimistic
                         NroRegUdapre = rsTesoreria.RecordCount
                         TxtTotalRegUDAPRE.Text = NroRegUdapre
                         While Not rsTesoreria.EOF
                             If rsTesoreria("fecha_pago") >= DTPInicio.Value And rsTesoreria("fecha_pago") <= DTPFin.Value And rsDatosBanco("cta_codigo") = DtCCuentaOrigen.Text Then
                                SW = 0
                                'Verificamos la coincidencia de codigos de pago del banco con tesoreria
                                If rsDatosBanco("Nro_Cheque") = rsTesoreria("numero_cheque_trf") And rsDatosBanco("cta_codigo") = rsTesoreria("cta_codigo") Then
                                              MsgBox "iguales"
        ''                                    If rsDatosBanco("Codigo_Org") <> rsTesoreria("Org_Codigo") Then
        ''                                        MsgBox "No coincide el organismo", vbInformation + vbCritical
        ''                                        MsgBox rsDatosBanco("Codigo_Org")
        ''                                        MsgBox rsTesoreria("Org_Codigo")
        ''                                    Else
        ''                                        sw = 1
        ''                                    End If
        ''                                    If rsDatosBanco("Cta_Codigo") <> rsTesoreria("Cta_Codigo") Then
        ''                                        MsgBox "No coincide la cuenta bancaria", vbInformation + vbCritical
        ''                                        MsgBox rsDatosBanco("Cta_Codigo")
        ''                                        MsgBox rsTesoreria("Cta_Codigo")
        ''                                    Else
        ''                                        sw = 1
        ''                                    End If
                                            If rsDatosBanco("importe") <> rsTesoreria("monto_bolivianos") Then
                                                MsgBox "No coincide el importe", vbInformation + vbCritical
                                                MsgBox rsDatosBanco("importe")
                                                MsgBox rsTesoreria("monto_bolivianos")
                                            Else
                                                SW = 1
                                            End If
                                            If SW <> 1 Then
                                                    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                                                    Set rsPAgoDetalle = New ADODB.Recordset
                                                    rsPAgoDetalle.Open "select * from Pago_Detalle", db, adOpenKeyset, adLockOptimistic
                                                    rsPAgoDetalle("estado_conciliacion") = "S"
                                                    rsPAgoDetalle.Update
                                                    NroRegConciliados = NroRegConciliados + 1
                                                    If rsDB.State = 1 Then rsDB.Close
                                                    Set rsDB = New ADODB.Recordset
                                                    rsDB.Open "select * from fc_DatosBanco", db, adOpenKeyset, adLockOptimistic
                                                    rsDB("estado_conciliacion") = "S"
                                                    rsDB.Update
                                            Else
                                                   NroRegNoCoincidentes = NroNocoincidentes + 1
                                            End If
                                     Else
                                      NroRegNoConciliados = NroRegNoConciliados + 1
                                End If
                       End If 'FECHA
                       rsTesoreria.MoveNext
                  Wend
               'End If  'FECHA
                    rsDatosBanco.MoveNext
        Wend
        TxtConciliados = NroRegConciliados
        TxtNoConciliados = NroRegNoConciliados
        TxtNoCoincidentes = NroRegNoCoincidentes
End If

If swFiltro = "MES" And CmbMes <> "" Then
        'Determinando el mes en numeral
            Select Case CmbMes.Text
                Case "ENERO"
                    mes_numeral = 1
                Case "FEBRERO"
                    mes_numeral = 2
                Case "MARZO"
                    mes_numeral = 3
                Case "ABRIL"
                    mes_numeral = 4
                Case "MAYO"
                    mes_numeral = 5
                Case "JUNIO"
                    mes_numeral = 6
                Case "JULIO"
                    mes_numeral = 7
                Case "AGOSTO"
                    mes_numeral = 8
                Case "SEPTIEMBRE"
                    mes_numeral = 9
                Case "OCTUBRE"
                    mes_numeral = 10
                Case "NOVIEMBRE"
                    mes_numeral = 11
                Case "DICIEMBRE"
                    mes_numeral = 12
            End Select
        
        'Abrir la tabla de los datos del banco
        Set rsDatosBanco = New ADODB.Recordset
        rsDatosBanco.Open "select * from fc_DatosBanco where month(Fecha_pago)='" & mes_numeral & "' and cta_codigo='" & DtCCuentaOrigen.Text & "' ", db, adOpenKeyset, adLockOptimistic
        NroRegBanco = rsDatosBanco.RecordCount
        TxtTotalRegBanco.Text = NroRegBanco
                While Not rsDatosBanco.EOF
                         Set rsTesoreria = New ADODB.Recordset
                         If rsTesoreria.State = 1 Then rsTesoreria.Close
                         rsTesoreria.Open "SELECT pagos.codigo_pago, pagos.ges_gestion, fc_cuenta_bancaria.Ges_gestion, fc_cuenta_bancaria.Bco_codigo, fc_cuenta_bancaria.Cta_codigo, pago_detalle.numero_cheque_trf  " & _
                                          "FROM ((pagos INNER JOIN pago_detalle ON (pagos.codigo_pago = pago_detalle.codigo_pago) AND (pagos.org_codigo = pago_detalle.org_codigo) AND (pagos.ges_gestion = pago_detalle.Ges_gestion)) LEFT JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo) INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario ", db, adOpenKeyset, adLockOptimistic
'                         rsTesoreria.Open "SELECT pago_detalle.numero_cheque_trf, fc_beneficiario.denominacion_beneficiario, pago_detalle.monto_Bolivianos, pago_detalle.codigo_pago,pago_detalle.monto_Dolares, pago_detalle.tipo_cambio, fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_codigo, pago_detalle.org_codigo " & _
'                         "FROM (pago_detalle INNER JOIN fc_beneficiario ON pago_detalle.codigo_beneficiario = fc_beneficiario.codigo_beneficiario) INNER JOIN fc_cuenta_bancaria ON pago_detalle.cta_codigo = fc_cuenta_bancaria.Cta_codigo where month(pago_detalle.fecha_pago)='" & mes_numeral & "' order by pago_detalle.fecha_pago", db, adOpenKeyset, adLockOptimistic
                         NroRegUdapre = rsTesoreria.RecordCount
                         TxtTotalRegUDAPRE.Text = NroRegUdapre
                         While Not rsTesoreria.EOF
                                SW = 0
                                'Verificamos la coincidencia de codigos de pago del banco con tesoreria
'                                If Not IsNull(rsTesoreria("numero_cheque_trf")) Then MsgBox rsTesoreria("numero_cheque_trf")
'                                If Not IsNull(rsDatosBanco("Nro_Cheque")) Then MsgBox rsDatosBanco("Nro_Cheque")


                                    '                                If Not IsNull(rsTesoreria("numero_cheque_trf")) And Not IsNull(rsDatosBanco("Nro_Cheque")) And rsTesoreria("numero_cheque_trf") = rsDatosBanco("Nro_Cheque") Then
                                    '                                       MsgBox "son iguales", vbInformation + vbCritical
                                    '                                End If


                                If rsDatosBanco("Nro_Cheque") = rsTesoreria("numero_cheque_trf") And rsDatosBanco("cta_codigo") = rsTesoreria("cta_codigo") And rsDatosBanco("cta_codigo") = DtCCuentaOrigen.Text Then
        ''                                    If rsDatosBanco("Codigo_Org") <> rsTesoreria("Org_Codigo") Then
        ''                                        MsgBox "No coincide el organismo", vbInformation + vbCritical
        ''                                        MsgBox rsDatosBanco("Codigo_Org")
        ''                                        MsgBox rsTesoreria("Org_Codigo")
        ''                                    Else
        ''                                        sw = 1
        ''                                    End If
        ''                                    If rsDatosBanco("Cta_Codigo") <> rsTesoreria("Cta_Codigo") Then
        ''                                        MsgBox "No coincide la cuenta bancaria", vbInformation + vbCritical
        ''                                        MsgBox rsDatosBanco("Cta_Codigo")
        ''                                        MsgBox rsTesoreria("Cta_Codigo")
        ''                                    Else
        ''                                        sw = 1
        ''                                    End If
                                           ' MsgBox "hola"
'                                            If rsDatosBanco("importe") <> rsTesoreria("monto_bolivianos") Then
'                                                MsgBox "No coincide el importe", vbInformation + vbCritical
'                                                MsgBox rsDatosBanco("importe")
'                                                MsgBox rsTesoreria("monto_bolivianos")
'                                            Else
'                                                sw = 1
'                                            End If
                                            If SW <> 1 Then
                                                    'Colocando el status de 'S' al estado de conciliaci�n de pago_detalle
                                                    Set rsPAgoDetalle = New ADODB.Recordset
                                                    If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                                                    rsPAgoDetalle.Open "select * from Pago_Detalle where numero_cheque_trf='" & rsDatosBanco("Nro_Cheque") & "' and cheque_o_trf='C' and cta_codigo='" & rsDatosBanco("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
                                                    If rsPAgoDetalle.RecordCount > 0 Then
                                                            rsPAgoDetalle("estado_conciliacion") = "S"
                                                            rsPAgoDetalle.Update
                                                            NroRegConciliados = NroRegConciliados + 1
                                                            If rsDB.State = 1 Then rsDB.Close
                                                            Set rsDB = New ADODB.Recordset
                                                            rsDB.Open "select * from fc_DatosBanco where Nro_Cheque='" & rsDatosBanco("Nro_Cheque") & "' and Transf_Cheq='C' and cta_codigo='" & rsDatosBanco("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
                                                            rsDB("estado_conciliacion") = "S"
                                                            rsDB.Update
                                                   End If
                                            Else
                                                   NroRegNoCoincidentes = NroNocoincidentes + 1
                                            End If
                                     Else
                                      NroRegNoConciliados = NroRegNoConciliados + 1
                                End If
                       rsTesoreria.MoveNext

                  Wend
                    rsDatosBanco.MoveNext
        Wend
        TxtConciliados = NroRegConciliados
        TxtNoConciliados = NroRegNoConciliados
        TxtNoCoincidentes = NroRegNoCoincidentes
        MsgBox "termin�"
End If


End Sub

Private Sub CmdConciliacionUDAPRE_Click()
End Sub

Private Sub CmdConciliacionFechaUDAPRE_Click()
'Validaci�n de fecha
If DtCCuentaOrigen.Text = "" Then
   MsgBox "Elija alguna cuenta", vbInformation + vbCritical, "Validaci�n de datos"
   Exit Sub
End If

If OptEgresos.Value = True Then op1 = "EGR"
If OptIngresos.Value = True Then op1 = "ING"
If OptTraspasos.Value = True Then op1 = "TRP"
If opttodos.Value = True Then op1 = "TDS"

  Set comConciliados = New ADODB.Command
  With comConciliados
        .CommandText = "Cel_Conciliacion_FechaGTZ"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 3, op1)
        .Parameters.Append op1
        Set Cta = .CreateParameter("Cuenta", adVarChar, adParamInput, 40, DtCCuentaOrigen.Text)
        .Parameters.Append Cta
        .ActiveConnection = db
        .Execute
    End With
MsgBox "Final de proceso"
End Sub

Private Sub CmdConciliaCuentas_Click()
Dim sql As String
Dim Resp As String
Dim rsTBanco As New ADODB.Recordset
Dim rsTGTZ As New ADODB.Recordset
Dim Total_Banco As Double
Dim Total_GTZ As Double
'validaci�n de datos
If DtCCuentaOrigen.Text = "" Then
   MsgBox "Introduzca c�digo de la cuenta !!", vbCritical + vbDefaultButton1, "Validaci�n de datos"
   Exit Sub
End If
'Totalizando registros
If rsTGTZ.State = 1 Then rsTGTZ.Close
rsTGTZ.Open "select * from pago_detalle where cta_codigo= '" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
If rsTGTZ.RecordCount > 0 Then
   Total_GTZ = rsTGTZ.RecordCount
Else
   Total_GTZ = 0
End If

If rsTBanco.State = 1 Then rsTGTZ.Close
rsTBanco.Open "select * from fc_datosBanco where cta_codigo= '" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
If rsTBanco.RecordCount > 0 Then
   Total_Banco = rsTBanco.RecordCount
Else
   Total_Banco = 0
End If

TxtTotalRegUDAPRE.Text = Total_GTZ
db.Execute "select count(*) as Total_Banco from fc_DatosBanco where cta_codigo= '1-297867'"
TxtTotalRegBanco = Total_Banco

Resp = MsgBox("Esta seguro de realizar la conciliaci�n?", vbInformation + vbYesNo)
If Resp = vbYes Then
         Set rsJoinDatos = New ADODB.Recordset
          If rsJoinDatos.State = 1 Then rsJoinDatos.Close
          sql = "SELECT fc_DatosBanco.Nro_Doc, " & _
                "fc_DatosBanco.Organismo, fc_DatosGTZ.Organismo, " & _
                "fc_DatosGTZ.Nro_Doc, fc_DatosBanco.*, fc_datosGTZ.* " & _
                "FROM fc_DatosBanco INNER JOIN " & _
                "fc_DatosGTZ ON " & _
                "fc_DatosBanco.Nro_Doc = fc_DatosGTZ.Nro_Doc and " & _
                "fc_DatosBanco.cta_codigo = fc_DatosGTZ.cta_codigo " & _
                "WHERE fc_DatosGTZ.procedencia=1 and fc_DatosGTZ.estado_conciliacion<>'S' and fc_DatosGTZ.cta_codigo='" & DtCCuentaOrigen.Text & "'"
                rsJoinDatos.Open sql, db, adOpenKeyset, adLockOptimistic
          If rsJoinDatos.RecordCount > 0 Then
          MsgBox rsJoinDatos.RecordCount
            While Not rsJoinDatos.EOF
                    'db.Execute "UPDATE pago_detalle set estado_conciliacion='S' where (numero_cheque_trf='" & rsJoinDatos("Nro_Doc") & "' and cta_codigo='" & rsJoinDatos("Cta_Codigo") & "' and ges_gestion='2002' and org_codigo='" & rsJoinDatos("organismo") & "' and codigo_pago='" & rsJoinDatos("nro_cmpte") & "')"
                    db.Execute "UPDATE fc_DatosBanco set estado_conciliacion='S' where Nro_Doc='" & rsJoinDatos("Nro_Doc") & "' and cta_codigo='" & rsJoinDatos("Cta_Codigo") & "'"
                    rsJoinDatos.MoveNext
            Wend
         
          End If
        MsgBox "Termina conciliaci�n de pagos", vbInformation + vbCritical, "Validaci�n de datos"
          
        '"fc_DatosBanco.Nro_Cmpte = fc_DatosGTZ.Nro_Cmpte AND "
        'Conciliando los ingresos (Con Cta y Nro de documento)
         Set rsJoinDatos = New ADODB.Recordset
          If rsJoinDatos.State = 1 Then rsJoinDatos.Close
          sql = "SELECT fc_DatosBanco.Nro_Doc, " & _
                "fc_DatosBanco.Organismo, fc_DatosGTZ.Organismo, " & _
                "fc_DatosGTZ.Nro_Doc " & _
                "FROM fc_DatosBanco INNER JOIN " & _
                "fc_DatosGTZ ON " & _
                "fc_DatosBanco.Organismo = fc_DatosGTZ.Organismo AND " & _
                "fc_DatosBanco.Nro_Doc = fc_DatosGTZ.Nro_Doc and " & _
                "fc_DatosBanco.cta_codigo= fc_DatosGTZ.cta_codigo " & _
                "WHERE fc_DatosGTZ.procedencia=2 and fc_DatosGTZ.estado_conciliacion<>'S'"
                
        rsJoinDatos.Open sql, db, adOpenKeyset, adLockOptimistic
          If rsJoinDatos.RecordCount > 0 Then
            While Not rsJoinDatos.EOF
                  db.Execute "UPDATE fo_ingresos set estado_conciliacion='S' where numero_documento='" & rsJoinDatos("Nro_Doc") & "' and cta_codigo='" & rsJoinDatos("Cta_Codigo") & "'"
                  db.Execute "UPDATE fc_DatosBanco set estado_conciliacion='S' where Nro_Doc='" & rsJoinDatos("Nro_Doc") & "' and cta_codigo='" & rsJoinDatos("Cta_Codigo") & "'"
                  rsJoinDatos.MoveNext
            Wend
          End If
          MsgBox "Termina conciliacion ingresos ", vbInformation + vbCritical, "Validaci�n de datos"
          Set rsJoinDatos = New ADODB.Recordset
          If rsJoinDatos.State = 1 Then rsJoinDatos.Close
          sql = "SELECT fc_DatosBanco.Nro_Doc, " & _
                "fc_DatosBanco.Organismo, fc_DatosGTZ.Organismo, " & _
                "fc_DatosGTZ.Nro_Doc " & _
                "FROM fc_DatosBanco INNER JOIN " & _
                "fc_DatosGTZ ON " & _
                "fc_DatosBanco.Organismo = fc_DatosGTZ.Organismo AND " & _
                "fc_DatosBanco.Nro_Doc = fc_DatosGTZ.Nro_Doc  and " & _
                "fc_DatosBanco.cta_codigo= fc_DatosGTZ.cta_codigo " & _
                "WHERE fc_DatosGTZ.procedencia=3"
                rsJoinDatos.Open sql, db, adOpenKeyset, adLockOptimistic
           If rsJoinDatos.RecordCount > 0 Then
            While Not rsJoinDatos.EOF
                  db.Execute "UPDATE CO_MOVIMIENTOPCO set estado_conciliacion='S'"
                  rsJoinDatos.MoveNext
            Wend
           End If
           MsgBox "Termina conciliacion pco's", vbInformation + vbCritical, "Validaci�n de datos"
End If
  
End Sub

Private Sub CmdConciliados_Click()

    'Validaci�n de fecha
    If FrmConciliacion.DtCCuentaOrigen.Text = "" Then
       MsgBox "Elija alguna cuenta", vbInformation + vbCritical, "Validaci�n de datos"
       Exit Sub
    End If
    
    If FrmConciliacion.DTPInicio.Value > FrmConciliacion.DTPFin.Value Or FrmConciliacion.DTPFin.Value < FrmConciliacion.DTPInicio.Value Then
         MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
         Exit Sub
    End If

    'Conciliados
    OptConciliados.Value = True
    
    FrmComparacion.LblTitulo.Caption = "C O N C I L I A D O S"
    FrmComparacion.TxtCompFIni.Text = DTPInicio.Value
    FrmComparacion.TxtCompFFin.Text = DTPFin.Value
    FrmComparacion.TxtCuentaBancaria.Text = DtCCuentaOrigen.Text
    FrmComparacion.Show
End Sub

Private Sub CmdDetalleNoCoincidente_Click()
'    swConciliados = "N"
'    'FrmComparacion.LblTitulo.Caption = "N O   C O N C I L I A D O S"
'    If swFiltro = "MES" Then
'        FrmComparacion.TxtCompMes.Text = CmbMes.Text
'    End If
'    If swFiltro = "FECHA" Then
'        FrmComparacion.TxtCompFIni.Text = DTPInicio.Value
'        FrmComparacion.TxtCompFFin.Text = DTPFin.Value
'    End If
'    FrmComparacion.TxtCuentaBancaria = DtCCuentaOrigen.Text
'    FrmComparacion.CmbCompA�o.Text = CmbA�o.Text
'    FrmComparacion.LblTitulo.Caption = LblTitulo.Caption
'    FrmComparacion.LblBanco.Caption = TxtBanco.Text
'    FrmComparacion.Show
End Sub

Private Sub CmdFechaGTZ_Click()
                ''''Dim sw As Integer
                ''''Dim condicion As String
                ''''Dim NroRegcoincidentes As Double
                ''''
                '''''AVI.Open "C:\Conciliacion Bancaria\AVIS\Search.avi"
                '''''AVI.Play
                ''''NroRegBanco = 0
                ''''NroRegUdapre = 0
                ''''NroRegConciliados = 0
                ''''NroRegNoConciliados = 0
                ''''MsgBox "Esperar el mensaje de terminado !!!!"
                ''''If swFiltro = "FECHA" Then
                ''''        'Abrir la tabla de los datos del banco
                ''''        Set rsUnion = New ADODB.Recordset
                ''''        If rsUnion.State = 1 Then rsUnion.Close
                ''''           rsUnion.Open "SELECT * FROM fc_DatosGTZ", db, adOpenKeyset, adLockOptimistic
                ''''           NroRegUdapre = rsUnion.RecordCount
                ''''           TxtTotalRegUDAPRE.Text = NroRegUdapre
                ''''           While Not rsUnion.EOF
                ''''                    If rsUnion("fecha_pago") >= DTPInicio.Value And rsUnion("fecha_pago") <= DTPFin.Value Then
                ''''                         Set rsDatosBanco = New ADODB.Recordset
                ''''                         rsDatosBanco.Open "select * from fc_DatosBanco where cta_codigo='" & DtCCuentaOrigen.Text & "' and estado_conciliacion<>'S' ", db, adOpenKeyset, adLockOptimistic
                ''''                         NroRegBanco = rsDatosBanco.RecordCount
                ''''                         TxtTotalRegBanco.Text = NroRegBanco
                ''''                         While Not rsDatosBanco.EOF
                ''''                                sw = 0
                ''''                                'Verificamos la coincidencia de codigos de pago del banco con tesoreria
                ''''                                        If rsUnion("Nro_Doc") = rsDatosBanco("Nro_Doc") And rsUnion("cta_codigo") = rsDatosBanco("cta_codigo") Then
                ''''                                                    'If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
                ''''                                                    Set rsPAgoDetalle = New ADODB.Recordset
                ''''                                                    rsPAgoDetalle.Open "select * from Pago_Detalle where numero_cheque_trf= '" & rsUnion("Nro_Doc") & "' and cta_codigo='" & DtCCuentaOrigen.Text & "' ", db, adOpenKeyset, adLockOptimistic
                ''''                                                    rsPAgoDetalle("estado_conciliacion") = "S"
                ''''                                                    rsPAgoDetalle.Update
                ''''                                                    NroRegConciliados = NroRegConciliados + 1
                ''''                                                    If rsDB.State = 1 Then rsDB.Close
                ''''                                                    Set rsDB = New ADODB.Recordset
                ''''                                                    rsDB.Open "select * from fc_DatosBanco where Nro_Doc = '" & rsUnion("Nro_Doc") & "' and cta_codigo = '" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
                ''''                                                    rsDB("estado_conciliacion") = "S"
                ''''                                                    NroRegConciliados = NroRegNoConciliados + 1
                ''''                                                    rsDB.Update
                ''''                                            Else
                ''''                                                   NroRegcoincidentes = NroRegcoincidentes + 1
                ''''                                            End If
                ''''                               rsDatosBanco.MoveNext
                ''''                          Wend
                ''''                  End If
                ''''            rsUnion.MoveNext
                ''''        Wend
                ''''        TxtConciliados = NroRegConciliados
                ''''        TxtNoConciliados = NroRegNoConciliados
                ''''        'TxtNoCoincidentes = NroRegNoCoincidentes
                ''''End If
            
End Sub

Private Sub CmdImprimirTotales_Click()
    db.Execute "DELETE FROM fc_DatosBanco"
'    MsgBox "Exitoso"

'No se puede por intervalo de fecha
x = "05/05/2002"

db.Execute "insert into fc_DatosBanco (Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario,  Nro_Doc, Bco_Codigo, Transf_Cheq, cta_codigo) values ('01461','111','" & x & "', 16000,6.05,'3462629LP','04010','uni','C','1-297809' ) "
db.Execute "insert into fc_DatosBanco (Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario,  Nro_Doc, Bco_Codigo, Transf_Cheq, cta_codigo) values ('00364','565','" & x & "'  , 46000,6.13,'3462629LP','02353','uni','C','1-297867' ) "
db.Execute "insert into fc_DatosBanco (Nro_Cmpte, Organismo, Fecha_Pago, Monto, Cambio, Beneficiario,  Nro_Doc, Bco_Codigo, Transf_Cheq, cta_codigo) values ('00014','111','" & x & "' , 56000,6.08,'3462629LP','00157','bcb','T','1-297809' ) "

     'Determinar las cuentas
      Set rssaf = New ADODB.Recordset
      rssaf.Open "select * from fc_DatosBanco ", db, adOpenKeyset, adLockOptimistic
      If rssaf.RecordCount > 0 Then
        While Not rssaf.EOF
          MsgBox rssaf("nro_cmpte") + " " + rssaf("organismo") + " " + rssaf("cta_codigo") + " " + rssaf("Transf_Cheq") + " " + rssaf("Nro_DOc") + " " + Format(rssaf("fecha_pago"), "ddmmyyyy") + rssaf("bco_Codigo")
          rssaf.MoveNext
        Wend
      End If
End Sub

Private Sub CmdModificar_Click()
    db.Execute "DELETE FROM to_DatosBanco"
    TxtTotalRegBanco.Text = ""
    TxtTotalRegUDAPRE.Text = ""
    
End Sub

Private Sub CmdNoCobrados_Click()
    FrmComparacion.LblTitulo.Caption = "N O   C O B R A D O S"
    FrmComparacion.Show
End Sub
Private Sub CmdNoConciliados_Click()


    'Validaci�n de fecha
    If FrmConciliacion.DtCCuentaOrigen.Text = "" Then
       MsgBox "Elija alguna cuenta", vbInformation + vbCritical, "Validaci�n de datos"
       Exit Sub
    End If
    
    If FrmConciliacion.DTPInicio.Value > FrmConciliacion.DTPFin.Value Or FrmConciliacion.DTPFin.Value < FrmConciliacion.DTPInicio.Value Then
         MsgBox "Seleccione un rango de fechas correcto", vbCritical + vbDefaultButton1
         Exit Sub
    End If


    'No Conciliados
    optNoConciliados.Value = True
    

    'procedimiento Almacenado
    If OptEgresos.Value = True Then op1 = "EGR"
    If OptIngresos.Value = True Then op1 = "ING"
    If OptTraspasos.Value = True Then op1 = "TRP"
    If opttodos.Value = True Then op1 = "TDS"

    Set comConciliados = New ADODB.Command
    With comConciliados
        .CommandText = "Cel_NoConciliacion_FechaGTZ"
        .CommandType = adCmdStoredProc
        Set fecha1 = .CreateParameter("FechaIni", adVarChar, adParamInput, 10, DTPInicio.Value)
        .Parameters.Append fecha1
        Set fecha2 = .CreateParameter("FechaFin", adVarChar, adParamInput, 10, DTPFin.Value)
        .Parameters.Append fecha2
        Set op1 = .CreateParameter("Opcion", adVarChar, adParamInput, 3, op1)
        .Parameters.Append op1
        Set Cta = .CreateParameter("Cuenta", adVarChar, adParamInput, 40, DtCCuentaOrigen.Text)
        .Parameters.Append Cta
        .ActiveConnection = db
        .Execute
    End With



    FrmComparacion.TxtCompFIni.Text = DTPInicio.Value
    FrmComparacion.TxtCompFFin.Text = DTPFin.Value
    FrmComparacion.TxtCuentaBancaria.Text = DtCCuentaOrigen.Text
    FrmComparacion.LblTitulo.Caption = LblTitulo.Caption + " - no conciliados"
    FrmComparacion.LblBanco.Caption = TxtBanco.Text
    FrmComparacion.CmdImprimirBanco.Caption = "Imprimir NO Conciliados Banco"
    FrmComparacion.CmdImprimirGTZ.Caption = "Imprimir NO Conciliados GTZ"
    
    
    
FrmComparacion.Show
End Sub

Private Sub CmdRevertidos_Click()
    FrmComparacion.LblTitulo.Caption = "R E V E R T I D O S"
    FrmComparacion.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DataCombo2_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText

End Sub

Private Sub DtCCuentaDes_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
End Sub

Private Sub Command1_Click()
End Sub

Private Sub CmdSeleccion_Click()
        'Borrando el contenido de la tabla temporal
        db.Execute "DELETE FROM fc_DatosGTZ"
        'Ejecutando avi de b�squeda
        'AVI.Open "C:\Conciliacion Bancaria\AVIS\Search.avi"
        'AVI.Play
                    'Abriendo fc_DatosGTZ
                    Set rsDG = New ADODB.Recordset
                    rsDG.Open "select * from fc_DatosGTZ", db, adOpenKeyset, adLockOptimistic
                               'Primera opcion del case
                               If swFiltro = "FECHA" Then
                                       'Buscando los registros de pago_detalle
                                       Set rsPD = New ADODB.Recordset
                                       'rsPD.Open "select * from pago_detalle WHERE estado_conciliacion<> 'S'", db, adOpenKeyset, adLockOptimistic
                                       rsPD.Open "select * from pago_detalle ", db, adOpenKeyset, adLockOptimistic
                                       If rsPD.RecordCount > 0 Then
                                       While Not rsPD.EOF
                                            'If rsPD("fecha_pago") >= DTPInicio.Value And rsPD("fecha_pago") <= DTPFin.Value And rsDatosBanco("cta_codigo") = DtCCuentaOrigen.Text Then
                                                rsDG.AddNew
                                                rsDG("Nro_Cmpte") = rsPD("codigo_pago")
                                                rsDG("Organismo") = rsPD("org_codigo")
                                                If Not IsNull(rsDG("Fecha_pago")) Then
                                                    rsDG("Fecha_pago") = Format(rsPD("Fecha_pago"), "dd/mm/yyyy")
                                                End If
                                                rsDG("Monto") = rsPD("Monto_Bolivianos")
                                                rsDG("Cambio") = rsPD("Tipo_Cambio")
                                                rsDG("Beneficiario") = rsPD("codigo_beneficiario")
                                                rsDG("Nro_Doc") = rsPD("numero_cheque_trf")
                                                rsDG("Transf_Cheq") = rsPD("cheque_o_trf")
                                                rsDG("Cta_Codigo") = rsPD("Cta_Codigo")
                                                rsDG("estado_conciliacion") = "N"
                                                rsDG("status") = "E"
                                                rsDG.Update
                                                If Not IsNull(rsPD("cheque_o_trf")) And rsPD("cheque_o_trf") = "R" Then
                                                    rsPD.AddNew
                                                    rsDG("Nro_Cmpte") = rsPD("codigo_pago")
                                                    rsDG("Organismo") = rsPD("org_codigo")
                                                    'rsDG("Fecha_pago") = rsPD("Fecha_pago")
                                                    rsDG("Monto") = rsPD("Monto_Bolivianos")
                                                    rsDG("Cambio") = rsPD("Tipo_Cambio")
                                                    rsDG("Beneficiario") = rsPD("codigo_beneficiario")
                                                    rsDG("Nro_Doc") = rsPD("numero_cheque_trf")
                                                    rsDG("Transf_Cheq") = rsPD("cheque_o_trf")
                                                    rsDG("Cta_Codigo") = rsPD("Cta_Codigo")
                                                    rsDG("Cta_Codigo") = rsPD("Cta_Codigo")
                                                    rsDG("estado_conciliacion") = "N"
                                                    rsDG("status") = "I"
                                                    rsDG.Update
                                                 End If
                                             '   End If
                                             rsPD.MoveNext
                                       Wend
            
                                       Set rsING = New ADODB.Recordset
                                       rsING.Open "select * from fo_ingresos where estado_conciliacion<> 'S'", db, adOpenKeyset, adLockOptimistic
                                       If rsING.RecordCount > 0 Then
                                       While rsDG.EOF
                                            'If rsING("fecha_ingreso") >= DTPInicio.Value And rsPD("fecha_ingreso") <= DTPFin.Value And rsDatosBanco("cta_codigo") = DtCCuentaOrigen.Text Then
                                                rsDG.AddNew
                                                rsDG("Nro_Cmpte") = rsING("codigo_pago")
                                                rsDG("Organismo") = rsING("org_codigo")
                                                If Not IsNull(rsDG("Fecha_pago")) Then rsDG("Fecha_pago") = Format(rsING("Fecha_ingreso"), "dd/mmm/yyyy")
                                                rsDG("Monto") = rsING("Monto_bolivianos")
                                                rsDG("Cambio") = rsING("Tipo_Cambio")
                                                rsDG("Beneficiario") = rsPD("codigo_beneficiario")
                                                rsDG("Nro_Doc") = rsPD("numero_documento")
                                                rsDG("Transf_Cheq") = rsPD("cheque_o_trf")
                                                rsDG("Cta_Codigo") = rsPD("Cta_Codigo")
                                                rsDG("estado_conciliacion") = "N"
                                                rsDG("status") = "I"
                                                rsPD.Update
            
                                                rsDG.MoveNext
                                             'end if
                                       Wend
                                     End If
            
                               'Case "MES"
                        End If
                     End If
    AVI.Stop
End Sub

Private Sub DtCCodigoBanco_Click(Area As Integer)
    DtCDescripcionBanco.BoundText = DtCCodigoBanco.BoundText
End Sub

Private Sub DtcCtaTGN_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtcCtaTGN.BoundText
    DtCCuentaOrigen.BoundText = DtcCtaTGN.BoundText
End Sub


Private Sub DtCCuentaOrigen_Change()
        'Determinar los bancos
        Set rsCta = New ADODB.Recordset
        If rsCta.State = 1 Then rsCta.Close
        rsCta.Open "select * from fc_Cuenta_Bancaria where cta_codigo='" & DtCCuentaOrigen.Text & "'", db, adOpenKeyset, adLockOptimistic
        If rsCta.RecordCount > 0 Then
            Set rsBco = New ADODB.Recordset
            If rsBco.State = 1 Then rsCta.Close
            rsBco.Open "select * from fc_Bancos where Bco_codigo='" & rsCta("Bco_codigo") & "'", db, adOpenKeyset, adLockOptimistic
            If rsCta.RecordCount > 0 Then
                TxtBanco.Text = rsBco("Bco_descripcion_larga")
                TxtCodigoBanco.Text = rsBco("Bco_codigo")
            End If
        End If

End Sub

Private Sub DtCCuentaOrigen_Click(Area As Integer)
    DtCCuentaOrigenDes.BoundText = DtCCuentaOrigen.BoundText
    DtcCtaTGN.BoundText = DtCCuentaOrigen.BoundText
    
End Sub

Private Sub DtCCuentaOrigenDes_Click(Area As Integer)
   DtcCtaTGN.BoundText = DtCCuentaOrigenDes.BoundText
   DtCCuentaOrigen.BoundText = DtCCuentaOrigenDes.BoundText
End Sub

Private Sub DtCDescripcionBanco_Click(Area As Integer)
    DtCCodigoBanco.BoundText = DtCDescripcionBanco.BoundText
End Sub
Private Sub Form_Load()
        
        'Determinar las cuentas
        Set rscuenta = New ADODB.Recordset
        rscuenta.Open "select * from fc_cuenta_bancaria order by Cta_codigo_tgn", db, adOpenKeyset, adLockOptimistic
        Set AdoCuenta.Recordset = rscuenta
    
        'Determinar los bancos
        Set rsBANCO = New ADODB.Recordset
        rsBANCO.Open "select * from fc_Bancos ", db, adOpenKeyset, adLockOptimistic
        Set AdoBanco.Recordset = rsBANCO
        DtCDescripcionBanco.BoundText = DtCCodigoBanco.BoundText
        
        'Abriendo Pago Detalle
        'If rsPAgoDetalle.State = 1 Then rsPAgoDetalle.Close
        Set rsPAgoDetalle = New ADODB.Recordset
        rsPAgoDetalle.Open "select * from Pago_Detalle", db, adOpenKeyset, adLockOptimistic
        Set AdoPagoDetalle.Recordset = rsPAgoDetalle
        
        DTPInicio.Value = Date
        DTPFin.Value = Date
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmCtaBco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Bancarias"
   ClientHeight    =   7800
   ClientLeft      =   540
   ClientTop       =   1830
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleMode       =   0  'User
   ScaleWidth      =   12486.77
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11790
      Begin VB.Label Label27 
         Caption         =   "Label3"
         Height          =   210
         Left            =   10230
         TabIndex        =   58
         Top             =   1095
         Width           =   1305
      End
      Begin VB.Label Label26 
         Caption         =   "HORA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9255
         TabIndex        =   57
         Top             =   825
         Width           =   900
      End
      Begin VB.Label Label11 
         Height          =   165
         Left            =   10185
         TabIndex        =   56
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   9195
         TabIndex        =   55
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "Label3"
         Height          =   225
         Left            =   1425
         TabIndex        =   25
         Top             =   1005
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad  Administrativa"
         Height          =   225
         Left            =   1365
         TabIndex        =   23
         Top             =   690
         Width           =   2595
      End
      Begin VB.Label Label2 
         Caption         =   "UNIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   105
         TabIndex        =   22
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CUENTAS BANCARIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   11565
      End
   End
   Begin VB.Frame Fradatos 
      Enabled         =   0   'False
      Height          =   6360
      Left            =   3840
      TabIndex        =   13
      Top             =   1365
      Width           =   8100
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmCta.frx":0000
         Left            =   2265
         List            =   "FrmCta.frx":0013
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   2775
         Width           =   540
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmCta.frx":002B
         Left            =   2265
         List            =   "FrmCta.frx":0035
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   3255
         Width           =   540
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FrmCta.frx":003F
         Left            =   2265
         List            =   "FrmCta.frx":0049
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   4200
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FrmCta.frx":0053
         Left            =   2265
         List            =   "FrmCta.frx":005D
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   3750
         Width           =   540
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   2325
         TabIndex        =   0
         Top             =   405
         Width           =   1080
      End
      Begin VB.TextBox Txtfecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   444
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   5970
         Width           =   1455
      End
      Begin VB.TextBox Txtusuario 
         Height          =   285
         Left            =   5280
         TabIndex        =   50
         Text            =   "Text6"
         Top             =   5970
         Width           =   1455
      End
      Begin VB.TextBox TxtHora 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   5970
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saldos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   105
         TabIndex        =   39
         Top             =   4500
         Width           =   7695
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   270
            Left            =   6015
            TabIndex        =   60
            Top             =   750
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   476
            _Version        =   393216
            Format          =   90832897
            CurrentDate     =   36724
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   270
            Left            =   1965
            TabIndex        =   59
            Top             =   750
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   476
            _Version        =   393216
            Format          =   90832897
            CurrentDate     =   36724
         End
         Begin VB.TextBox Text11 
            DataField       =   "Cta_saldo_actual"
            DataSource      =   "adocta"
            Height          =   285
            Left            =   6000
            TabIndex        =   11
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            DataField       =   "Cta_saldo_inicial"
            DataSource      =   "adocta"
            Height          =   285
            Left            =   1920
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Fecha Saldo Actual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   43
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label24 
            Caption         =   "Saldo Actual:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "Fecha Apertura:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Saldo Apertura:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   180
            TabIndex        =   40
            Top             =   360
            Width           =   1695
         End
      End
      Begin MSAdodcLib.Adodc AdoMoneda 
         Height          =   375
         Left            =   5565
         Top             =   2115
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Caption         =   "Adodc1"
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
      Begin VB.TextBox Text4 
         DataField       =   "Cta_codigo_tgn"
         DataSource      =   "adocta"
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc adoBancos 
         Height          =   330
         Left            =   5880
         Top             =   1440
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Caption         =   "Adodc1"
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
      Begin VB.TextBox Text3 
         DataField       =   "Cta_descripcion_larga"
         DataSource      =   "adocta"
         Height          =   435
         Left            =   2295
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   840
         Width           =   5415
      End
      Begin MSDataListLib.DataCombo dtcB 
         Bindings        =   "FrmCta.frx":0067
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Top             =   1800
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bco_codigo"
         BoundColumn     =   "Bco_codigo"
         Text            =   ""
      End
      Begin VB.TextBox text1 
         DataField       =   "Cta_codigo"
         DataSource      =   "adocta"
         Height          =   285
         Left            =   5160
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dtcMon 
         Bindings        =   "FrmCta.frx":007F
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   2280
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipo_moneda"
         BoundColumn     =   "Tipo_Moneda"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcBco 
         Bindings        =   "FrmCta.frx":0097
         Height          =   315
         Left            =   2295
         TabIndex        =   34
         Top             =   1800
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Bco_descripcion_larga"
         BoundColumn     =   "Bco_codigo"
         Text            =   ""
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   53
         Top             =   5730
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Hora_Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   52
         Top             =   5730
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5520
         TabIndex        =   51
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Origen de Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   2760
         Width           =   1980
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "10=Finanaciamiento Propio, 70=Credito, 80=Donaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   44
         Top             =   2805
         Width           =   4995
      End
      Begin VB.Label Label18 
         Caption         =   "1=Activa, 0=Inactiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2865
         TabIndex        =   38
         Top             =   3780
         Width           =   3300
      End
      Begin VB.Label Label17 
         Caption         =   "E=Cta. Especial, F=Cta. Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   37
         Top             =   4230
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "C=Cuenta Corriente, H=Caja de AHorro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2895
         TabIndex        =   36
         Top             =   3300
         Width           =   3465
      End
      Begin VB.Label Label5 
         Caption         =   "Gestión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Moneda de la Cta.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   2220
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   1980
      End
      Begin VB.Label Label23 
         Caption         =   "Estado de la Cta.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "Sigla Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Label Label21 
         Caption         =   "Clasificación Cta.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   4200
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Label Label14 
         Caption         =   "Denominación Cta.:"
         DataField       =   "Cta_descripcion_larga"
         DataSource      =   "adocta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Banco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6315
      Left            =   0
      TabIndex        =   12
      Top             =   1395
      Width           =   1305
      Begin VB.CommandButton Cmdimprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":00AF
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3840
         Width           =   1005
      End
      Begin VB.CommandButton Cmdborrar 
         Caption         =   "Borrar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":0719
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1965
         Width           =   1005
      End
      Begin VB.CommandButton Cmd_busqueda 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":0D83
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2880
         Width           =   1005
      End
      Begin VB.CommandButton Cmdeditar 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":11C5
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1140
         Width           =   1005
      End
      Begin VB.CommandButton Cmdadicionar 
         Caption         =   "Adicionar"
         Height          =   720
         Left            =   180
         Picture         =   "FrmCta.frx":1607
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   300
         Width           =   990
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   210
         Picture         =   "FrmCta.frx":1A49
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4800
         Width           =   1005
      End
      Begin VB.CommandButton Cmdsalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":1E8B
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4800
         Width           =   1005
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   840
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Cmdaceptar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   195
         Picture         =   "FrmCta.frx":22CD
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   285
         Width           =   1005
      End
   End
   Begin MSDataGridLib.DataGrid grdlista 
      Bindings        =   "FrmCta.frx":270F
      Height          =   6000
      Left            =   1305
      TabIndex        =   48
      Top             =   1410
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   10583
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "Ges_gestion"
         Caption         =   "Ges_gestion"
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
         DataField       =   "Bco_codigo"
         Caption         =   "Bco_codigo"
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
         DataField       =   "Cta_codigo"
         Caption         =   "Cta_codigo"
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
         DataField       =   "Cta_tipo"
         Caption         =   "Cta_tipo"
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
         DataField       =   "Cta_moneda"
         Caption         =   "Cta_moneda"
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
         DataField       =   "Cta_codigo_tgn"
         Caption         =   "Cta_codigo_tgn"
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
         DataField       =   "Cta_descripcion_larga"
         Caption         =   "Cta_descripcion_larga"
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
         DataField       =   "Cta_fecha_ape"
         Caption         =   "Cta_fecha_ape"
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
         DataField       =   "Cta_saldo_inicial"
         Caption         =   "Cta_saldo_inicial"
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
         DataField       =   "Fte_codigo"
         Caption         =   "Fte_codigo"
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
      BeginProperty Column10 
         DataField       =   "Cta_patron"
         Caption         =   "Cta_patron"
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
      BeginProperty Column11 
         DataField       =   "Cta_activo"
         Caption         =   "Cta_activo"
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
      BeginProperty Column12 
         DataField       =   "Cta_Fecha_saldo"
         Caption         =   "Cta_Fecha_saldo"
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
      BeginProperty Column13 
         DataField       =   "Campo1"
         Caption         =   "Campo1"
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
      BeginProperty Column14 
         DataField       =   "Cta_saldo_actual"
         Caption         =   "Cta_saldo_actual"
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
      BeginProperty Column15 
         DataField       =   "Cta_fecha_mod"
         Caption         =   "Cta_fecha_mod"
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
      BeginProperty Column16 
         DataField       =   "Cta_hora_mod"
         Caption         =   "Cta_hora_mod"
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
      BeginProperty Column17 
         DataField       =   "Cta_usuario_mod"
         Caption         =   "Cta_usuario_mod"
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
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adolista 
      Height          =   375
      Left            =   1305
      Top             =   7380
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Adodc1"
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
Attribute VB_Name = "FrmCtaBco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstctaban As New ADODB.Recordset
Dim rstbancos As New ADODB.Recordset
Dim rstmoneda As New ADODB.Recordset
Dim CAMPOS As ADODB.Field
'Para busqueda
'==== busquedas ====
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
'Dim queryinicial As String
'Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim sql_banco As String


Private Sub Adolista_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   If pRecordset.EOF Or pRecordset.BOF Then
'      cmdeditar.Enabled = False
'      cmdborrar.Enabled = False
      
      Text1.Text = Empty
      Text2.Text = Empty
      Text3.Text = Empty
      Text4.Text = Empty
      Text9.Text = ""
      Text11.Text = ""
      DTP1.Value = Date
      DTP2.Value = Date
      
     Exit Sub
   End If
   
'   cmdeditar.Enabled = True
'   cmdborrar.Enabled = True
   
   
   
   Select Case pRecordset.EditMode
      Case adEditInProgress
      Case adEditNone
         Text1.Text = IIf(IsNull(pRecordset("Cta_codigo")), "", pRecordset("Cta_codigo"))
         Text2.Text = IIf(IsNull(pRecordset("ges_gestion")), "", pRecordset("ges_gestion"))
         Text3.Text = IIf(IsNull(pRecordset("cta_descripcion_larga")), "", pRecordset("cta_descripcion_larga"))
         dtcB.Text = IIf(IsNull(pRecordset("bco_codigo")), "", pRecordset("bco_codigo"))
         rstbancos.MoveFirst
         rstbancos.Find "bco_codigo = '" & pRecordset!bco_codigo & "'"
         If Not rstbancos.EOF Then dtcBco.Text = rstbancos!Bco_descripcion_larga & "" Else dtcBco.Text = ""
         dtcMon.Text = IIf(IsNull(pRecordset("cta_moneda")), "", pRecordset("cta_moneda"))
         Text4.Text = IIf(IsNull(pRecordset("cta_codigo_tgn")), "", pRecordset("cta_codigo_tgn"))
         Combo2.Text = IIf(IsNull(pRecordset("cta_tipo")), "", pRecordset("cta_tipo"))
         Combo1.Text = IIf(IsNull(pRecordset("fte_codigo")), "", pRecordset("fte_codigo"))
         Combo3.Text = IIf(IsNull(pRecordset("cta_patron")), "", pRecordset("cta_patron"))
         Combo4.Text = IIf(IsNull(pRecordset("cta_activo")), "", pRecordset("cta_activo"))
         Text9.Text = IIf(IsNull(pRecordset("cta_saldo_inicial")), "", pRecordset("cta_saldo_inicial"))
         DTP1.Value = IIf(IsNull(pRecordset("cta_fecha_ape")), "", pRecordset("cta_fecha_ape"))
         Text11.Text = IIf(IsNull(pRecordset("cta_saldo_actual")), "", pRecordset("cta_saldo_actual"))
         DTP2.Value = IIf(IsNull(pRecordset("cta_fecha_saldo")), "", pRecordset("cta_fecha_saldo"))
         txtFecha.Text = IIf(IsNull(pRecordset("cta_fecha_mod")), "", pRecordset("cta_fecha_mod"))
         TxtHora.Text = IIf(IsNull(pRecordset("cta_hora_mod")), "", pRecordset("cta_hora_mod"))
         Txtusuario.Text = IIf(IsNull(pRecordset("cta_usuario_mod")), "", pRecordset("cta_usuario_mod"))
      Case adEditDelete
      Case adEditAdd
   End Select
   Adolista.Caption = CStr(Adolista.Recordset.AbsolutePosition) & " de " & CStr(Adolista.Recordset.RecordCount)
End Sub




Private Sub Cmd_BSalir_Click()
'Fra_Busqueda.Visible = False
End Sub

Private Sub cmd_Ejecutar_Click()
'  If ValidaCriterio(CmbCampo.Text, CmbOperador.Text, TxtValor.Text) = 2 Then
'        If (Not rstctaban.BOF) Then
'            rstctaban.MoveFirst
'            rstctaban.Find CmbCampo.Text & " " & CmbOperador.Text & " '" & TxtValor.Text & "'", , adSearchForward
'            cmd_Ejecutar.Enabled = True
'            Cmdeditar.Enabled = True
'            Cmdborrar.Enabled = True
'
'         End If
'    Else
'      MsgBox errCriterio, vbExclamation, "ERROR"
'
'    End If
End Sub
   
Private Sub cmdAceptar_Click()
Dim SQL_FOR As String
Dim rstbanaux As New ADODB.Recordset
Dim SW As Boolean
On Error GoTo errorAceptar
   With Adolista
              
                   If Text1 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text1.SetFocus
                        Exit Sub
                     End If
                       If Text2 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text2.SetFocus
                        Exit Sub
                     End If
                       If Text3 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text3.SetFocus
                        Exit Sub
                     End If
                      If Text4 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text4.SetFocus
                        Exit Sub
                     End If
                   If Text9 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text9.SetFocus
                        Exit Sub
                     End If
                     If Not IsNumeric(Text9.Text) Then
                      MsgBox "El Saldo de Apertura debe ser un Valor Numerico"
                      Text9.SetFocus
                      Exit Sub
                     End If
                     If Text11 = "" Then
                       MsgBox "INTRODUZCA DATOS"
                        Text11.SetFocus
                        Exit Sub
                     End If
                     
                    If Not IsNumeric(Text11.Text) Then
                      MsgBox "El Saldo Actual debe ser un Valor Numerico"
                      Text11.SetFocus
                      Exit Sub
                     End If
                 Set rstbanaux = New ADODB.Recordset
                 SQL_FOR = "select * from fc_cuenta_Bancaria where cta_codigo = '" & Text1.Text & "'"
                 rstbanaux.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic, adCmdText
                If rstbanaux.RecordCount > 0 And Text1.Enabled Then
                   SW = True
                    MsgBox " CODIGO DUPLICADO"
                    Text1.SetFocus
                    Exit Sub
                End If
                      db.BeginTrans
                      SW = False
                If Text1.Enabled Then
                    .Recordset.AddNew
                    .Recordset("cta_codigo") = Text1.Text
                End If
                    .Recordset("ges_gestion").Value = Text2.Text
                    .Recordset("cta_descripcion_larga").Value = Text3.Text
                    .Recordset("bco_codigo").Value = dtcB.BoundText
                    .Recordset("cta_moneda").Value = dtcMon.BoundText
                    .Recordset("cta_codigo_tgn").Value = Text4.Text
                    .Recordset("cta_tipo").Value = Trim(Combo2.Text)
                    .Recordset("fte_codigo").Value = Trim(Combo1.Text)
                    .Recordset("cta_patron").Value = Trim(Combo3.Text)
                    .Recordset("cta_activo").Value = Trim(Combo4.Text)
                    .Recordset("cta_saldo_inicial").Value = Format(Text9.Text, "###,###,##.0,00")
                    .Recordset("cta_fecha_ape").Value = Format(DTP1.Value, "dd/mm/yyyy")
                    .Recordset("cta_saldo_actual").Value = Format(Text11.Text, "###,###,##.0,00")
                    .Recordset("cta_fecha_saldo").Value = Format(DTP2.Value, "dd/mm/yyyy")
                    .Recordset("cta_usuario_mod").Value = frmLogin.txtUserName.Text
                    .Recordset("cta_fecha_mod").Value = Date
                    .Recordset("cta_hora_mod").Value = Format(Time, "hh:mm:ss")
                    .Recordset.Update
                    .Recordset.Requery
                    db.CommitTrans
                    
        
  End With
   Call Cmdadicionar_Click
     
   Call cmdCancelar_Click
   
   Exit Sub

errorAceptar:
   
   Call pErrorRst(db.Errors)
   
   Adolista.Recordset.CancelUpdate
   
   db.RollbackTrans
End Sub
' Private Sub text9_keypress(KeyAscii As interger)
'  KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
' End Sub
'  Private Sub text11_keypress(KeyAscii As interger)
'  KeyAscii = IIf(Chr(KeyAscii) Like "[0-9,'.']", KeyAscii, 0)
' End Sub
 
 Private Sub Cmdadicionar_Click()
   Text1.Enabled = True
   Adolista.Enabled = False
   'grdlista.Enabled = True
   fraDatos.Enabled = True
   
  ' cmdborrar.Visible = False
   Cmd_busqueda.Visible = False
   Cmdimprimir.Visible = False
   CmdSalir.Visible = False
   Cmdeditar.Visible = False
   CmdAdicionar.Visible = False

   cmdAceptar.Visible = True
   CmdCancelar.Visible = True
   
   Text1.Text = Empty
   Text2.Text = Empty
   Text3.Text = Empty
   Text4.Text = Empty
   Combo2.Text = Combo2.List(0)
   Combo1.Text = Combo1.List(0)
   Combo3.Text = Combo3.List(0)
   Combo4.Text = Combo4.List(0)
   Text9.Text = ""
   DTP1.Value = Date
   Text11.Text = ""
   DTP2.Value = Date
   Text2.SetFocus
End Sub

'Private Sub Cmdborrar_Click()
'   Dim Mensaje As String
'
'On Error GoTo errorDelete
'
'   Mensaje = "¿Borrar: " & _
'               Text1.Text & " " & _
'               Trim(Text3.Text) & "?"
'   If MsgBox(Mensaje, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar:") = vbYes Then
'      db.BeginTrans
'      adoLista.Recordset.Delete
'      db.CommitTrans
'   End If
'
'   Exit Sub
'errorDelete:
'
'   Dim e As ADODB.Error
'
'   For Each e In db.Errors
'      MsgBox "Error No. " & e.Number & " " & e.Description
'   Next
'
'   db.RollbackTrans
'
'End Sub

Private Sub Cmd_Busqueda_Click()
'fra_Busqueda.Visible = True
'Fradatos.Enabled = True
'dul
'    Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
'    Set ClBuscaGrid.Conexión = db
'    ClBuscaGrid.EsTdbGrid = False
'    Set ClBuscaGrid.GridTrabajo = grdlista
'    ClBuscaGrid.QueryUtilizado = sql_banco
'    Set ClBuscaGrid.RecordsetTrabajo = Adolista.Recordset
'    'ClBuscaGrid.CamposVisibles = "11010011"
'    ClBuscaGrid.Ejecutar
    
  PosibleApliqueFiltro = False
  Dim rsNada As ADODB.Recordset
  Dim GrSqlAux As String
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = grdlista
  ClBuscaGrid.QueryUtilizado = sql_banco
  Set ClBuscaGrid.RecordsetTrabajo = Adolista.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True

End Sub

Private Sub cmdCancelar_Click()
  On Error Resume Next
   Text1.Enabled = True
   fraDatos.Enabled = False
   Adolista.Recordset.Requery
   'grdlista.ReBind
   'cmdborrar.Visible = True
   Cmd_busqueda.Visible = True
   Cmdimprimir.Visible = True
   CmdSalir.Visible = True
   Cmdeditar.Visible = True
   cmdAceptar.Visible = False
   CmdAdicionar.Visible = True
   CmdCancelar.Visible = False
   Adolista.Enabled = True
   'grdlista.Enabled = True
   Adolista.Recordset.Requery
   'grdlista.ReBind
End Sub


Private Sub cmdEditar_Click()
   Adolista.Enabled = False
   'grdlista.Enabled = False
   fraDatos.Enabled = True
   
  ' cmdborrar.Visible = False
   Cmd_busqueda.Visible = False
   Cmdimprimir.Visible = False
   CmdSalir.Visible = False
   CmdAdicionar.Visible = False
   Cmdeditar.Visible = False
   cmdAceptar.Visible = True
   CmdCancelar.Visible = True
   
   Text1.Enabled = False
   Text2.Enabled = True
   Text3.Enabled = True
   Text4.Enabled = True
   Text9.Enabled = True
   DTP1.Enabled = True
   Text11.Enabled = True
   DTP2.Enabled = True
   
   Text2.SetFocus
End Sub

Private Sub Cmdimprimir_Click()
  Dim iResult As Integer
    'CrystalReport1.ReportFileName = App.Path & "\clasificadores\bancos\crybancos.rpt"
     CrystalReport1.WindowShowPrintSetupBtn = True
     CrystalReport1.WindowShowRefreshBtn = True
  CrystalReport1.ReportFileName = "\SAF-2000\Clasificadores\Tesoreria\cuentas bancarias\cryctabco.rpt"
  iResult = CrystalReport1.PrintReport
  If iResult <> 0 Then
      MsgBox CrystalReport1.LastErrorNumber & " : " & CrystalReport1.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
CrystalReport1.WindowState = crptMaximized

'RepCtaBco.Show

'   rptModalidadSeleccion.Show vbModal
End Sub


Private Sub CmdSalir_Click()
   Unload Me
End Sub


Private Sub dtcB_Click(Area As Integer)
dtcBco.BoundText = dtcB.BoundText
End Sub

Private Sub dtcBco_Click(Area As Integer)
    dtcB.BoundText = dtcBco.BoundText
End Sub
Private Sub Form_Load()
   Dim sql_moneda As String
   Dim sql_bancos As String
   Label7.Caption = frmLogin.txtUserName.Text
   Label11.Caption = Format(Time, "hh:mm:ss")
   Label27.Caption = Date
   fraDatos.Enabled = False
   CmdBorrar.Visible = True
   Cmd_busqueda.Visible = True
   Cmdimprimir.Visible = True
   CmdSalir.Visible = True
   cmdAceptar.Visible = False
   CmdCancelar.Visible = False
     
   Set rstbancos = New ADODB.Recordset
   sql_bancos = "select* from fc_bancos" ' order by bco_codigo"
   rstbancos.Open sql_bancos, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstbancos.Sort = "BCO_CODIGO"
   Set adoBancos.Recordset = rstbancos
   
   Set rstmoneda = New ADODB.Recordset
   sql_moneda = "select * from tipo_moneda" ' order by tipo_moneda"
   rstmoneda.Open sql_moneda, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstmoneda.Sort = "TIPO_MONEDA"
   Set AdoMoneda.Recordset = rstmoneda
   
   Set rstctaban = New ADODB.Recordset
   sql_banco = "select * from fc_cuenta_Bancaria" ' order by cta_codigo"
   rstctaban.Open sql_banco, db, adOpenKeyset, adLockOptimistic, adCmdText
   rstctaban.Sort = "cta_codigo"
   Set Adolista.Recordset = rstctaban
   
 
     Set ClBuscaGrid = Nothing
   
	Call SeguridadSet(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If (rstctaban.State = adStateClosed) Then rstctaban.Close
   'Set rstctaban = Nothing

End Sub


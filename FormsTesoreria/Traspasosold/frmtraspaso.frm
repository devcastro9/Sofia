VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmtraspasos 
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   6015
      Left            =   1320
      TabIndex        =   118
      Top             =   960
      Width           =   3255
      Begin Crystal.CrystalReport CryCompTraspasos 
         Left            =   2640
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   1680
         TabIndex        =   121
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptSAprobar 
         Caption         =   "Sin Aprobar"
         Height          =   255
         Left            =   240
         TabIndex        =   120
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DtGrid_comprobante 
         Height          =   5295
         Left            =   120
         TabIndex        =   119
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Cod_Comp"
            Caption         =   "Cod_Comp"
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
            DataField       =   "Tipo_Comp"
            Caption         =   "Tipo_Comp"
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
            DataField       =   "codigo_beneficiario"
            Caption         =   "codigo_beneficiario"
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
            DataField       =   "status"
            Caption         =   "status"
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
            DataField       =   "org_codigo"
            Caption         =   "org_codigo"
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
            DataField       =   "cod_trans"
            Caption         =   "cod_trans"
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
         EndProperty
      End
   End
   Begin VB.Frame frame_moneda 
      Caption         =   "Tipo de Moneda"
      Height          =   495
      Left            =   4680
      TabIndex        =   87
      Top             =   4080
      Width           =   6435
      Begin VB.OptionButton optdolares 
         Caption         =   "Dólares"
         Height          =   270
         Left            =   4200
         TabIndex        =   89
         Top             =   165
         Width           =   1590
      End
      Begin VB.OptionButton optbolivianos 
         Caption         =   "Bolivianos"
         Height          =   270
         Left            =   1680
         TabIndex        =   88
         Top             =   165
         Width           =   1350
      End
   End
   Begin VB.Frame Fra_Busqueda 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1905
      Left            =   5220
      TabIndex        =   86
      Top             =   3480
      Visible         =   0   'False
      Width           =   3876
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   3615
         Begin VB.ComboBox CboStatus 
            Height          =   288
            ItemData        =   "frmtraspaso.frx":0000
            Left            =   2520
            List            =   "frmtraspaso.frx":000A
            TabIndex        =   122
            Top             =   240
            Visible         =   0   'False
            Width           =   792
         End
         Begin VB.TextBox Text_Valor 
            Height          =   336
            Left            =   2520
            TabIndex        =   108
            Text            =   "1"
            Top             =   240
            Width           =   900
         End
         Begin VB.ComboBox Cmbo_Operador 
            Height          =   315
            ItemData        =   "frmtraspaso.frx":0014
            Left            =   1560
            List            =   "frmtraspaso.frx":0027
            TabIndex        =   107
            Text            =   "="
            Top             =   240
            Width           =   672
         End
         Begin VB.ComboBox Cmbo_Atributo 
            DataMember      =   "comprobante"
            DataSource      =   "dtetraspasos"
            Height          =   315
            ItemData        =   "frmtraspaso.frx":003C
            Left            =   120
            List            =   "frmtraspaso.frx":0052
            TabIndex        =   106
            Text            =   "Cod_Comp"
            Top             =   240
            Width           =   1284
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   101
         Top             =   960
         Width           =   3615
         Begin VB.CommandButton Cmd_Normal 
            Caption         =   "Normal"
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
            Left            =   1440
            TabIndex        =   104
            Top             =   240
            Width           =   888
         End
         Begin VB.CommandButton Cmd_BSalir 
            Caption         =   "Salir"
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
            Left            =   2520
            TabIndex        =   103
            Top             =   240
            Width           =   936
         End
         Begin VB.CommandButton cmd_Ejecutar 
            Caption         =   "Ejecutar"
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
            Left            =   120
            TabIndex        =   102
            Top             =   255
            Width           =   996
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   1335
      TabIndex        =   85
      Top             =   7065
      Width           =   2670
      Begin VB.CommandButton cmdfin 
         Height          =   555
         Left            =   1905
         Picture         =   "frmtraspaso.frx":009F
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdinicio 
         Height          =   555
         Left            =   105
         Picture         =   "frmtraspaso.frx":04E1
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdsgte 
         Height          =   555
         Left            =   1365
         Picture         =   "frmtraspaso.frx":0923
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdatras 
         DisabledPicture =   "frmtraspaso.frx":0D65
         Height          =   555
         Left            =   705
         Picture         =   "frmtraspaso.frx":11A7
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   165
         Width           =   600
      End
   End
   Begin VB.Frame Frame_aprobacion 
      Caption         =   "Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   4980
      TabIndex        =   84
      Top             =   1560
      Visible         =   0   'False
      Width           =   5235
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   180
         TabIndex        =   109
         Top             =   240
         Width           =   4935
         Begin VB.OptionButton optindividual 
            Caption         =   "Individual"
            Height          =   195
            Left            =   1080
            TabIndex        =   115
            Top             =   240
            Width           =   1050
         End
         Begin VB.OptionButton optconjunto 
            Caption         =   "Conjunto"
            Height          =   225
            Left            =   2640
            TabIndex        =   114
            Top             =   240
            Width           =   1080
         End
         Begin VB.ComboBox cboaprob_inicio 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmtraspaso.frx":15E9
            Left            =   1185
            List            =   "frmtraspaso.frx":15EB
            TabIndex        =   111
            Top             =   600
            Width           =   1125
         End
         Begin VB.ComboBox cbo_aprob_final 
            Height          =   315
            Left            =   3480
            TabIndex        =   110
            Top             =   600
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label20 
            Caption         =   "No. Comprob "
            Height          =   225
            Left            =   120
            TabIndex        =   113
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lblcomprob 
            Caption         =   "No. Comprob "
            Height          =   225
            Left            =   2400
            TabIndex        =   112
            Top             =   600
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmd_aprob_aceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   840
         TabIndex        =   117
         Top             =   1440
         Width           =   1350
      End
      Begin VB.CommandButton cmd_aprob_cancel 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3000
         TabIndex        =   116
         Top             =   1440
         Width           =   1350
      End
   End
   Begin VB.Frame FraGlobal 
      Enabled         =   0   'False
      Height          =   2955
      Left            =   4680
      TabIndex        =   32
      Top             =   1080
      Width           =   6420
      Begin VB.Frame Frame_Plan 
         Caption         =   "Plan_cuentas"
         Height          =   2655
         Left            =   1440
         TabIndex        =   38
         Top             =   3720
         Visible         =   0   'False
         Width           =   7335
         Begin VB.CommandButton Cmd_Eligir 
            Caption         =   "Elegir"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   2160
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid DtGrid_Plan 
            Height          =   1815
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3201
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
      End
      Begin VB.TextBox Txt_glosa 
         Height          =   570
         Left            =   840
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2175
         Width           =   5370
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   0
         TabIndex        =   37
         Top             =   -48
         Width           =   7110
      End
      Begin VB.TextBox Text_Tipo 
         Height          =   288
         Left            =   2685
         TabIndex        =   36
         Text            =   "Comprobante de Traspasos"
         Top             =   255
         Width           =   2430
      End
      Begin VB.TextBox Txt_Fecha 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   288
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   700
         Width           =   1095
      End
      Begin VB.TextBox Txt_ges 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00400000&
         Height          =   288
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   700
         Width           =   735
      End
      Begin VB.TextBox TxtComprobante 
         Appearance      =   0  'Flat
         DataField       =   "codigo_pago"
         DataSource      =   "AdoRegularizacion"
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
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox Txt_Respaldo 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo D1documento 
         Bindings        =   "frmtraspaso.frx":15ED
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo_Documento"
         BoundColumn     =   "Denominacion_documento"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo D2descripcion 
         Bindings        =   "frmtraspaso.frx":160B
         DataField       =   "Denominacion_documento"
         DataSource      =   "Adodcdocumento"
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Top             =   1080
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_documento"
         BoundColumn     =   "Codigo_Documento"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo d2beneficiario 
         Bindings        =   "frmtraspaso.frx":1629
         DataField       =   "denominacion_beneficiario"
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo d1beneficiario 
         Bindings        =   "frmtraspaso.frx":164A
         DataField       =   "codigo_beneficiario"
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Top             =   1800
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "denominacion_beneficiario"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   50
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label_Respaldo 
         Caption         =   "Numero de Respaldo"
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label_AntComp 
         Caption         =   "Tipo Comprobante Anterior:"
         Height          =   285
         Left            =   135
         TabIndex        =   47
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label Label_Fecha 
         Caption         =   "Fecha:"
         Height          =   285
         Left            =   4560
         TabIndex        =   46
         Top             =   735
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Glosa"
         Height          =   252
         Left            =   120
         TabIndex        =   45
         Top             =   2292
         Width           =   636
      End
      Begin VB.Label Label5 
         Caption         =   "Gestion:"
         Height          =   285
         Left            =   3120
         TabIndex        =   44
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiario:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1830
         Width           =   870
      End
      Begin VB.Label Label11 
         Caption         =   "Documento Respaldo"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1560
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Nro Comprobante:"
         Enabled         =   0   'False
         Height          =   288
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.Frame FraOpcionesDetalle 
      Height          =   6885
      Left            =   60
      TabIndex        =   31
      Top             =   990
      Width           =   1180
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":166B
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   6060
         Width           =   930
      End
      Begin VB.CommandButton cmdimprime_grid 
         Caption         =   "Imprime Grid"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":1AAD
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   5290
         Width           =   930
      End
      Begin VB.CommandButton Cmd_Busqueda 
         Caption         =   "Busqueda"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":1EEF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1330
         Width           =   945
      End
      Begin VB.CommandButton CmdAgregarDetalle 
         Caption         =   "Adicionar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":2331
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton Cmd_GrabaM 
         Caption         =   "Grabar"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":2773
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2580
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Modificar 
         Caption         =   "Modificar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":2BB5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   740
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "Cancelar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":2FF7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3340
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Aprobar 
         Caption         =   "Aprobar"
         Height          =   705
         Left            =   135
         Picture         =   "frmtraspaso.frx":30F9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   945
      End
      Begin VB.CommandButton CmdEstado 
         Caption         =   "&Estado"
         Height          =   495
         Left            =   180
         TabIndex        =   9
         Top             =   6270
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton Cmd_Copiar 
         Caption         =   "Copiar"
         Height          =   645
         Left            =   135
         Picture         =   "frmtraspaso.frx":353B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton Cmd_IMPRIMIR 
         Caption         =   "Imprimir"
         Height          =   645
         Left            =   135
         Picture         =   "frmtraspaso.frx":3A6D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4640
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   11160
      TabIndex        =   0
      Top             =   0
      Width           =   11220
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "C O M P R O B  A N  T E -  C O N T A B L E -  T R A S P A S O S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   705
         TabIndex        =   83
         Top             =   285
         Width           =   7845
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "C O M P R O B  A N  T E -  C O N T A B L E -  M A N U A L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2955
         TabIndex        =   30
         Top             =   1125
         Width           =   8415
      End
      Begin VB.Label Label7 
         Height          =   225
         Left            =   9000
         TabIndex        =   29
         Top             =   600
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7740
         TabIndex        =   28
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1020
         TabIndex        =   24
         Top             =   615
         Width           =   2460
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   75
         TabIndex        =   23
         Top             =   630
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3165
      Left            =   4680
      TabIndex        =   49
      Top             =   4680
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5583
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   420
      TabCaption(0)   =   "Crédito"
      TabPicture(0)   =   "frmtraspaso.frx":40D7
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Fram_AsientoH"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Débito"
      TabPicture(1)   =   "frmtraspaso.frx":40F3
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Fram_AsientoD"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Fram_AsientoD 
         Caption         =   "Asiento_Debito"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2685
         Left            =   240
         TabIndex        =   66
         Top             =   480
         Width           =   5955
         Begin VB.ComboBox cbod_sub2 
            Height          =   315
            ItemData        =   "frmtraspaso.frx":410F
            Left            =   3840
            List            =   "frmtraspaso.frx":411C
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.Frame FrameD_CtaCorriente 
            Caption         =   "Cuentas corrientes de Bancos"
            Height          =   1575
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   5715
            Begin VB.ComboBox cbod_aux1_denom 
               Height          =   315
               Left            =   2280
               TabIndex        =   22
               Top             =   360
               Width           =   3165
            End
            Begin VB.ComboBox cbod_aux1 
               Height          =   315
               Left            =   960
               TabIndex        =   21
               Top             =   360
               Width           =   1260
            End
            Begin MSDataListLib.DataCombo TxtD_Nom3_Corriente 
               Height          =   315
               Left            =   2280
               TabIndex        =   70
               Top             =   1080
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Nom2_Corriente 
               Height          =   315
               Left            =   2280
               TabIndex        =   71
               Top             =   720
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Aux3_Corriente 
               Height          =   315
               Left            =   960
               TabIndex        =   72
               Top             =   1080
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Aux2_Corriente 
               Height          =   315
               Left            =   960
               TabIndex        =   73
               Top             =   720
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "Descripcion:"
               Height          =   255
               Left            =   2955
               TabIndex        =   99
               Top             =   105
               Width           =   1080
            End
            Begin VB.Label Label14 
               Caption         =   "Auxiliar 1:"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label13 
               Caption         =   "Auxiliar 2:"
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "Auxiliar 3"
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   1200
               Width           =   735
            End
         End
         Begin VB.TextBox TxtD_Dls 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3210
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   615
            Width           =   1515
         End
         Begin VB.TextBox TxtD_Bs 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   975
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label lbld_cuenta 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1111"
            Height          =   270
            Left            =   720
            TabIndex        =   100
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbld_sub1 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "02"
            Height          =   270
            Left            =   2280
            TabIndex        =   82
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label_Cuenta 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   105
            TabIndex        =   81
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "MontoDls"
            Height          =   255
            Left            =   2355
            TabIndex        =   80
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label_MontoBs 
            Caption         =   "Monto_Bs"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label_Cta2 
            Caption         =   "Sub_Cta2:"
            Height          =   255
            Left            =   3000
            TabIndex        =   78
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label_Cta1 
            Caption         =   "Sub_Cta1:"
            Height          =   255
            Left            =   1440
            TabIndex        =   77
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Fram_AsientoH 
         Caption         =   "Asiento_Credito"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2715
         Left            =   -74880
         TabIndex        =   51
         Top             =   255
         Width           =   6090
         Begin VB.TextBox Txt_Cambio 
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   260
            Width           =   540
         End
         Begin VB.Frame FrameH_CtaCorriente 
            Caption         =   "Cuentas corrientes de Bancos"
            Height          =   1515
            Left            =   120
            TabIndex        =   52
            Top             =   1020
            Width           =   5865
            Begin VB.ComboBox cboH_aux1 
               Height          =   315
               Left            =   960
               TabIndex        =   94
               Top             =   360
               Width           =   1245
            End
            Begin VB.ComboBox cboh_aux1_denom 
               Height          =   315
               Left            =   2400
               TabIndex        =   93
               Top             =   360
               Width           =   3165
            End
            Begin MSDataListLib.DataCombo TxtH_Nom3_Corriente 
               Height          =   315
               Left            =   2055
               TabIndex        =   53
               Top             =   1620
               Width           =   3180
               _ExtentX        =   5609
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Nom2_Corriente 
               Height          =   315
               Left            =   2400
               TabIndex        =   95
               Top             =   720
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Aux3_Corriente 
               Height          =   315
               Left            =   960
               TabIndex        =   96
               Top             =   1080
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Aux2_Corriente 
               Height          =   315
               Left            =   960
               TabIndex        =   97
               Top             =   720
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Nom3_aux3 
               Height          =   315
               Left            =   2400
               TabIndex        =   98
               Top             =   1080
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "Descripcion:"
               Height          =   270
               Left            =   2430
               TabIndex        =   57
               Top             =   120
               Width           =   1080
            End
            Begin VB.Label Label17 
               Caption         =   "Auxiliar 3"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   1095
               Width           =   735
            End
            Begin VB.Label Label16 
               Caption         =   "Auxiliar 2:"
               Height          =   255
               Left            =   105
               TabIndex        =   55
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "Auxiliar 1:"
               Height          =   240
               Left            =   120
               TabIndex        =   54
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.TextBox Txth_dls 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3330
            TabIndex        =   19
            Top             =   675
            Width           =   1215
         End
         Begin VB.TextBox Txth_Bs 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   975
            TabIndex        =   18
            Top             =   675
            Width           =   1350
         End
         Begin VB.ComboBox cboH_sub2 
            Height          =   315
            ItemData        =   "frmtraspaso.frx":412C
            Left            =   3480
            List            =   "frmtraspaso.frx":4139
            TabIndex        =   16
            Top             =   260
            Width           =   690
         End
         Begin VB.Label Label_Dl 
            Caption         =   "MontoDls"
            Height          =   255
            Left            =   2520
            TabIndex        =   65
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label_Bs 
            Caption         =   "Monto_Bs"
            Height          =   255
            Left            =   135
            TabIndex        =   64
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label_Cambio 
            Caption         =   "Cambio_Dl:"
            Height          =   255
            Left            =   4320
            TabIndex        =   63
            Top             =   360
            Width           =   840
         End
         Begin VB.Label LabelD_Cuenta 
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   585
         End
         Begin VB.Label LabelH_Cta2 
            Caption         =   "Sub_Cta2:"
            Height          =   200
            Left            =   2640
            TabIndex        =   61
            Top             =   345
            Width           =   735
         End
         Begin VB.Label LabelH_Cuenta 
            Alignment       =   2  'Center
            Caption         =   "Sub_Cta1:"
            Height          =   210
            Left            =   1320
            TabIndex        =   60
            Top             =   345
            Width           =   735
         End
         Begin VB.Label lblh_cuenta 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1111"
            Height          =   270
            Left            =   720
            TabIndex        =   59
            Top             =   255
            Width           =   540
         End
         Begin VB.Label lblh_sub1 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "02"
            Height          =   270
            Left            =   2160
            TabIndex        =   58
            Top             =   260
            Width           =   435
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodcbeneficiario 
      Height          =   330
      Left            =   3990
      Top             =   7560
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
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
   Begin MSAdodcLib.Adodc Adodcdocumento 
      Height          =   330
      Left            =   3990
      Top             =   7335
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
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
Attribute VB_Name = "frmtraspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public num_comprobante As Integer ' vaiable donde se almacena nùmero de comprobante
Public TTipoC As Double
Public TFecha As Date
'****QUERIES
Dim sql_grid  As String
Dim rstipocambio As ADODB.Recordset
Dim sql_TC As String
'********RECORDSETS
Dim rscomprobante1 As ADODB.Recordset
Dim rsdocumento As ADODB.Recordset
Dim rsbenef_traspaso As ADODB.Recordset
Dim rscta_corriente As ADODB.Recordset
Dim rsComprobante As ADODB.Recordset
Dim rsdiario As ADODB.Recordset
Dim rsCorrelativo As ADODB.Recordset
Dim rscomprobante_M As ADODB.Recordset
Dim rscompro_N As ADODB.Recordset
Dim rsPago As ADODB.Recordset
Dim rspago_detalle As ADODB.Recordset
Dim rsRepCab As ADODB.Recordset
Dim rsRepDet As ADODB.Recordset
Dim rsPlan_cuentas As ADODB.Recordset
Dim rsnombre_cta As ADODB.Recordset
Dim rsfc_cuenta_bancaria As ADODB.Recordset
Dim rsbenef  As ADODB.Recordset
Dim rsimprgrid  As ADODB.Recordset
'*******************
Public sw1 As Integer
Public sw2 As Integer

'
'Private Sub cbobeneficiario_Change()
'  With dtetraspasos
'  If .rsbenef_traspaso.State = 1 Then .rsbenef_traspaso.Close
'  .rsbenef_traspaso.Open
'   .rsbenef_traspaso.MoveFirst
'    .rsbenef_traspaso.Find "codigo_beneficiario= '" & Trim(Me.cbobeneficiario.Text) & "'"
'   Me.cbobenef_deno.Text = .rsbenef_traspaso!denominacion_beneficiario
'  End With
'End Sub
'Private Sub cbobeneficiario_Click()
''''With dtetraspasos
''''   .rsbenef_traspaso.MoveFirst
''''   .rsbenef_traspaso.Find "codigo_beneficiario= '" & Trim(Me.cbobeneficiario.Text) & "'"
''''   Me.cbobenef_deno.Text = .rsbenef_traspaso!denominacion_beneficiario
''''  End With
'End Sub

Private Sub cbod_aux1_Change()
'On Error GoTo err1:
'With dtetraspasos
'If rscta_corriente.State = 1 Then .rscta_corriente.Close
'rscta_corriente.Open
'''''rscta_corriente.MoveFirst
'''''rscta_corriente.Filter = adFilterNone
'''''rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cbod_aux1.Text) & "'"
'''''Me.cbod_aux1_denom.Text = rscta_corriente!Cta_descripcion_larga
'End With
'Exit Sub
err1:
If Err.Number = 3021 Then
'MsgBox "Formulario con datos incompletos", vbExclamation + vbDefaultButton1, "SAF/2000"
Call limpiar
Me.DtGrid_comprobante.SetFocus
End If
End Sub
Private Sub cbod_aux1_Click()
With dtetraspasos
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cbod_aux1.Text) & "'"
Me.cbod_aux1_denom.Text = rscta_corriente!cta_descripcion_larga
End With
End Sub
Private Sub cbod_aux1_denom_Change()
'On Error GoTo err2
'With dtetraspasos
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
'b = rscta_corriente.RecordCount
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_descripcion_larga= '" & Trim(Me.cbod_aux1_denom.Text) & "'"
If rscta_corriente.RecordCount = 0 Then
    Exit Sub
Else
    Me.cbod_aux1.Text = rscta_corriente!cta_codigo
End If
'End With
'Exit Sub
err2:
If Err.Number = 3021 Then MsgBox "Comprobante con datos incompletos", vbExclamation + vbDefaultButton1, "SAF/2000"
End Sub

Private Sub cbod_aux1_denom_Click()
'''With dtetraspasos
'''If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'''.rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
'''.rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_descripcion_larga= '" & Trim(Me.cbod_aux1_denom.Text) & "'"
Me.cbod_aux1.Text = rscta_corriente!cta_codigo
'''nd With
End Sub

Private Sub cbod_aux1_LostFocus()
With dtetraspasos

  End With
End Sub

Private Sub cbod_sub2_Click()
'With dtetraspasos
  Me.cbod_aux1.Clear
  Me.cbod_aux1_denom.Clear
  'If rscta_corriente.State = 1 Then rscta_corriente.Close
  'rscta_corriente.Open
  rscta_corriente.Filter = adFilterNone
  rscta_corriente.MoveFirst
Do While Not rscta_corriente.EOF
  Select Case Me.cbod_sub2.Text
  Case "01"
    If rscta_corriente!Fte_Codigo = "41" Then
      Me.cbod_aux1.AddItem rscta_corriente!cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
    End If
  Case "02"
    If rscta_corriente!Fte_Codigo = "43" Then
      Me.cbod_aux1.AddItem rscta_corriente!cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
    End If
  Case "03"
    If rscta_corriente!Fte_Codigo = "80" Then
      Me.cbod_aux1.AddItem rscta_corriente!cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
    End If
  End Select
  rscta_corriente.MoveNext
  Loop
  Me.cbod_aux1.Text = Me.cbod_aux1.List(0)
'End With
End Sub
Private Sub cboH_aux1_Change()
'If rscta_corriente.State = 1 Then rscta_corriente.Close
'rscta_corriente.Open
'MsgBox rscta_corriente.RecordCount
rscta_corriente.Filter = adFilterNone
'rscta_corriente.MoveFirst
If rscta_corriente.RecordCount = 0 Then
    Exit Sub
Else
rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
End If
'rscta_corriente.Find "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
'MsgBox rscta_corriente.RecordCount
Me.cboh_aux1_denom.Text = rscta_corriente!cta_descripcion_larga
End Sub

Private Sub cboH_aux1_Click()
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'.rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
Me.cboh_aux1_denom.Text = rscta_corriente!cta_descripcion_larga

End Sub
Private Sub cboh_aux1_denom_Click()
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'.rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_descripcion_larga= '" & Trim(Me.cboh_aux1_denom.Text) & "'"
Me.cboH_aux1.Text = rscta_corriente!cta_codigo
End Sub

Private Sub cboH_sub2_Click()
Me.cboH_aux1.Clear
'If rscta_corriente.State = 1 Then rscta_corriente.Close
 ' rscta_corriente.Open
  rscta_corriente.Filter = adFilterNone
  rscta_corriente.MoveFirst
Do While Not rscta_corriente.EOF
Select Case Me.cboH_sub2.Text
Case "01"
  If rscta_corriente!Fte_Codigo = "41" Then
    Me.cboH_aux1.AddItem rscta_corriente!cta_codigo
    Me.cboh_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
  End If
Case "02"
  If rscta_corriente!Fte_Codigo = "43" Then
   Me.cboH_aux1.AddItem rscta_corriente!cta_codigo
   Me.cboh_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
  End If
Case "03"
  If rscta_corriente!Fte_Codigo = "80" Then
    Me.cboH_aux1.AddItem rscta_corriente!cta_codigo
    Me.cboh_aux1_denom.AddItem rscta_corriente!cta_descripcion_larga
  End If
End Select
rscta_corriente.MoveNext
Loop
Me.cboH_aux1.Text = Me.cboH_aux1.List(0)
End Sub

Private Sub Cmbo_Atributo_Change()
  If Cmbo_Atributo.Text = "status" Then
    Me.CboStatus.Visible = True
    Me.Text_Valor.Visible = False
  Else
    Me.CboStatus.Visible = False
    Me.Text_Valor.Visible = True
  End If
End Sub

Private Sub Cmbo_Atributo_LostFocus()
  If Cmbo_Atributo.Text = "status" Then
    Me.CboStatus.Visible = True
    Me.Text_Valor.Visible = False
  Else
    Me.CboStatus.Visible = False
    Me.Text_Valor.Visible = True
  End If
End Sub

Private Sub cmd_aprob_aceptar_Click()
db.BeginTrans
If sw1 = 1 Then
        '********CAMBIO DE STATUS A APROBADO
        Set rscomprobante_M = New ADODB.Recordset
        If rscomprobante_M.State = 1 Then rscomprobante_M.Close
          rscomprobante_M.Open "select * from Co_Comprobante_M where cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockOptimistic
          rscomprobante_M.MoveFirst
     'MsgBox rscomprobante_M!Cod_Comp & rscomprobante_M!Tipo_Comp

     If rscomprobante_M!Status = "N" Then
              rscomprobante_M!Status = "P"
              rscomprobante_M!Fecha_A = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
              rscomprobante_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
              rscomprobante_M.Update
            Set rsdiario = New ADODB.Recordset
            If rsdiario.State = 1 Then rsdiario.Close
              rsdiario.Open "SELECT * FROM CO_Diario " & _
              "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockOptimistic
              rsdiario!Cod_Comp_C = Val(Trim(Me.cboaprob_inicio.Text))
              rsdiario.Update
            Set rsPago = New ADODB.Recordset
            Set rspago_detalle = New ADODB.Recordset
          If rsPago.State = 1 Then rsPago.Close
            rsPago.Open "SELECT * FROM pagos WHERE (ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
          
          '.Connection1.BeginTrans
          '*********ADICION A LA TABLA PAGO
            rsPago.AddNew
            rsPago!Ges_gestion = Trim(rscomprobante_M!Ges_gestion)
            rsPago!org_codigo = "999"
            rsPago!codigo_pago = Trim(rscomprobante_M!Cod_Comp)
            '.rspago!nro_comprobante_anterior = .rscomprobante!Cod_Comp
            rsPago!tipo_comp = "TRP"
            rsPago!Codigo_orden = Trim(rscomprobante_M!Num_respaldo)
            rsPago!Codigo_documento = Trim(rscomprobante_M!Codigo_documento)
            rsPago!fecha_egreso = CDate(Format(rscomprobante_M!Fecha_A, "dd/mm/yyyy"))
            rsPago!monto_bolivianos = rsdiario!D_MontoBs
            rsPago!monto_Dolares = rsdiario!D_MontoDl
            rsPago!liquido_pagar = rsdiario!D_MontoBs
            rsPago!estado_aprobacion = "N"
            rsPago!estado_contabilidad = "P"
            rsPago!Justificacion = Trim(rscomprobante_M!Glosa)
            rsPago!estado_pagado = "S"
            rsPago!Usr_Usuario = Trim(rscomprobante_M!Usr_Usuario)
            rsPago!fecha_aprueba = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
            rsPago!hora_aprueba = (Format(Time, "hh:mm:ss"))
            rsPago!fecha_registro = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
            rsPago!Hora_Registro = (Format(Time, "hh:mm:ss"))
            '********ADICION A LA TABLA PAGO DETALLE
            If rspago_detalle.State = 1 Then rspago_detalle.Close
            rspago_detalle.Open "SELECT * FROM pago_detalle WHERE (Ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
            'MsgBox rspago_detalle.RecordCount
            rspago_detalle.AddNew
            rspago_detalle!Ges_gestion = Trim(rscomprobante_M!Ges_gestion)
            rspago_detalle!org_codigo = "999"
            rspago_detalle!codigo_pago = Trim(Str(rscomprobante_M!Cod_Comp))
            rspago_detalle!codigo_Pago_detalle = "1"
            rspago_detalle!Codigo_Beneficiario = Trim(rscomprobante_M!Codigo_Beneficiario)
            rspago_detalle!tipo_cambio = rsdiario!D_Cambio
            rspago_detalle!monto_total = rsdiario!D_MontoBs
            rspago_detalle!departamento = "La Paz"
            rspago_detalle!honorarios = "N"
            ''''''''''''
            rspago_detalle!cta_codigo_destino = Trim(rsdiario!d_cta_Larga)
            rscta_corriente.Filter = "cta_codigo='" & rsdiario!d_cta_Larga & "'"
            rspago_detalle!banco_destino = Trim(rscta_corriente!Bco_descripcion_larga)
            rscta_corriente.Filter = adFilterNone
            rspago_detalle!cta_codigo = Trim(rsdiario!h_cta_Larga)
            rspago_detalle!cheque_o_trf = "R"   'prefijo para traspasos
            rspago_detalle!tipo_cambio = rsdiario!D_Cambio
            rspago_detalle!estado_aprobacion = "N"
            rspago_detalle!monto_bolivianos = rsdiario!D_MontoBs
            rspago_detalle!monto_Dolares = rsdiario!D_MontoDl
            rspago_detalle!fecha_pago = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
            'rspago_detalle!departamento=
            'rspago_detalle!beneficiario_destino=
            
            rspago_detalle!Usr_Usuario = Trim(rscomprobante_M!Usr_Usuario)
            rspago_detalle!fecha_registro = Format(Me.TFecha, "dd/mm/yyyy")
            rspago_detalle!Hora_Registro = Format(Time, "hh:mm:ss")
           '********ACTUALIZACION MOVIMIENTOS DEBE Y HABER EN LA CUENTA BANCARIA
            Set rsfc_cuenta_bancaria = New ADODB.Recordset
            'CTA BANCARIA DEL DEBE, LA QUE RECIBE
            If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
            rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!d_cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
            rsfc_cuenta_bancaria.MoveFirst
            rsfc_cuenta_bancaria!Cta_Saldo_Debe = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Debe), 0, rsfc_cuenta_bancaria!Cta_Saldo_Debe) + rsdiario!D_MontoBs
            rsfc_cuenta_bancaria.Update
            ' CTA BANCARIA DEL HABER, LA QUE DA
            If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
            rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!h_cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
            rsfc_cuenta_bancaria!Cta_Saldo_Haber = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Haber), 0, rsfc_cuenta_bancaria!Cta_Saldo_Haber) + Val(rsdiario!h_MontoBs)
            rsfc_cuenta_bancaria.Update
            rsPago.Update
            rspago_detalle.Update

          '.Connection1.CommitTrans
          
      MsgBox "Comprobante aprobado", vbInformation + vbDefaultButton1, "SAS/2000"
    
        'If .rscomprobante.State = 1 Then .rscomprobante.Close
        'rscomprobante.Filter = "status ='N'"
        '.rscomprobante.Open
        
     Else '*******estado comprobante
        MsgBox "El comprobante " & Trim(Me.cboaprob_inicio) & " ya está aprobado", vbExclamation + vbDefaultButton1
        Me.cboaprob_inicio.SetFocus
        Exit Sub
    End If

Else '***del swich
  If sw1 = 0 And (Me.cboaprob_inicio.Text < Me.cbo_aprob_final.Text) Then
  'abrimos diario y comprobante_M entre rangos
        Set rscomprobante_M = New ADODB.Recordset
        If rscomprobante_M.State = 1 Then rscomprobante_M.Close
        rscomprobante_M.Open " Select * from co_comprobante_M where cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text), db, adOpenKeyset, adLockOptimistic
        Set rsdiario = New ADODB.Recordset
        
        If rsdiario.State = 1 Then rsdiario.Close
        rsdiario.Open " select * from Co_Diario where cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text), db, adOpenKeyset, adLockOptimistic
        
        For i = Val(Trim(Me.cboaprob_inicio)) To Val(Trim(Me.cbo_aprob_final))
            rscomprobante_M.Filter = adFilterNone
            rscomprobante_M.Filter = "cod_comp=" & i
            'MsgBox rscomprobante_M.RecordCount + rscomprobante_M!Cod_Comp
          '********CAMBIO DE STATUS A APROBADO
            'rscomprobante_M.MoveFirst

            If rscomprobante_M!Status = "N" Then
                    rscomprobante_M!Status = "P"
                    rscomprobante_M!Fecha_A = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
                    rscomprobante_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
                    rscomprobante_M.Update
                    rsdiario.MoveFirst
                    rsdiario.Filter = adFilterNone
                    rsdiario.Filter = "cod_comp=" & i
                    'rsdiario.Find "cod_comp=" & i
                    
                    If rsPago.State = 1 Then rsPago.Close
                    rsPago.Open "SELECT * FROM pagos WHERE (ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
                    If rspago_detalle.State = 1 Then rspago_detalle.Close
                    rspago_detalle.Open "SELECT * FROM pago_detalle WHERE (Ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
                  '*********ADICION A LA TABLA PAGO
                    rsPago.AddNew
                    rsPago!Ges_gestion = rscomprobante_M!Ges_gestion
                    rsPago!org_codigo = "999"
                    rsPago!codigo_pago = rscomprobante_M!Cod_Comp
                    '.rspago!nro_comprobante_anterior = .rscomprobante!Cod_Comp
                    rsPago!tipo_comp = "PCE"
                    rsPago!Codigo_orden = rscomprobante_M!Num_respaldo
                    rsPago!Codigo_documento = rscomprobante_M!Codigo_documento
                    rsPago!fecha_egreso = CDate(Format(rscomprobante_M!Fecha_A, "dd/mm/yyyy"))
                    rsPago!monto_bolivianos = rsdiario!D_MontoBs
                    rsPago!monto_Dolares = rsdiario!D_MontoDl
                    rsPago!liquido_pagar = rsdiario!D_MontoBs
                    'celia rspago!estado_aprobacion = "N" o "A"
                    rsPago!estado_contabilidad = "P"
                    rsPago!Justificacion = rscomprobante_M!Glosa
                    rsPago!estado_pagado = "S"
                    rsPago!Usr_Usuario = rscomprobante_M!Usr_Usuario
                    rsPago!fecha_aprueba = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
                    rsPago!hora_aprueba = (Format(Time, "hh:mm:ss"))
                    rsPago!fecha_registro = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
                    rsPago!Hora_Registro = (Format(Time, "hh:mm:ss"))
                    '********ADICION A LA TABLA PAGO DETALLE
                    rspago_detalle.AddNew
                    rspago_detalle!Ges_gestion = rscomprobante_M!Ges_gestion
                    rspago_detalle!org_codigo = "999"
                    rspago_detalle!codigo_pago = Str(rscomprobante_M!Cod_Comp)
                    rspago_detalle!codigo_Pago_detalle = "1"
                    rspago_detalle!Codigo_Beneficiario = rscomprobante_M!Codigo_Beneficiario
                    rspago_detalle!tipo_cambio = rsdiario!D_Cambio
                    rspago_detalle!monto_total = rsdiario!D_MontoBs
                    rspago_detalle!departamento = "La Paz"
                    rspago_detalle!honorarios = "N"
                    ''''''''''''
                    rspago_detalle!cta_codigo_destino = rsdiario!d_cta_Larga
                    rscta_corriente.Filter = "cta_codigo='" & rsdiario!d_cta_Larga & "'"
                    rspago_detalle!banco_destino = rscta_corriente!Bco_descripcion_larga
                    rscta_corriente.Filter = adFilterNone
                    rspago_detalle!cta_codigo = rsdiario!h_cta_Larga
                    rspago_detalle!cheque_o_trf = "R"   ' prefijo para traspasos
                    rspago_detalle!tipo_cambio = rsdiario!D_Cambio
                    rspago_detalle!estado_aprobacion = "N"
                    rspago_detalle!monto_bolivianos = rsdiario!D_MontoBs
                    rspago_detalle!monto_Dolares = rsdiario!D_MontoDl
                    rspago_detalle!fecha_pago = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
                    'rspago_detalle!departamento=
                    'rspago_detalle!beneficiario_destino=
                    
                    rspago_detalle!Usr_Usuario = rscomprobante_M!Usr_Usuario
                    rspago_detalle!fecha_registro = Format(Me.TFecha, "dd/mm/yyyy")
                    rspago_detalle!Hora_Registro = Format(Time, "hh:mm:ss")
                   '********ACTUALIZACION MOVIMIENTOS DEBE Y HABER EN LA CUENTA BANCARIA
                    Set rsfc_cuenta_bancaria = New ADODB.Recordset
                    'CTA BANCARIA DEL DEBE, LA QUE RECIBE
                    If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
                    rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!d_cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
                    rsfc_cuenta_bancaria.MoveFirst
                    rsfc_cuenta_bancaria!Cta_Saldo_Debe = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Debe), 0, rsfc_cuenta_bancaria!Cta_Saldo_Debe) + rsdiario!D_MontoBs
                    rsfc_cuenta_bancaria.Update
                    ' CTA BANCARIA DEL HABER, LA QUE DA
                    If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
                    rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!h_cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
                    rsfc_cuenta_bancaria!Cta_Saldo_Haber = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Haber), 0, rsfc_cuenta_bancaria!Cta_Saldo_Haber) + Val(rsdiario!h_MontoBs)
                    rsfc_cuenta_bancaria.Update
                    rsPago.Update
                    rspago_detalle.Update

            Else '******* si esta aprobado
               MsgBox " El comprobante " & i & "ya está aprobado", vbExclamation + vbDefaultButton1
            End If
    Next
  
  MsgBox "Comprobantes aprobados", vbInformation + vbDefaultButton1, "SAF/2000"
Else
  MsgBox "Introduzca nuevamente el rango", vbCritical + vbDefaultButton1, "SAF/2000"
End If

End If ' del sw
db.CommitTrans
'        Me.Frame_aprobacion.Visible = False
'        Me.FraGlobal.Visible = True
'        Me.Frame_aprobacion.Visible = True
'        Me.Cmd_Aprobar.Enabled = True
'        Me.Cmd_Busqueda.Enabled = True
'        Me.Cmd_Cancelar.Enabled = True
'        Me.Cmd_Copiar.Enabled = True
'        Me.Cmd_Eligir.Enabled = True
'        Me.Cmd_GrabaM.Enabled = True
'        Me.Cmd_IMPRIMIR.Enabled = True
'        Me.Cmd_Modificar.Enabled = True
'        Me.CmdAgregarDetalle.Enabled = True
'        Me.CmdEstado.Enabled = True
'        Me.CmdSalir.Enabled = True
        Me.cbo_aprob_final.Clear
        Me.cboaprob_inicio.Clear
        rsComprobante.Requery
        'MsgBox rscomprobante.RecordCount
        rsComprobante.Filter = adFilterNone
        'MsgBox rscomprobante.RecordCount
        rsComprobante.Filter = "status='N'"
        'MsgBox rscomprobante.RecordCount
        If rsComprobante.RecordCount <> 0 Then
          Do While Not rsComprobante.EOF
            Me.cboaprob_inicio.AddItem rsComprobante!Cod_Comp
            Me.cbo_aprob_final.AddItem rsComprobante!Cod_Comp
            rsComprobante.MoveNext
          Loop
        End If
          'rscomprobante.Filter = adFilterNone
        Set Me.DtGrid_comprobante.DataSource = rsComprobante

End Sub

Private Sub cmd_aprob_cancel_Click()
    Me.cboaprob_inicio.Clear
    Me.Frame_aprobacion.Visible = False
    Me.FraGlobal.Visible = True
    Me.FraGlobal.Visible = True
    Me.Frame_aprobacion.Visible = True
    Me.Cmd_Aprobar.Enabled = True
    Me.Cmd_Busqueda.Enabled = True
    Me.Cmd_Cancelar.Enabled = True
    Me.Cmd_Copiar.Enabled = True
    Me.Cmd_Eligir.Enabled = True
    Me.Cmd_GrabaM.Enabled = True
    Me.Cmd_IMPRIMIR.Enabled = True
    Me.Cmd_Modificar.Enabled = True
    Me.CmdAgregarDetalle.Enabled = True
    Me.CmdEstado.Enabled = True
    Me.CmdSalir.Enabled = True
    Me.Frame_aprobacion.Visible = False
    rsComprobante.Requery
    rsComprobante.Filter = adFilterNone
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
End Sub

Private Sub Cmd_Aprobar_Click()
'Me.Cmbo_Atributo = Clear
'With dtetraspasos
'If .rscomprobante.State = 1 Then .rscomprobante.Close
Me.cbo_aprob_final.Clear
Me.cboaprob_inicio.Clear
rsComprobante.Filter = adFilterNone
rsComprobante.Filter = "status ='N'"
'.rscomprobante.Open
If rsComprobante.RecordCount <> 0 Then
Do While Not rsComprobante.EOF
Me.cboaprob_inicio.AddItem rsComprobante!Cod_Comp
Me.cbo_aprob_final.AddItem rsComprobante!Cod_Comp
rsComprobante.MoveNext
Loop
'***
'MsgBox rscomprobante.RecordCount
Set Me.DtGrid_comprobante.DataSource = rsComprobante
Me.DtGrid_comprobante.Refresh
Me.FraGlobal.Visible = True
Me.Frame_aprobacion.Visible = True
Me.Cmd_Aprobar.Enabled = False
Me.Cmd_Busqueda.Enabled = False
Me.Cmd_Cancelar.Enabled = False
Me.Cmd_Copiar.Enabled = False
Me.Cmd_Eligir.Enabled = False
Me.Cmd_GrabaM.Enabled = False
Me.Cmd_IMPRIMIR.Enabled = False
Me.Cmd_Modificar.Enabled = False
Me.CmdAgregarDetalle.Enabled = False
Me.CmdEstado.Enabled = False
Me.CmdSalir.Enabled = False
'rscomprobante.Filter = adFilterNone
Else
MsgBox "No existen comprobantes para aprobar", vbExclamation + vbDefaultButton1

End If
'End With

End Sub

Private Sub Cmd_BSalir_Click()
With dtetraspasos
   'rscomprobante.Filter = adFilterNone
   'MsgBox .rscomprobante.RecordCount
   Set Me.DtGrid_comprobante.DataSource = rsComprobante
   Me.DtGrid_comprobante.Refresh
   'Me.Fram_AsientoD.Enabled = True
   'Me.Fram_AsientoH.Enabled = True
   'Me.Cmd_Busqueda.Enabled = False
'   Me.Cmd_Copiar.Enabled = False
'   Me.Cmd_GenAciento.Enabled = False
'   Me.Cmd_GrabaM.Enabled = True
'   Me.Cmd_IMPRIMIR.Enabled = True
'   Me.CmdAgregarDetalle.Enabled = True
'   Me.CmdSalir.Enabled = True
Me.Cmd_Copiar.Enabled = True
   Me.Cmd_Modificar.Enabled = True
   Me.Cmd_Busqueda.Enabled = True
   Me.Cmd_Copiar.Enabled = True
   Me.Cmd_Aprobar.Enabled = True
    Me.CmdEstado.Enabled = True
   Me.Cmd_GrabaM.Enabled = True
   Me.Cmd_IMPRIMIR.Enabled = False
   Me.CmdAgregarDetalle.Enabled = True
   Me.CmdSalir.Enabled = True

Me.FraGlobal.Visible = False
Me.SSTab1.Visible = False
Me.Fra_Busqueda.Visible = True
Me.FraGlobal.Visible = True
Me.Fra_Busqueda.Visible = False
Me.SSTab1.Visible = True

End With
End Sub

Private Sub Cmd_Cancelar_Click()
'With dtetraspasos
Me.FraGlobal.Enabled = False
Me.Fram_AsientoD.Enabled = False
Me.Fram_AsientoH.Enabled = False
   rsComprobante.Filter = adFilterNone
Set Me.DtGrid_comprobante.DataSource = rsComprobante
  Me.DtGrid_comprobante.Refresh
'End With
  Call limpiar
  Me.Cmd_GrabaM.Enabled = False
  Me.CmdSalir.Enabled = True
  Me.Cmd_Modificar.Enabled = True
  Me.CmdAgregarDetalle.Enabled = True
  Me.Cmd_Aprobar.Enabled = True
  Me.Cmd_Busqueda.Enabled = True
  Me.Cmd_Copiar.Enabled = True
  Me.Cmd_Eligir.Enabled = True
  Me.Cmd_IMPRIMIR.Enabled = True
  Me.DtGrid_comprobante.Enabled = True
  Me.frame_moneda.Visible = False
  'Me.FraGlobal.Enabled = True
  'Me.Fram_AsientoD.Enabled = True
  'Me.Fram_AsientoH.Enabled = True

End Sub

Private Sub Cmd_Copiar_Click()
sw2 = 1
'With dtetraspasos
   rsComprobante.Filter = adFilterNone
   'MsgBox .rscomprobante.RecordCount
   Set Me.DtGrid_comprobante.DataSource = rsComprobante
   Me.Fram_AsientoD.Enabled = True
   Me.Fram_AsientoH.Enabled = True
   Me.Cmd_Busqueda.Enabled = False
   Me.Cmd_Copiar.Enabled = False
   'Me.Cmd_GenAciento.Enabled = False
   Me.Cmd_GrabaM.Enabled = True
   Me.Cmd_IMPRIMIR.Enabled = True
   Me.CmdAgregarDetalle.Enabled = True
   Me.CmdSalir.Enabled = True
   Me.frame_moneda.Visible = True
   Me.TxtComprobante = ""
'End With
End Sub
Private Sub cmd_Ejecutar_Click()
'MsgBox Me.Cmbo_Atributo.List(0)
'With dtetraspasos
'rsComprobante.Requery
'rsComprobante.Resync
'rsComprobante.MoveFirst

rsComprobante.Filter = adFilterNone
'rsComprobante.RecordCount
Select Case Cmbo_Atributo.Text
 Case "Cod_Comp"
  '  If rscomprobante.State = 1 Then rscomprobante.Close
    Select Case Me.Cmbo_Operador.Text
    Case "="
      rsComprobante.Filter = "cod_comp =" & Val(Me.Text_Valor)
    Case ">"
        rsComprobante.Filter = "cod_comp >" & Val(Me.Text_Valor)
    '  .rscomprobante.Open
    '  Set Me.DtGrid_comprobante.DataSource = .rscomprobante
    Case "<"
      rsComprobante.Filter = "cod_comp <" & Val(Me.Text_Valor)
    Case "<="
        rsComprobante.Filter = "cod_comp <=" & Val(Me.Text_Valor)
    Case ">="
        rsComprobante.Filter = "cod_comp >=" & Val(Me.Text_Valor)
    End Select
    'rscomprobante.Open
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
      
 Case "Codigo_Beneficiario"
 'If rscomprobante.State = 1 Then rscomprobante.Close
    Select Case Me.Cmbo_Operador.Text
    Case "="
      rsComprobante.Filter = "codigo_beneficiario=" & Trim(Me.Text_Valor)
    Case ">", "<", "<=", ">="
      rsComprobante.Filter = "codigo_beneficiario >" & Trim(Me.Text_Valor)
    End Select
    'rsComprobante.Open
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
      
 Case "cod_trans"
 'If rscomprobante.State = 1 Then rscomprobante.Close
    Select Case Me.Cmbo_Operador.Text
    Case "="
      rsComprobante.Filter = "cod_trans =" & Val(Me.Text_Valor)
    Case ">"
        rsComprobante.Filter = "cod_trans  >" & Val(Me.Text_Valor)
    '  .rscomprobante.Open
    '  Set Me.DtGrid_comprobante.DataSource = .rscomprobante
    Case "<"
      rsComprobante.Filter = "cod_trans  <" & Val(Me.Text_Valor)
    Case "<="
       rsComprobante.Filter = "cod_trans  <=" & Val(Me.Text_Valor)
    Case ">="
       rsComprobante.Filter = "cod_trans  >=" & Val(Me.Text_Valor)
    End Select
    rsComprobante.Open
        Set Me.DtGrid_comprobante.DataSource = rsComprobante
 Case "cod_comp"
     Select Case Me.Cmbo_Operador.Text
     Case "="
        rsComprobante.Filter = "cod_comp='" & Trim(Me.Text_Valor) & "'"
     Case Else
        rsComprobante.Filter = "cod_comp='" & Trim(Me.Text_Valor) & "'"
     End Select
 Case "org_codigo"
     Select Case Me.Cmbo_Operador.Text
     Case "="
        rsComprobante.Filter = "org_codigo='" & Trim(Me.Text_Valor) & "'"
     Case Else
        rsComprobante.Filter = "org_codigo='" & Trim(Me.Text_Valor) & "'"
     End Select
 Case "tipo_comp"
     Select Case Me.Cmbo_Operador.Text
     Case "="
        rsComprobante.Filter = "tipo_comp='" & Trim(Me.Text_Valor) & "'"
     Case Else
        rsComprobante.Filter = "tipo_comp='" & Trim(Me.Text_Valor) & "'"
     End Select
 Case "status"
     Select Case Me.Cmbo_Operador.Text
     Case "="
        rsComprobante.Filter = "status='" & Trim(Me.CboStatus.Text) & "'"
     Case Else
        rsComprobante.Filter = "status='" & Trim(Me.CboStatus.Text) & "'"
     End Select
 End Select
If rsComprobante.RecordCount = 0 Then
  MsgBox "No existe ese registro", vbExclamation, "SAF/2000"
  'Me.FraGlobal.Visible = True
  rsComprobante.Filter = adFilterNone
  Set Me.DtGrid_comprobante.DataSource = rsComprobante
  Me.DtGrid_comprobante.Refresh
  Me.Cmd_Copiar.Enabled = True
  Me.Cmd_Modificar.Enabled = True
  Me.Cmd_Busqueda.Enabled = True
  Me.Cmd_Copiar.Enabled = True
  Me.Cmd_Aprobar.Enabled = True
  Me.CmdEstado.Enabled = True
  Me.Cmd_GrabaM.Enabled = False
  Me.Cmd_IMPRIMIR.Enabled = False
  Me.CmdAgregarDetalle.Enabled = True
  Me.CmdSalir.Enabled = True
  Me.FraGlobal.Visible = False
  Me.SSTab1.Visible = False
  Me.Fra_Busqueda.Visible = True
  Me.FraGlobal.Visible = True
  Me.Fra_Busqueda.Visible = False
  Me.SSTab1.Visible = True
End If
Me.Txt_Cambio.Text = Str(1)
Me.Fra_Busqueda.Visible = False
'If rsComprobante!Status = "S" Then Me.Cmd_Modificar.Enabled = False
'End With

End Sub

Private Sub Cmd_GrabaM_Click()
'On Error GoTo err3
  If Val(Me.Txth_Bs.Text) <= 0 Or Val(Me.Txt_Cambio) <= 0 Then
      MsgBox "Complete los montos", vbCritical + vbDefaultButton1
      Me.Txt_Cambio.SetFocus
      Exit Sub
  End If
  
  rscta_corriente.MoveFirst
  rscta_corriente.Filter = adFilterNone
  rscta_corriente.Filter = "Cta_codigo = '" & Trim(Me.cboH_aux1) & "'"
  'MsgBox rscta_corriente.RecordCount
  If Val(Me.Txth_Bs) > rscta_corriente!Cta_saldo_actual Then
    MsgBox "El saldo en la cuenta es : " & Str(Round(rscta_corriente!Cta_saldo_actual, 2)), vbCritical + vbDefaultButton1
    Me.Txth_Bs.SetFocus
  End If
  Set rscomprobante_M = New ADODB.Recordset
  If rscomprobante_M.State = 1 Then rscomprobante_M.Close
  rscomprobante_M.Open "SELECT * FROM Co_Comprobante_M " & _
  "where co_COmprobante_M.Ges_gestion='9999'", db, adOpenKeyset, adLockOptimistic
  'MsgBox rscomprobante_M.RecordCount
  Set rsdiario = New ADODB.Recordset
  If rsdiario.State = 1 Then rsdiario.Close
  rsdiario.Open "SELECT * FROM CO_Diario " & _
  "WHERE (H_Cuenta = '9999')", db, adOpenKeyset, adLockOptimistic
  db.BeginTrans 'inicio de la transaccion
  
  Set rsnombre_cta = New ADODB.Recordset
    '****ADICION ALCOMPROBANTE_M
    rscomprobante_M.AddNew
    Call genera_codigo
    rscomprobante_M!Cod_Comp = Trim(Str(num_comprobante))
    rscomprobante_M!tipo_comp = "TRP"
    rscomprobante_M!Cod_Trans = 0
    rscomprobante_M!Cod_Trans_Detalle = "1"
    rscomprobante_M!org_codigo = "999"
    rscomprobante_M!Ges_gestion = Trim(Me.Txt_ges)
    rscomprobante_M!Num_respaldo = Trim(Me.Txt_Respaldo)
    rscomprobante_M!Fecha_A = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
    rscomprobante_M!Codigo_Beneficiario = Trim(Me.d1beneficiario.Text)
    rscomprobante_M!Codigo_documento = Trim(Me.D1documento.Text)
    rscomprobante_M!Glosa = Trim(Me.Txt_glosa)
    rscomprobante_M!Status = "N"
    '****'''''revisar codigo de usuario
    rscomprobante_M!Usr_Usuario = GlUsuario ' variable global de usuario glusuario
    rscomprobante_M!fecha_registro = CDate(Format(Me.TFecha, "mm/dd/yyyy"))
    rscomprobante_M!Hora_Registro = Format(Time, "hh:mm:ss")
    '********ADICION AL DIARIO
    rsdiario.AddNew
    rsdiario!tipo_comp = "TRP"
    rsdiario!Cod_Comp_C = 0
    rsdiario!D_Cuenta = "1111"
    If rsnombre_cta.State = 1 Then rsnombre_cta.Close
    rsnombre_cta.Open "SELECT NombreCta  From CC_Plan_Cuentas " & _
    "WHERE Cuenta = '" & Trim(Me.lbld_cuenta) & "' AND SubCta1 = '" & Trim(Me.lbld_sub1) & _
    "' AND  SubCta2 = '" & Trim(Me.cbod_sub2.Text) & "' AND (Aux1 = '02') " & _
    "AND (Aux2 = '00') AND (Aux3 = '00')", db, adOpenKeyset, adLockReadOnly
    If rsnombre_cta.RecordCount <> 0 Then
      rsdiario!D_Nombre = rsnombre_cta!NombreCta
    Else
      rsdiario!D_Nombre = ""
    End If
    rsdiario!D_SubCta1 = "02"
    rsdiario!D_SubCta2 = Trim(Me.cbod_sub2.Text)
    rsdiario!d_Aux1 = "02"
    rsdiario!d_Aux2 = "00"
    rsdiario!d_Aux3 = "00"
    rsdiario!d_cta_Larga = Trim(Me.cbod_aux1.Text)
    rsdiario!d_des_Larga = Trim(Me.cbod_aux1_denom.Text)
    rsdiario!D_MontoBs = Val(Me.TxtD_Bs)
    rsdiario!D_MontoDl = Val(Me.TxtD_Dls)
    rsdiario!D_Cambio = Val(Me.Txt_Cambio)
  '*************
    rsdiario!H_Cuenta = "1111"
    If rsnombre_cta.State = 1 Then rsnombre_cta.Close
    If rsnombre_cta.State = 1 Then rsnombre_cta.Close
     rsnombre_cta.Open "SELECT NombreCta  From CC_Plan_Cuentas " & _
    "WHERE Cuenta = '" & Trim(Me.lblh_cuenta) & "' AND SubCta1 = '" & Trim(Me.lblh_sub1) & _
    "' AND  SubCta2 = '" & Trim(Me.cboH_sub2.Text) & "' AND (Aux1 = '02') " & _
    "AND (Aux2 = '00') AND (Aux3 = '00')", db, adOpenKeyset, adLockReadOnly
    
'    rsnombre_cta.Open " SELECT CC_Plan_Cuentas.NombreCta" & _
'    "From CC_Plan_Cuentas where CC_Plan_Cuentas.Cuenta = '" & Trim(Me.lblh_cuenta) & _
'    "' AND CC_Plan_Cuentas.SubCta1= '" & Trim(Me.lblh_sub1) & "' AND CC_Plan_Cuentas.SubCta2= '" & Trim(Me.cboH_sub2) & _
'    "' AND CC_Plan_Cuentas.Aux1 = '02' AND CC_Plan_Cuentas.Aux2 = '00'" & _
'    " AND CC_Plan_Cuentas.Aux3= '00'", db, adOpenKeyset, adLockOptimistic
    If rsnombre_cta.RecordCount <> 0 Then
      rsdiario!H_Nombre = Trim(rsnombre_cta!NombreCta)
    Else
      rsdiario!H_Nombre = ""
    End If
    rsdiario!H_SubCta1 = "02"
    rsdiario!H_SubCta2 = Trim(Me.cboH_sub2.Text)
    rsdiario!h_Aux1 = "02"
    rsdiario!h_Aux2 = "00"
    rsdiario!h_Aux3 = "00"
    rsdiario!h_cta_Larga = Trim(Me.cboH_aux1.Text)
    rsdiario!h_des_Larga = Trim(Me.cboh_aux1_denom.Text)
    rsdiario!h_MontoBs = Val(Me.Txth_Bs)
    rsdiario!h_MontoDl = Val(Me.Txth_dls)
    rsdiario!h_Cambio = Val(Me.Txt_Cambio)
    '''''revisar codigo de usuario
    rsdiario!Usr_Usuario = GlUsuario ' variable global de usuario
    rsdiario!fecha_registro = CDate(Format(Me.TFecha, "dd/mm/yyyy"))
    rsdiario!Hora_Registro = Format(Time, "hh:mm:ss")
    rsdiario!Cod_Comp = Trim(Str(num_comprobante))
    rsdiario.Update
    rscomprobante_M.Update
    db.CommitTrans
    MsgBox "Registro el comprobante " & num_comprobante & " TRP", vbInformation + vbDefaultButton1, "SAF/2000"
    Me.TxtComprobante = num_comprobante
    
    rsComprobante.Requery
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
    Me.DtGrid_comprobante.Refresh
    Me.Cmd_GrabaM.Enabled = False
    Me.CmdSalir.Enabled = True
    Me.Cmd_Modificar.Enabled = True
    Me.CmdAgregarDetalle.Enabled = True
    Me.Cmd_Aprobar.Enabled = True
    Me.Cmd_Busqueda.Enabled = True
    Me.Cmd_Copiar.Enabled = True
    Me.Cmd_Eligir.Enabled = True
    Me.Cmd_IMPRIMIR.Enabled = True
    Me.FraGlobal.Enabled = False
    Me.DtGrid_comprobante.Enabled = True
    Me.frame_moneda.Visible = False
    Me.Fram_AsientoD.Enabled = False
    Me.Fram_AsientoH.Enabled = False
    Me.SSTab1.Tab = 0
Exit Sub
err3:
db.RollbackTrans
MsgBox "Error al actualizar los datos"
End Sub

Private Sub Cmd_IMPRIMIR_Click()
    Dim literales As String
    Dim literalCry As String
    Dim decimal2 As String
    Dim RECSETAUX As ADODB.Recordset
    Set RECSETAUX = New ADODB.Recordset
    If RECSETAUX.State = 1 Then RECSETAUX.Close
    RECSETAUX.Open " SELECT  distinct Co_Comprobante_M.Cod_Comp,Co_Comprobante_M.Tipo_Comp,CO_Diario.D_MontoBs " & _
        " From CO_Diario,CO_Comprobante_M WHERE CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp AND " & _
        " CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp and CO_Diario.Tipo_Comp='" & rsComprobante!tipo_comp & _
        "' and CO_Diario.Cod_Comp=" & Val(rsComprobante!Cod_Comp), db, adOpenDynamic, adLockOptimistic
    If RECSETAUX.RecordCount <> 0 Then
        literalCry = Str(Int(RECSETAUX!D_MontoBs))
        decimal2 = Str(Round((RECSETAUX!D_MontoBs - Val(literalCry)), 2) * 100)
        literales = Literal(literalCry) + " " + decimal2 + "/100  Bolivianos"
        Dim IResult As Integer
        CryCompTraspasos.Destination = crptToWindow
        CryCompTraspasos.ReportFileName = App.Path & "\FormsContabilidad\reportes\CryComprob_Conta.rpt"
        CryCompTraspasos.StoredProcParam(0) = RECSETAUX!Cod_Comp
        CryCompTraspasos.StoredProcParam(1) = RECSETAUX!tipo_comp
        CryCompTraspasos.StoredProcParam(2) = "gaby"
        CryCompTraspasos.StoredProcParam(3) = literales
        IResult = CryCompTraspasos.PrintReport
        If IResult <> 0 Then
               MsgBox CryCompTraspasos.LastErrorNumber & " : " & CryCompTraspasos.LastErrorString, vbCritical + vbOKOnly, "Error..."
        End If
    End If




'Set rsRepCab = New ADODB.Recordset
'Set rsRepDet = New ADODB.Recordset
'
''With dtetraspasos
'' ******** Imprime
''    Set rsRepCab = New ADODB.Recordset
'    If rsRepCab.State = 1 Then rsRepCab.Close
'     rsRepCab.Open "SELECT Co_comprobante_rep.* FROM Co_comprobante_rep", db, adOpenKeyset, adLockOptimistic
'    '"select * from co_comprobante_rep ", db, adOpenKeyset, adLockOptimistic
'        While Not rsRepCab.EOF And rsRepCab.RecordCount > 0
'                rsRepCab.Delete
'                rsRepCab.MoveNext
'        Wend
'        '"select * from co_comprobante_rep ", db, adOpenKeyset, adLockOptimistic
'        rsRepCab.AddNew
'        rsRepCab("cod_comp") = rscomprobante1!Cod_Comp
'        rsRepCab("tipo_comp") = rscomprobante1!tipo_Comp
'        rsRepCab("ges_gestion") = "2000"  ' ERRRRRRRRRRRRRRRRRRRRRRRRRR
'        If Not IsNull(rsComprobante!Cod_Trans) Then
'            rsRepCab("cod_trans") = rscomprobante1!Cod_Trans
'          Else
'            rsRepCab("cod_trans") = "-"
'        End If
'        rsRepCab("Num_respaldo") = rscomprobante1!Num_respaldo
'        rsRepCab("fecha_A") = CDate(rscomprobante1!Fecha_A)
'        If rscomprobante1!Codigo_Beneficiario <> "" Then
'            rsRepCab("codigo_beneficiario") = rscomprobante1!Codigo_Beneficiario
'        Else
'            rsRepCab("codigo_beneficiario") = "-"
'        End If
'        rsRepCab("codigo_documento") = rscomprobante1!Codigo_documento
'        rsRepCab("glosa") = rscomprobante1!Glosa
'        If rsComprobante("status") = "S" Then
'            rsComprobante("status") = "A"
'           Else
'            rsRepCab("status") = "S"
'        End If
'        rsRepCab.Update
'       'Set rsRepDet = New ADODB.Recordset
'       If rsRepDet.State = 1 Then rsRepDet.Close
'       rsRepDet.Open "SELECT co_diario_rep.* FROM co_diario_rep ", db, adOpenKeyset, adLockOptimistic
'       'rsRepDet.Open "select * from co_diario_rep ", db, adOpenKeyset, adLockOptimistic
'       While Not rsRepDet.EOF And rsRepDet.RecordCount > 0
'                    rsRepDet.Delete
'                    rsRepDet.MoveNext
'       Wend
'
'        If rsRepDet.State = 1 Then rsRepDet.Close
'        rsRepDet.Open
'        '"select * from co_diario_rep ", db, adOpenKeyset, adLockOptimistic
'        If Not IsNull(rsRepCab("cod_comp")) And Not IsNull(rsRepDet("tipo_comp")) Then
''           Set rsPlanCta = New ADODB.Recordset
''           rsPlanCta.Open "select * from co_diario where cod_comp=" & rsRepCab("cod_comp") & " and tipo_comp='" & rsRepCab("tipo_comp") & "'", db, adOpenKeyset, adLockOptimistic
'
''           Set rsDetalle = New ADODB.Recordset
''
' '          .rscomprobante.Open
'           '"select * from co_diario where cod_comp=" & rsRepCab("cod_comp") & " and tipo_comp='" & rsRepCab("tipo_comp") & "'", db, adOpenKeyset, adLockOptimistic
'
'            'Set DtGDetalle.DataSource = rsDetalle
''            If .rscomprobante.State = 1 Then .rscomprobante.Close
''            .rscomprobante.Open
'          'If .rscomprobante.RecordCount > 0 Then
'           Set rsnombre_cta = New ADODB.Recordset
'
'          'While Not .rscomprobante.EOF
'            rsRepDet.AddNew
'            rsRepDet("cod_comp") = rscomprobante1!Cod_Comp
'            rsRepDet("tipo_comp") = rscomprobante1!tipo_Comp
'            rsRepDet("d_cuenta") = rscomprobante1!D_Cuenta
'            rsRepDet("d_subcta1") = rscomprobante1!D_SubCta1
'            rsRepDet("D_subcta2") = rscomprobante1!D_SubCta2
'            If rsnombre_cta.State = 1 Then rsnombre_cta.Close
'            rsnombre_cta.Open "SELECT NombreCta From CC_Plan_Cuentas WHERE (SubCta1 = '" & Trim(rscomprobante1!D_SubCta1) & _
'            "') AND (SubCta2 = '" & Trim(rscomprobante1!D_SubCta2) & "') AND  (Aux1 = '" & Trim(rscomprobante1!d_Aux1) & _
'            "') AND (Aux2 = '" & Trim(rscomprobante1!d_Aux2) & "') AND (Aux3 = '" & Trim(rscomprobante1!d_Aux3) & "') AND  (Cuenta = '" & _
'            Trim(rscomprobante1!D_Cuenta) & "')", db, adOpenKeyset, adLockReadOnly
'''            rsnombre_cta.Open "SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas  where " & _
''            "(CC_Plan_Cuentas.Cuenta)= '" & rscomprobante1!d_cuenta & "' AND (CC_Plan_Cuentas.SubCta1)= '" & _
''            rscomprobante1!d_subcta1 & "' and  CC_Plan_Cuentas.SubCta2 ='" & rscomprobante1!d_subcta2 & _
''            " ' AND (CC_Plan_Cuentas.Aux1)= '" & rscomprobante1!d_aux1 & "' AND (CC_Plan_Cuentas.Aux2)= ' " & rscomprobante1!d_aux2 & "' AND (CC_Plan_Cuentas.Aux3)= '" & _
''            rscomprobante1!d_aux3 & "'", db, adOpenKeyset, adLockReadOnly
'            MsgBox rsnombre_cta.RecordCount
'            '.nombre_cta .rscomprobante1!d_cuenta, .rscomprobante1!d_subcta1, .rscomprobante1!d_subcta2, .rscomprobante1!d_aux1, .rscomprobante1!d_aux2, .rscomprobante1!d_aux3
'
'            rsRepDet("d_nombre") = rsnombre_cta!NombreCta
'            rsRepDet("d_Aux1") = rscomprobante1("d_Aux1")
'            rsRepDet("d_Aux2") = rscomprobante1("d_Aux2")
'            rsRepDet("d_Aux3") = rscomprobante1("d_Aux3")
'            rsRepDet("D_Cta_larga") = rscomprobante1("D_Cta_larga")
'            rsRepDet("D_MontoBs") = rscomprobante1("D_MontoBs")
'            rsRepDet("D_MontoDl") = rscomprobante1("D_MontoDl")
'            rsRepDet("D_Cambio") = rscomprobante1("D_Cambio")
'            rsRepDet("D_Des_Larga") = rscomprobante1("D_Des_Larga")
'            rsRepDet("H_cuenta") = rscomprobante1("H_cuenta")
'            rsRepDet("H_subcta1") = rscomprobante1("H_subcta1")
'            rsRepDet("H_subcta2") = rscomprobante1("H_subcta2")
'            If rsnombre_cta.State = 1 Then rsnombre_cta.Close
'
'             rsnombre_cta.Open "SELECT NombreCta From CC_Plan_Cuentas WHERE (SubCta1 = '" & Trim(rscomprobante1!H_SubCta1) & _
'            "') AND (SubCta2 = '" & Trim(rscomprobante1!H_SubCta2) & "') AND  (Aux1 = '" & Trim(rscomprobante1!h_Aux1) & _
'            "') AND (Aux2 = '" & Trim(rscomprobante1!h_Aux2) & "') AND (Aux3 = '" & Trim(rscomprobante1!h_Aux3) & "') AND  (Cuenta = '" & _
'            Trim(rscomprobante1!H_Cuenta) & "')", db, adOpenKeyset, adLockReadOnly
'            rsRepDet("h_nombre") = rsnombre_cta!NombreCta
'
''
'            rsRepDet("H_Aux1") = rscomprobante1("H_Aux1")
'            rsRepDet("H_Aux2") = rscomprobante1("H_Aux2")
'            rsRepDet("H_Aux3") = rscomprobante1("H_Aux3")
'            If rscomprobante1("H_Cta_larga") <> "" Then
'                rsRepDet("H_Cta_larga") = rscomprobante1("H_Cta_larga")
'              Else
'                rsRepDet("H_Cta_larga") = "-"
'            End If
'            rsRepDet("H_MontoBs") = rscomprobante1("H_MontoBs")
'            LiteralCry = Str(Int(rscomprobante1("H_MontoBs")))
'            Decimal2 = Str(Round((rscomprobante1("H_MontoBs") - Val(LiteralCry)), 2) * 100)
'            rsRepDet("H_MontoDl") = rscomprobante1("H_MontoDl")
'            rsRepDet("H_Cambio") = rscomprobante1("H_Cambio")
'            rsRepDet("H_Des_Larga") = rscomprobante1("H_Des_Larga")
'            rsRepDet("literal") = Literal(LiteralCry) + " " + Decimal2 + "/100  Bolivianos"
'            'rsRepDet("literal") = Literal(LiteralCry) + " Bolivianos"
'
'            rsRepDet.Update
'            '.rs.MoveNext
'
'          'Wend
'          'End If
'        End If
'
' RepComprob_Conta.Show
'End With

End Sub

Private Sub Cmd_Modificar_Click()
 'With dtetraspasos
  Set Me.DtGrid_comprobante.DataSource = rsComprobante
 'End With
Me.FraGlobal.Enabled = True
Me.Fram_AsientoD.Enabled = True
Me.Fram_AsientoH.Enabled = True
Me.Cmd_Busqueda.Enabled = False
Me.Cmd_Copiar.Enabled = False
'Me.Cmd_GenAciento.Enabled = False
Me.Cmd_GrabaM.Enabled = True
Me.Cmd_IMPRIMIR.Enabled = True
Me.CmdAgregarDetalle.Enabled = True
Me.CmdSalir.Enabled = True
Me.frame_moneda.Visible = True
End Sub


Private Sub Cmd_Normal_Click()
'With dtetraspasos
'Set recsetauxrel = New ADODB.Recordset
'recsetauxrel.CursorLocation = adUseClient
 rsComprobante.Filter = adFilterNone
  Set Me.DtGrid_comprobante.DataSource = rsComprobante
  Fra_Busqueda.Visible = False
    
'    If .rscomprobante.State = 1 Then .rscomprobante.Close
'    .rscomprobante.Open
'    .rscomprobante.Filter = adFilterNone
'    Set Me.DtGrid_comprobante.DataSource = .rscomprobante
'      Fra_Busqueda.Visible = False
'End With
End Sub

Private Sub CmdAgregarDetalle_Click()
  Me.SSTab1.Tab = 0
  Me.frame_moneda.Visible = True
  Me.Cmd_GrabaM.Enabled = True
  Me.CmdSalir.Enabled = False
  Me.Cmd_Modificar.Enabled = False
  Me.CmdAgregarDetalle.Enabled = False
  Me.Cmd_Aprobar.Enabled = False
  Me.Cmd_Busqueda.Enabled = False
  Me.Cmd_Copiar.Enabled = False
  Me.Cmd_Eligir.Enabled = False
  Me.Cmd_IMPRIMIR.Enabled = False
  Me.CmdEstado.Enabled = False
  Me.FraGlobal.Enabled = True
  Me.Fram_AsientoD.Enabled = True
  Me.Fram_AsientoH.Enabled = True
  Me.DtGrid_comprobante.Enabled = False
  Call limpiar
  Me.Txt_Fecha = Format(Me.TFecha, "dd/mm/yyyy")
  Me.Txt_Cambio = Me.TTipoC
  Me.Txt_ges = Year(Me.TFecha)
  Me.cbod_sub2.Clear
  Me.cboH_sub2.Clear
  '''Me.cbodocumento.Clear
  Me.D1documento.Text = ""
  ''Me.cbodoc_descrip.Clear
  Me.D2descripcion.Text = ""
  Me.d1beneficiario.Text = ""
  Me.d2beneficiario.Text = ""
  ''Me.cbobeneficiario.Clear
  ''Me.cbobenef_deno.Clear
  Me.cbod_sub2.AddItem "01"
  Me.cbod_sub2.AddItem "02"
  Me.cbod_sub2.AddItem "03"
  Me.cboH_sub2.AddItem "01"
  Me.cboH_sub2.AddItem "02"
  Me.cboH_sub2.AddItem "03"
  Me.FraGlobal.Enabled = True
Me.Fram_AsientoD.Enabled = True
Me.Fram_AsientoH.Enabled = True
'  With dtetraspasos
'   If .rsdocumento.State = 1 Then .rsdocumento.Close
'     .rsdocumento.Open
'    If .rsbenef_traspaso.State = 1 Then .rsbenef_traspaso.Close
'    .rsbenef_traspaso.Open
    Me.cbod_sub2.Text = Me.cbod_sub2.List(0)
    Me.cboH_sub2.Text = Me.cboH_sub2.List(0)
' End With
Me.CmdEstado.Enabled = True
End Sub


Private Sub DataGrid1_Click()

End Sub

Private Sub CmdEstado_Click()
  rsComprobante.Filter = "status ='N'"
If rsComprobante.RecordCount <> 0 Then
 
Set Me.DtGrid_comprobante.DataSource = rsComprobante
    Me.DtGrid_comprobante.Refresh
Else

    MsgBox "No existen comprobante para aprobar", vbInformation + vbDefaultButton1, "SAF/2000"
    rscompro_N.Filter = adFilterNone
    rsComprobante.Filter = adFilterNone
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
 End If
  Me.Fram_AsientoD.Enabled = True
  Me.Fram_AsientoH.Enabled = True
  Me.FraGlobal.Enabled = True
  Me.Cmd_Modificar = False
  Me.Cmd_GrabaM.Enabled = True
End Sub

Private Sub Cmd_Busqueda_Click()
   Me.Fra_Busqueda.Visible = True
End Sub
Private Sub cmdimprime_grid_Click()
Set rsbenef = New ADODB.Recordset
Set rsimprgrid = New ADODB.Recordset
db.Execute " truncate table impresion_grid"

If rsimprgrid.State = 1 Then rsimprgrid.Close
    rsimprgrid.Open " select * from impresion_grid", db, adOpenKeyset, adLockOptimistic
'MsgBox rsimprgrid.RecordCount
    'AdodcAprob.Recordset.MoveFirst
rsComprobante.MoveFirst
Do While Not rsComprobante.EOF
  rsimprgrid.AddNew
  rsimprgrid!Cod_Comp = rsComprobante!Cod_Comp
  rsimprgrid!tipo_comp = rsComprobante!tipo_comp
  rsimprgrid!Codigo_Beneficiario = rsComprobante!Codigo_Beneficiario
  rsimprgrid!Cod_Trans = rsComprobante!Cod_Trans
  rsimprgrid!org_codigo = rsComprobante!org_codigo
  rsimprgrid!Status = rsComprobante!Status
  If rsbenef.State = 1 Then rsbenef.Close
    rsbenef.Open "select denominacion_beneficiario,codigo_beneficiario from fc_beneficiario where codigo_beneficiario = '" & rsComprobante!Codigo_Beneficiario & "'", db, adOpenKeyset, adLockReadOnly
  If rsbenef.RecordCount <> 0 Then
    rsimprgrid!denom_beneficiario = rsbenef!denominacion_beneficiario
  Else
    rsimprgrid!denom_beneficiario = " "
  End If
  rsimprgrid.Update
  rsComprobante.MoveNext
Loop
Repgrid.Show
End Sub

Private Sub CmdSalir_Click()
  Set Me.DtGrid_comprobante.DataSource = Nothing
  Unload Me
  
End Sub
Private Sub Cmdatras_Click()
If rsComprobante.BOF Then
    rsComprobante.MoveNext
  Else
    rsComprobante.MovePrevious
  End If
End Sub

Private Sub Cmdsgte_Click()
If rsComprobante.EOF Then
    rsComprobante.MovePrevious
  Else
    rsComprobante.MoveNext
  End If
End Sub

Private Sub Cmdinicio_Click()
  rsComprobante.MoveFirst
End Sub

Private Sub Cmdfin_Click()
  rsComprobante.MoveLast
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub d1beneficiario_Change()
Me.d2beneficiario.Text = Me.d1beneficiario.BoundText
End Sub


Private Sub D1documento_Change()
Me.D2descripcion.Text = Me.D1documento.BoundText
End Sub

Private Sub d2beneficiario_Click(Area As Integer)
Me.d1beneficiario.Text = Me.d2beneficiario.BoundText

End Sub

Private Sub D2descripcion_Click(Area As Integer)
Me.D1documento.Text = Me.D2descripcion.BoundText
End Sub


Private Sub DtGrid_comprobante_Click()
'error 6160 de acceso de datos
Call limpiar
Me.TxtComprobante = Me.DtGrid_comprobante.Columns(0).Value
Me.Cmd_Modificar.Enabled = True
Set rscomprobante1 = New ADODB.Recordset
If rscomprobante1.State = 1 Then rscomprobante1.Close
rscomprobante1.Open "SELECT Co_Comprobante_M.Cod_Comp, " & _
"Co_Comprobante_M.Tipo_Comp, Co_Comprobante_M.cod_trans," & _
"Co_Comprobante_M.cod_trans_detalle, Co_Comprobante_M.org_codigo," & _
"Co_Comprobante_M.ges_gestion, Co_Comprobante_M.Num_Respaldo," & _
"Co_Comprobante_M.Fecha_A,Co_Comprobante_M.codigo_beneficiario," & _
"Co_Comprobante_M.codigo_documento,Co_Comprobante_M.Glosa, Co_Comprobante_M.status," & _
"CO_Diario.Cod_Comp_C, CO_Diario.D_Cuenta,CO_Diario.D_Nombre, CO_Diario.D_Subcta1," & _
"CO_Diario.D_SubCta2, CO_Diario.D_Aux1,CO_Diario.D_Aux2, CO_Diario.D_Aux3," & _
"CO_Diario.D_Cta_Larga, CO_Diario.D_Des_Larga,CO_Diario.D_MontoBs, CO_Diario.D_MontoDl," & _
"CO_Diario.D_Cambio, CO_Diario.H_Cuenta,  CO_Diario.H_Nombre, CO_Diario.H_SubCta1," & _
"CO_Diario.H_SubCta2, CO_Diario.H_Aux1, CO_Diario.H_Aux2, CO_Diario.H_Aux3," & _
"CO_Diario.H_Cta_Larga, CO_Diario.H_Des_Larga,CO_Diario.H_MontoBs, CO_Diario.H_MontoDl," & _
"CO_Diario.H_Cambio FROM Co_Comprobante_M INNER JOIN " & _
"CO_Diario ON  Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp AND " & _
" Co_Comprobante_M.Tipo_Comp = CO_Diario.Tipo_Comp where " & _
" co_comprobante_M.cod_comp=" & Val(rsComprobante!Cod_Comp) & _
" and Co_Comprobante_M.Tipo_Comp='" & rsComprobante!tipo_comp & "'", db, adOpenKeyset, adLockOptimistic
  If rscomprobante1.RecordCount <> 0 Then
    Me.Txt_ges = rscomprobante1!Ges_gestion
    If IsNull(rscomprobante1!Fecha_A) Then
      Me.Txt_Fecha = " "
    Else
      Me.Txt_Fecha = Format(rscomprobante1!Fecha_A, "dd/mm/yyyy")
    End If
  '''Me.cbodocumento.Text = .rscomprobante!codigo_documento
    Me.D1documento.Text = rscomprobante1!Codigo_documento
    Me.Txt_Respaldo = rscomprobante1!Num_respaldo
  '''Me.cbobeneficiario.Text = .rscomprobante!codigo_beneficiario
    Me.d1beneficiario.Text = rscomprobante1!Codigo_Beneficiario
    Me.Txt_glosa = rscomprobante1!Glosa
    Me.lblh_cuenta = rscomprobante1!H_Cuenta
    Me.lblh_sub1 = rscomprobante1!H_SubCta1
    Me.cboH_sub2 = rscomprobante1!H_SubCta2
    Me.Txt_Cambio = Val(rscomprobante1!h_Cambio)
    Me.Txth_Bs = Val(rscomprobante1!D_MontoBs)
    Me.Txth_dls = Val(rscomprobante1!h_MontoDl)
    Me.cboH_aux1.Text = rscomprobante1!h_cta_Larga
    Me.cboh_aux1_denom.Text = rscomprobante1!h_des_Larga
    Me.lbld_cuenta = rscomprobante1!D_Cuenta
    Me.lbld_sub1 = rscomprobante1!D_SubCta1
    Me.cbod_sub2 = rscomprobante1!D_SubCta2
    Me.TxtD_Bs = Val(rscomprobante1!D_MontoBs)
    Me.TxtD_Dls = Val(rscomprobante1!D_MontoDl)
    Me.cbod_aux1.Text = rscomprobante1!d_cta_Larga
    Me.cbod_aux1_denom.Text = rscomprobante1!d_des_Larga
    If rscomprobante1!Status = "S" Or rscomprobante1!Status = "A" Then
      Me.Cmd_Modificar.Enabled = False
    End If
  Else
  MsgBox "Comprobantes sin datos", vbCritical + vbDefaultButton1
  End If
End Sub

Private Sub Form_Load()
Me.SSTab1.Tab = 0
Me.frame_moneda.Visible = False
Me.FraGlobal.Enabled = False
Me.Fram_AsientoD.Enabled = False
Me.Fram_AsientoH.Enabled = False
Me.Cmd_GrabaM.Enabled = False
'***
Set rsPago = New ADODB.Recordset
Set rspago_detalle = New ADODB.Recordset
'*************recordset para el grid
Set rsComprobante = New ADODB.Recordset
Me.OptSAprobar.Value = True
'**********recordset para el documento
Set rsdocumento = New ADODB.Recordset
If rsdocumento.State = 1 Then rsdocumento.Close
rsdocumento.Open "SELECT Codigo_Documento, Denominacion_documento FROM ac_documento_respaldo" & _
" ORDER BY Codigo_Documento", db, adOpenKeyset, adLockOptimistic
Set Me.Adodcdocumento.Recordset = rsdocumento
'*********recordset para el beneficiario
Set rsbenef_traspaso = New ADODB.Recordset
If rsbenef_traspaso.State = 1 Then rsbenef_traspaso.Close
rsbenef_traspaso.Open "SELECT fc_beneficiario.codigo_beneficiario, fc_beneficiario.denominacion_beneficiario " & _
"FROM fc_beneficiario, fc_bancos WHERE fc_beneficiario.denominacion_beneficiario = fc_bancos.Bco_descripcion_larga", db, adOpenKeyset, adLockOptimistic
'MsgBox rsbenef_traspaso.RecordCount
Set Me.Adodcbeneficiario.Recordset = rsbenef_traspaso
'**********recordset para cuentas bancarias
Set rscta_corriente = New ADODB.Recordset
  If rscta_corriente.State = 1 Then rscta_corriente.Close
  rscta_corriente.Open "SELECT fc_bancos.Bco_codigo, fc_bancos.Ges_gestion, fc_bancos.bco_descripcion_larga, " & _
  "fc_cuenta_bancaria.Cta_codigo,fc_cuenta_bancaria.Cta_Saldo_Debe," & _
  "fc_cuenta_bancaria.Cta_Saldo_Haber,fc_cuenta_bancaria.Cta_saldo_inicial," & _
  "fc_cuenta_bancaria.Fte_codigo,fc_cuenta_bancaria.Cta_Acumulado," & _
  "fc_cuenta_bancaria.Cta_descripcion_larga,fc_cuenta_bancaria.Cta_saldo_actual " & _
  "FROM fc_cuenta_bancaria INNER JOIN  fc_bancos ON fc_cuenta_bancaria.Bco_codigo = fc_bancos.Bco_codigo " & _
  "AND fc_cuenta_bancaria.Ges_gestion = fc_bancos.Ges_gestion ", db, adOpenKeyset, adLockOptimistic
'recordset para el tipo de cambio
    Set rstipocambio = New ADODB.Recordset
    sql_TC = "select fecha_cambio, Cambio_Oficial  from ac_tipo_cambio  where fecha_cambio = (select max(fecha_cambio) as expr1 from ac_tipo_cambio)"
    rstipocambio.Open sql_TC, db, adOpenKeyset, adLockReadOnly
    TTipoC = rstipocambio!cambio_oficial
    TFecha = rstipocambio!fecha_cambio
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)
With dtetraspasos
'If .rscomprobante.State = 1 Then .rscomprobante.Close
'.rscomprobante.Open
  Set Me.DtGrid_comprobante.DataSource = .rsComprobante
End With
End Sub

Private Sub optbolivianos_Click()
  If Me.optbolivianos.Value = True Then
    Me.Txth_dls.Enabled = False
    Me.Txth_dls.BackColor = &HE0E0E0
    Me.Txth_Bs.Enabled = True
    Me.Txth_Bs.BackColor = &HFFFFFF
  End If
End Sub

Private Sub optconjunto_Click()
Me.cboaprob_inicio.Enabled = True
Me.lblcomprob.Visible = True
Me.cbo_aprob_final.Visible = True
sw1 = 0
End Sub

Private Sub optdolares_Click()
If Me.optdolares.Value = True Then
  Me.Txth_Bs.Enabled = False
  Me.Txth_Bs.BackColor = &HE0E0E0
  Me.Txth_dls.Enabled = True
  Me.Txth_dls.BackColor = &HFFFFFF
End If
End Sub

Private Sub optindividual_Click()
Me.cboaprob_inicio.Enabled = True
Me.lblcomprob.Visible = False
Me.cbo_aprob_final.Visible = False
sw1 = 1
End Sub
Private Sub OptSAprobar_Click()
If Me.OptSAprobar.Value = True Then
    If rsComprobante.State = 1 Then rsComprobante.Close
    sql_grid = "SELECT Co_Comprobante_M.Cod_Comp, " & _
    "Co_Comprobante_M.Tipo_Comp, Co_Comprobante_M.codigo_beneficiario," & _
    "Co_Comprobante_M.status,Co_Comprobante_M.cod_trans,Co_Comprobante_M.org_codigo " & _
    "FROM CO_Diario INNER JOIN  Co_Comprobante_M ON CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp " & _
    "AND  CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp " & _
    "WHERE (CO_Diario.D_Cuenta = '1111') AND (CO_Diario.D_Subcta1 = '02') AND " & _
    "(CO_Diario.H_Cuenta = '1111') AND (CO_Diario.H_SubCta1 = '02') AND " & _
    "(Co_Comprobante_M.Tipo_Comp = 'PCE' OR Co_Comprobante_M.Tipo_Comp = 'TRP' )" & _
    " and  Co_Comprobante_M.Status='N' ORDER BY Co_Comprobante_M.Cod_Comp"
    rsComprobante.Open sql_grid, db, adOpenKeyset, adLockOptimistic
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
End If
End Sub

Private Sub OptTodos_Click()
If Me.OptTodos.Value = True Then
    If rsComprobante.State = 1 Then rsComprobante.Close
    sql_grid = "SELECT Co_Comprobante_M.Cod_Comp, " & _
    "Co_Comprobante_M.Tipo_Comp, Co_Comprobante_M.codigo_beneficiario," & _
    "Co_Comprobante_M.status,Co_Comprobante_M.cod_trans,Co_Comprobante_M.org_codigo " & _
    "FROM CO_Diario INNER JOIN  Co_Comprobante_M ON CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp " & _
    "AND  CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp " & _
    "WHERE (CO_Diario.D_Cuenta = '1111') AND (CO_Diario.D_Subcta1 = '02') AND " & _
    "(CO_Diario.H_Cuenta = '1111') AND (CO_Diario.H_SubCta1 = '02') AND " & _
    "(Co_Comprobante_M.Tipo_Comp = 'PCE' OR Co_Comprobante_M.Tipo_Comp = 'TRP' )" & _
    " and  Co_Comprobante_M.Status <>'E' ORDER BY Co_Comprobante_M.Cod_Comp"
    rsComprobante.Open sql_grid, db, adOpenKeyset, adLockOptimistic
    Set Me.DtGrid_comprobante.DataSource = rsComprobante
End If
End Sub

Private Sub Txt_Cambio_Change()
  Me.TxtD_Dls = Round(Val(Me.TxtD_Bs) * Val(Me.Txt_Cambio), 2)
  Me.TxtD_Bs = Me.TxtD_Bs
  Me.TxtD_Bs = Me.TxtD_Bs
End Sub
Private Sub Txt_Cambio_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        KeyAscii = 0        'Para que no "pite"
        SendKeys "{tab}"    'Envia una pulsación TAB
    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
    'Si después de añadirle la tecla actual no es un número...
        If Not IsNumeric("0" & Me.Txt_Cambio.Text & Chr(KeyAscii)) Then
        '... se desecha esa tecla y se avisa de que no es correcta
            Beep
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub Txth_Bs_Change()
If Me.optbolivianos.Value = True Then
  If Val(Me.Txt_Cambio) <> 0 Then
    Me.Txth_dls = Round(Val(Me.Txth_Bs) / Val(Me.Txt_Cambio), 2)
    Me.TxtD_Bs = Me.Txth_Bs
    Me.TxtD_Dls = Me.Txth_dls
  Else
    Exit Sub
  End If
Else
  Exit Sub
End If
End Sub
Private Sub Txth_Bs_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        KeyAscii = 0        'Para que no "pite"
        SendKeys "{tab}"    'Envia una pulsación TAB
    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
    'Si después de añadirle la tecla actual no es un número...
        If Not IsNumeric("0" & Txth_Bs.Text & Chr(KeyAscii)) Then
        '... se desecha esa tecla y se avisa de que no es correcta
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Txth_dls_Change()
If Me.optdolares.Value = True Then
  If Val(Me.Txt_Cambio) <> 0 Then
    Me.Txth_Bs = Round(Val(Me.Txth_dls) * Val(Me.Txt_Cambio), 2)
    Me.TxtD_Bs = Me.Txth_Bs
    Me.TxtD_Dls = Me.Txth_dls
  Else
    Exit Sub
  End If
Else
  Exit Sub
End If
End Sub

Private Sub Txth_dls_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0        'Para que no "pite"
        SendKeys "{tab}"    'Envia una pulsación TAB
    ElseIf KeyAscii <> 8 Then   'El 8 es la tecla de borrar (backspace)
    'Si después de añadirle la tecla actual no es un número...
        If Not IsNumeric("0" & Me.Txth_dls.Text & Chr(KeyAscii)) Then
        '... se desecha esa tecla y se avisa de que no es correcta
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Public Sub limpiar()
Me.Txt_Cambio = ""
Me.Txt_Fecha = ""
Me.Txt_Fecha = ""
Me.Txt_glosa = ""
Me.Txt_Respaldo = ""
Me.TxtComprobante = ""
Me.TxtD_Bs = ""
Me.TxtD_Dls = ""
Me.Txth_Bs = ""
Me.Txth_dls = ""
Me.lbld_cuenta = "1111"
Me.lbld_sub1 = "02"
Me.lblh_cuenta = "1111"
Me.lblh_sub1 = "02"
End Sub
Public Sub genera_codigo()
'With dtetraspasos
Set rsCorrelativo = New ADODB.Recordset
If rsCorrelativo.State = 1 Then rsCorrelativo.Close
  rsCorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'cmbte')", db, adOpenKeyset, adLockOptimistic
  rsCorrelativo.MoveFirst
  num_comprobante = rsCorrelativo!Numero_correlativo + 1
  rsCorrelativo!Numero_correlativo = rsCorrelativo!Numero_correlativo + 1
  rsCorrelativo.Update
'End With
End Sub


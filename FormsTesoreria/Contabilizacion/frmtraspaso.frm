VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmtraspasos 
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame_moneda 
      Caption         =   "Tipo de Moneda"
      Height          =   495
      Left            =   4125
      TabIndex        =   109
      Top             =   3930
      Width           =   5235
      Begin VB.OptionButton optdolares 
         Caption         =   "Dólares"
         Height          =   270
         Left            =   2550
         TabIndex        =   111
         Top             =   165
         Width           =   1590
      End
      Begin VB.OptionButton optbolivianos 
         Caption         =   "Bolivianos"
         Height          =   270
         Left            =   585
         TabIndex        =   110
         Top             =   165
         Width           =   1350
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
      Left            =   5040
      TabIndex        =   108
      Top             =   3360
      Visible         =   0   'False
      Width           =   3876
      Begin VB.ComboBox Cmbo_Atributo 
         DataMember      =   "comprobante"
         DataSource      =   "dtetraspasos"
         Height          =   315
         ItemData        =   "frmtraspaso.frx":0000
         Left            =   240
         List            =   "frmtraspaso.frx":0016
         TabIndex        =   34
         Text            =   "Cod_Comp"
         Top             =   426
         Width           =   1284
      End
      Begin VB.ComboBox Cmbo_Operador 
         Height          =   315
         ItemData        =   "frmtraspaso.frx":0063
         Left            =   1635
         List            =   "frmtraspaso.frx":0076
         TabIndex        =   35
         Text            =   "="
         Top             =   426
         Width           =   672
      End
      Begin VB.TextBox Text_Valor 
         Height          =   336
         Left            =   2550
         TabIndex        =   36
         Text            =   "1"
         Top             =   405
         Width           =   900
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
         Left            =   270
         TabIndex        =   37
         Top             =   1335
         Width           =   996
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
         Left            =   2610
         TabIndex        =   39
         Top             =   1290
         Width           =   936
      End
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
         Left            =   1455
         TabIndex        =   38
         Top             =   1305
         Width           =   888
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   1335
      TabIndex        =   107
      Top             =   7065
      Width           =   2670
      Begin VB.CommandButton cmdfin 
         Height          =   555
         Left            =   1905
         Picture         =   "frmtraspaso.frx":008B
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdinicio 
         Height          =   555
         Left            =   105
         Picture         =   "frmtraspaso.frx":04CD
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdsgte 
         Height          =   555
         Left            =   1365
         Picture         =   "frmtraspaso.frx":090F
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdatras 
         DisabledPicture =   "frmtraspaso.frx":0D51
         Height          =   555
         Left            =   705
         Picture         =   "frmtraspaso.frx":1193
         Style           =   1  'Graphical
         TabIndex        =   40
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
      Height          =   1830
      Left            =   4200
      TabIndex        =   104
      Top             =   1680
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmd_aprob_cancel 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   3060
         TabIndex        =   33
         Top             =   1395
         Width           =   1350
      End
      Begin VB.CommandButton cmd_aprob_aceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   765
         TabIndex        =   32
         Top             =   1395
         Width           =   1350
      End
      Begin VB.ComboBox cbo_aprob_final 
         Height          =   315
         Left            =   3630
         TabIndex        =   31
         Top             =   765
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.ComboBox cboaprob_inicio 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmtraspaso.frx":15D5
         Left            =   1185
         List            =   "frmtraspaso.frx":15D7
         TabIndex        =   30
         Top             =   795
         Width           =   1125
      End
      Begin VB.OptionButton optconjunto 
         Caption         =   "Conjunto"
         Height          =   360
         Left            =   2370
         TabIndex        =   29
         Top             =   285
         Width           =   960
      End
      Begin VB.OptionButton optindividual 
         Caption         =   "Individual"
         Height          =   195
         Left            =   540
         TabIndex        =   28
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblcomprob 
         Caption         =   "No. Comprob "
         Height          =   225
         Left            =   2400
         TabIndex        =   106
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label20 
         Caption         =   "No. Comprob "
         Height          =   225
         Left            =   120
         TabIndex        =   105
         Top             =   870
         Width           =   1065
      End
   End
   Begin MSDataGridLib.DataGrid DtGrid_comprobante 
      Height          =   5895
      Left            =   1380
      TabIndex        =   1
      Top             =   1080
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   10398
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
   Begin VB.Frame FraGlobal 
      Enabled         =   0   'False
      Height          =   2805
      Left            =   4140
      TabIndex        =   47
      Top             =   1020
      Width           =   5220
      Begin VB.Frame Frame_Plan 
         Caption         =   "Plan_cuentas"
         Height          =   2655
         Left            =   1440
         TabIndex        =   53
         Top             =   3720
         Visible         =   0   'False
         Width           =   7335
         Begin VB.CommandButton Cmd_Eligir 
            Caption         =   "Elegir"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   2160
            Width           =   1695
         End
         Begin MSDataGridLib.DataGrid DtGrid_Plan 
            Height          =   1815
            Left            =   240
            TabIndex        =   55
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
         Left            =   720
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2175
         Width           =   4290
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   0
         TabIndex        =   52
         Top             =   -48
         Width           =   7110
      End
      Begin VB.TextBox Text_Tipo 
         Height          =   288
         Left            =   2685
         TabIndex        =   51
         Text            =   "Comprobante de Traspasos"
         Top             =   255
         Width           =   2430
      End
      Begin VB.TextBox Txt_Fecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   3
         EndProperty
         Height          =   288
         Left            =   6012
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Txt_ges 
         Height          =   288
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   696
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   48
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
         Left            =   1725
         MaxLength       =   9
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin MSDataListLib.DataCombo D1documento 
         Bindings        =   "frmtraspaso.frx":15D9
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   1215
         TabIndex        =   11
         Top             =   1125
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo_Documento"
         BoundColumn     =   "Denominacion_documento"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo D2descripcion 
         Bindings        =   "frmtraspaso.frx":15F7
         DataField       =   "Denominacion_documento"
         DataSource      =   "Adodcdocumento"
         Height          =   315
         Left            =   2505
         TabIndex        =   12
         Top             =   1065
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_documento"
         BoundColumn     =   "Codigo_Documento"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo d2beneficiario 
         Bindings        =   "frmtraspaso.frx":1615
         DataField       =   "denominacion_beneficiario"
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo d1beneficiario 
         Bindings        =   "frmtraspaso.frx":1636
         DataField       =   "codigo_beneficiario"
         DataSource      =   "Adodcbeneficiario"
         Height          =   315
         Left            =   1260
         TabIndex        =   14
         Top             =   1800
         Width           =   1305
         _ExtentX        =   2302
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
         TabIndex        =   65
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label_Respaldo 
         Caption         =   "Numero de Respaldo"
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label Label_AntComp 
         Caption         =   "Tipo Comprobante Anterior:"
         Height          =   285
         Left            =   135
         TabIndex        =   62
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label Label_Fecha 
         Caption         =   "Fecha:"
         Height          =   288
         Left            =   5304
         TabIndex        =   61
         Top             =   732
         Width           =   612
      End
      Begin VB.Label Label8 
         Caption         =   "Glosa"
         Height          =   252
         Left            =   120
         TabIndex        =   60
         Top             =   2292
         Width           =   636
      End
      Begin VB.Label Label5 
         Caption         =   "Gestion:"
         Height          =   285
         Left            =   3285
         TabIndex        =   59
         Top             =   735
         Width           =   750
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiario:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   1830
         Width           =   870
      End
      Begin VB.Label Label11 
         Caption         =   "Documento Respaldo"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   1035
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Nro Comprobante:"
         Enabled         =   0   'False
         Height          =   288
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.Frame FraOpcionesDetalle 
      Height          =   6885
      Left            =   60
      TabIndex        =   46
      Top             =   990
      Width           =   1180
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":1657
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   6060
         Width           =   930
      End
      Begin VB.CommandButton cmdimprime_grid 
         Caption         =   "Imprime Grid"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":1A99
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   5290
         Width           =   930
      End
      Begin VB.CommandButton Cmd_Busqueda 
         Caption         =   "Busqueda"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":1EDB
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1330
         Width           =   945
      End
      Begin VB.CommandButton CmdAgregarDetalle 
         Caption         =   "Adicionar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":231D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton Cmd_GrabaM 
         Caption         =   "Grabar"
         Height          =   765
         Left            =   135
         Picture         =   "frmtraspaso.frx":275F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2570
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Modificar 
         Caption         =   "Modificar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":2BA1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   740
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Cancelar 
         Caption         =   "Cancelar"
         Height          =   585
         Left            =   135
         Picture         =   "frmtraspaso.frx":2FE3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3340
         Width           =   945
      End
      Begin VB.CommandButton Cmd_Aprobar 
         Caption         =   "Aprobar"
         Height          =   705
         Left            =   135
         Picture         =   "frmtraspaso.frx":30E5
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3930
         Width           =   945
      End
      Begin VB.CommandButton CmdEstado 
         Caption         =   "&Estado"
         Height          =   495
         Left            =   180
         TabIndex        =   10
         Top             =   6270
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton Cmd_Copiar 
         Caption         =   "Copiar"
         Height          =   645
         Left            =   135
         Picture         =   "frmtraspaso.frx":3527
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton Cmd_IMPRIMIR 
         Caption         =   "Imprimir"
         Height          =   645
         Left            =   135
         Picture         =   "frmtraspaso.frx":3A59
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4640
         Width           =   945
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9705
      TabIndex        =   0
      Top             =   0
      Width           =   9765
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
         TabIndex        =   103
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
         TabIndex        =   45
         Top             =   1125
         Width           =   8415
      End
      Begin VB.Label Label7 
         Height          =   225
         Left            =   10485
         TabIndex        =   44
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7740
         TabIndex        =   43
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Administrativa Financiera"
         Height          =   225
         Left            =   1020
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   630
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3165
      Left            =   4140
      TabIndex        =   64
      Top             =   4620
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   5583
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "Crédito"
      TabPicture(0)   =   "frmtraspaso.frx":40C3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fram_AsientoH"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Débito"
      TabPicture(1)   =   "frmtraspaso.frx":40DF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fram_AsientoD"
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
         Height          =   2445
         Left            =   -74820
         TabIndex        =   84
         Top             =   360
         Width           =   5475
         Begin VB.ComboBox cbod_sub2 
            Height          =   315
            ItemData        =   "frmtraspaso.frx":40FB
            Left            =   3525
            List            =   "frmtraspaso.frx":4108
            TabIndex        =   23
            Top             =   240
            Width           =   750
         End
         Begin VB.Frame FrameD_CtaCorriente 
            Caption         =   "Cuentas corrientes de Bancos"
            Height          =   1335
            Left            =   45
            TabIndex        =   87
            Top             =   1080
            Width           =   5400
            Begin VB.ComboBox cbod_aux1_denom 
               Height          =   315
               Left            =   2145
               TabIndex        =   25
               Top             =   270
               Width           =   3165
            End
            Begin VB.ComboBox cbod_aux1 
               Height          =   315
               Left            =   795
               TabIndex        =   24
               Top             =   240
               Width           =   1260
            End
            Begin MSDataListLib.DataCombo TxtD_Nom3_Corriente 
               Height          =   315
               Left            =   2145
               TabIndex        =   88
               Top             =   960
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Nom2_Corriente 
               Height          =   315
               Left            =   2145
               TabIndex        =   89
               Top             =   615
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Aux3_Corriente 
               Height          =   315
               Left            =   825
               TabIndex        =   90
               Top             =   960
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtD_Aux2_Corriente 
               Height          =   315
               Left            =   840
               TabIndex        =   91
               Top             =   615
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin VB.Label Label14 
               Caption         =   "Auxiliar 1:"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   285
               Width           =   735
            End
            Begin VB.Label Label13 
               Caption         =   "Auxiliar 2:"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   690
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "Auxiliar 3"
               Height          =   195
               Left            =   120
               TabIndex        =   93
               Top             =   1035
               Width           =   735
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "Descripcion:"
               Height          =   255
               Left            =   2955
               TabIndex        =   92
               Top             =   105
               Width           =   1080
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
            TabIndex        =   86
            Top             =   615
            Width           =   1395
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
            TabIndex        =   85
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label lbld_sub1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "02"
            Height          =   255
            Left            =   2145
            TabIndex        =   102
            Top             =   285
            Width           =   435
         End
         Begin VB.Label lbld_cuenta 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1111"
            Height          =   240
            Left            =   795
            TabIndex        =   101
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label_Cuenta 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   105
            TabIndex        =   100
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "MontoDls"
            Height          =   255
            Left            =   2355
            TabIndex        =   99
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label_MontoBs 
            Caption         =   "Monto_Bs"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label_Cta2 
            Caption         =   "Sub_Cta2:"
            Height          =   255
            Left            =   2640
            TabIndex        =   97
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label_Cta1 
            Caption         =   "Sub_Cta1:"
            Height          =   255
            Left            =   1320
            TabIndex        =   96
            Top             =   330
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
         Height          =   2595
         Left            =   120
         TabIndex        =   66
         Top             =   255
         Width           =   5490
         Begin VB.TextBox Txt_Cambio 
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
            Left            =   4860
            TabIndex        =   20
            Top             =   330
            Width           =   540
         End
         Begin VB.Frame FrameH_CtaCorriente 
            Caption         =   "Cuentas corrientes de Bancos"
            Height          =   1410
            Left            =   60
            TabIndex        =   67
            Top             =   1020
            Width           =   5385
            Begin VB.ComboBox cboh_aux1_denom 
               Height          =   315
               Left            =   2175
               TabIndex        =   19
               Top             =   360
               Width           =   3180
            End
            Begin VB.ComboBox cboH_aux1 
               Height          =   315
               Left            =   900
               TabIndex        =   18
               Top             =   345
               Width           =   1185
            End
            Begin MSDataListLib.DataCombo TxtH_Nom3_Corriente 
               Height          =   315
               Left            =   2055
               TabIndex        =   68
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
               Left            =   2160
               TabIndex        =   69
               Top             =   690
               Width           =   3165
               _ExtentX        =   5583
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Aux3_Corriente 
               Height          =   315
               Left            =   870
               TabIndex        =   70
               Top             =   1020
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo TxtH_Aux2_Corriente 
               Height          =   315
               Left            =   885
               TabIndex        =   71
               Top             =   675
               Width           =   1215
               _ExtentX        =   2143
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
               TabIndex        =   75
               Top             =   120
               Width           =   1080
            End
            Begin VB.Label Label17 
               Caption         =   "Auxiliar 3"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   1095
               Width           =   735
            End
            Begin VB.Label Label16 
               Caption         =   "Auxiliar 2:"
               Height          =   255
               Left            =   105
               TabIndex        =   73
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "Auxiliar 1:"
               Height          =   240
               Left            =   120
               TabIndex        =   72
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3330
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   675
            Width           =   1350
         End
         Begin VB.ComboBox cboH_sub2 
            Height          =   315
            ItemData        =   "frmtraspaso.frx":4118
            Left            =   3330
            List            =   "frmtraspaso.frx":4125
            TabIndex        =   17
            Top             =   345
            Width           =   690
         End
         Begin VB.Label Label_Dl 
            Caption         =   "MontoDls"
            Height          =   255
            Left            =   2520
            TabIndex        =   83
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label_Bs 
            Caption         =   "Monto_Bs"
            Height          =   255
            Left            =   135
            TabIndex        =   82
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label_Cambio 
            Caption         =   "Cambio_Dl:"
            Height          =   255
            Left            =   4050
            TabIndex        =   81
            Top             =   375
            Width           =   840
         End
         Begin VB.Label LabelD_Cuenta 
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   60
            TabIndex        =   80
            Top             =   360
            Width           =   600
         End
         Begin VB.Label LabelH_Cta2 
            Caption         =   "Sub_Cta2:"
            Height          =   210
            Left            =   2505
            TabIndex        =   79
            Top             =   345
            Width           =   735
         End
         Begin VB.Label LabelH_Cuenta 
            Alignment       =   2  'Center
            Caption         =   "Sub_Cta1:"
            Height          =   210
            Left            =   1320
            TabIndex        =   78
            Top             =   345
            Width           =   735
         End
         Begin VB.Label lblh_cuenta 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1111"
            Height          =   240
            Left            =   690
            TabIndex        =   77
            Top             =   345
            Width           =   540
         End
         Begin VB.Label lblh_sub1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "02"
            Height          =   240
            Left            =   2100
            TabIndex        =   76
            Top             =   330
            Width           =   330
         End
      End
   End
End
Attribute VB_Name = "frmtraspasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public num_comprobante As Integer ' vaiable donde se almacena nùmero de comprobante
'********RECORDSETS
Dim rscomprobante1 As ADODB.Recordset
Dim rsdocumento As ADODB.Recordset
Dim rsbenef_traspaso As ADODB.Recordset
Dim rscta_corriente As ADODB.Recordset
Dim rsComprobante As ADODB.Recordset
Dim rsdiario As ADODB.Recordset
Dim rscorrelativo As ADODB.Recordset
Dim rscomprobante_M As ADODB.Recordset
Dim rscompro_N As ADODB.Recordset
Dim rspago As ADODB.Recordset
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
Me.cbod_aux1_denom.Text = rscta_corriente!Cta_descripcion_larga
End With
End Sub
Private Sub cbod_aux1_denom_Change()
'On Error GoTo err2
'With dtetraspasos
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_descripcion_larga= '" & Trim(Me.cbod_aux1_denom.Text) & "'"
Me.cbod_aux1.Text = rscta_corriente!Cta_codigo
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
Me.cbod_aux1.Text = rscta_corriente!Cta_codigo
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
    If rscta_corriente!Fte_codigo = "41" Then
      Me.cbod_aux1.AddItem rscta_corriente!Cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
    End If
  Case "02"
    If rscta_corriente!Fte_codigo = "43" Then
      Me.cbod_aux1.AddItem rscta_corriente!Cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
    End If
  Case "03"
    If rscta_corriente!Fte_codigo = "80" Then
      Me.cbod_aux1.AddItem rscta_corriente!Cta_codigo
      Me.cbod_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
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
rscta_corriente.MoveFirst
rscta_corriente.Filter = adFilterNone
rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
'rscta_corriente.Find "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
'MsgBox rscta_corriente.RecordCount
Me.cboh_aux1_denom.Text = rscta_corriente!Cta_descripcion_larga
End Sub

Private Sub cboH_aux1_Click()
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'.rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_codigo= '" & Trim(Me.cboH_aux1.Text) & "'"
Me.cboh_aux1_denom.Text = rscta_corriente!Cta_descripcion_larga

End Sub
Private Sub cboh_aux1_denom_Click()
'If .rscta_corriente.State = 1 Then .rscta_corriente.Close
'.rscta_corriente.Open
rscta_corriente.Filter = adFilterNone
rscta_corriente.MoveFirst
rscta_corriente.Filter = "cta_descripcion_larga= '" & Trim(Me.cboh_aux1_denom.Text) & "'"
Me.cboH_aux1.Text = rscta_corriente!Cta_codigo
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
  If rscta_corriente!Fte_codigo = "41" Then
    Me.cboH_aux1.AddItem rscta_corriente!Cta_codigo
    Me.cboh_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
  End If
Case "02"
  If rscta_corriente!Fte_codigo = "43" Then
   Me.cboH_aux1.AddItem rscta_corriente!Cta_codigo
   Me.cboh_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
  End If
Case "03"
  If rscta_corriente!Fte_codigo = "80" Then
    Me.cboH_aux1.AddItem rscta_corriente!Cta_codigo
    Me.cboh_aux1_denom.AddItem rscta_corriente!Cta_descripcion_larga
  End If
End Select
rscta_corriente.MoveNext
Loop
Me.cboH_aux1.Text = Me.cboH_aux1.List(0)
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
              rscomprobante_M!Status = "S"
              rscomprobante_M!Fecha_A = CDate(Format(Date, "dd/mm/yyyy"))
              rscomprobante_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
              rscomprobante_M.Update
            Set rsdiario = New ADODB.Recordset
            If rsdiario.State = 1 Then rsdiario.Close
              rsdiario.Open "SELECT * FROM CO_Diario " & _
              "WHERE Cod_comp=" & Val(Trim(Me.cboaprob_inicio.Text)), db, adOpenKeyset, adLockOptimistic
              rsdiario!cod_comp_C = Val(Trim(Me.cboaprob_inicio.Text))
              rsdiario.Update
            Set rspago = New ADODB.Recordset
            Set rspago_detalle = New ADODB.Recordset
          If rspago.State = 1 Then rspago.Close
            rspago.Open "SELECT * FROM pagos WHERE (ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
          
          '.Connection1.BeginTrans
          '*********ADICION A LA TABLA PAGO
            rspago.AddNew
            rspago!Ges_gestion = Trim(rscomprobante_M!Ges_gestion)
            rspago!org_codigo = "999"
            rspago!codigo_pago = Trim(rscomprobante_M!Cod_Comp)
            '.rspago!nro_comprobante_anterior = .rscomprobante!Cod_Comp
            rspago!tipo_comp = "TRP"
            rspago!codigo_orden = Trim(rscomprobante_M!Num_Respaldo)
            rspago!codigo_documento = Trim(rscomprobante_M!codigo_documento)
            rspago!fecha_egreso = CDate(Format(rscomprobante_M!Fecha_A, "dd/mm/yyyy"))
            rspago!monto_Bolivianos = rsdiario!D_MontoBs
            rspago!monto_Dolares = rsdiario!D_MontoDl
            rspago!liquido_pagar = rsdiario!D_MontoBs
            rspago!estado_aprobacion = "N"
            rspago!estado_contabilidad = "P"
            rspago!justificacion = Trim(rscomprobante_M!Glosa)
            rspago!estado_pagado = "S"
            rspago!Usr_Usuario = Trim(rscomprobante_M!Usr_Usuario)
            rspago!fecha_aprueba = CDate(Format(Date, "dd/mm/yyyy"))
            rspago!hora_aprueba = (Format(Time, "hh:mm:ss"))
            rspago!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
            rspago!hora_registro = (Format(Time, "hh:mm:ss"))
            '********ADICION A LA TABLA PAGO DETALLE
            If rspago_detalle.State = 1 Then rspago_detalle.Close
            rspago_detalle.Open "SELECT * FROM pago_detalle WHERE (Ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
            'MsgBox rspago_detalle.RecordCount
            rspago_detalle.AddNew
            rspago_detalle!Ges_gestion = Trim(rscomprobante_M!Ges_gestion)
            rspago_detalle!org_codigo = "999"
            rspago_detalle!codigo_pago = Trim(Str(rscomprobante_M!Cod_Comp))
            rspago_detalle!codigo_pago_detalle = "1"
            rspago_detalle!Codigo_Beneficiario = Trim(rscomprobante_M!Codigo_Beneficiario)
            rspago_detalle!tipo_cambio = rsdiario!D_Cambio
            rspago_detalle!monto_total = rsdiario!D_MontoBs
            rspago_detalle!departamento = "La Paz"
            rspago_detalle!honorarios = "N"
            ''''''''''''
            rspago_detalle!cta_codigo_destino = Trim(rsdiario!D_Cta_Larga)
            rscta_corriente.Filter = "cta_codigo='" & rsdiario!D_Cta_Larga & "'"
            rspago_detalle!banco_destino = Trim(rscta_corriente!bco_descripcion_larga)
            rscta_corriente.Filter = adFilterNone
            rspago_detalle!Cta_codigo = Trim(rsdiario!H_Cta_Larga)
            rspago_detalle!cheque_o_trf = "T"
            rspago_detalle!tipo_cambio = rsdiario!D_Cambio
            rspago_detalle!estado_aprobacion = "N"
            rspago_detalle!monto_Bolivianos = rsdiario!D_MontoBs
            rspago_detalle!monto_Dolares = rsdiario!D_MontoDl
            rspago_detalle!fecha_pago = CDate(Format(Date, "dd/mm/yyyy"))
            'rspago_detalle!departamento=
            'rspago_detalle!beneficiario_destino=
            
            rspago_detalle!Usr_Usuario = Trim(rscomprobante_M!Usr_Usuario)
            rspago_detalle!fecha_registro = Format(Date, "dd/mm/yyyy")
            rspago_detalle!hora_registro = Format(Time, "hh:mm:ss")
           '********ACTUALIZACION MOVIMIENTOS DEBE Y HABER EN LA CUENTA BANCARIA
            Set rsfc_cuenta_bancaria = New ADODB.Recordset
            'CTA BANCARIA DEL DEBE, LA QUE RECIBE
            If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
            rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!D_Cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
            rsfc_cuenta_bancaria.MoveFirst
            rsfc_cuenta_bancaria!Cta_Saldo_Debe = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Debe), 0, rsfc_cuenta_bancaria!Cta_Saldo_Debe) + rsdiario!D_MontoBs
            rsfc_cuenta_bancaria.Update
            ' CTA BANCARIA DEL HABER, LA QUE DA
            If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
            rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!H_Cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
            rsfc_cuenta_bancaria!Cta_Saldo_Haber = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Haber), 0, rsfc_cuenta_bancaria!Cta_Saldo_Haber) + Val(rsdiario!H_MontoBs)
            rsfc_cuenta_bancaria.Update
            rspago.Update
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
                    rscomprobante_M!Status = "S"
                    rscomprobante_M!Fecha_A = CDate(Format(Date, "dd/mm/yyyy"))
                    rscomprobante_M!Cod_Trans = Trim(Me.cboaprob_inicio.Text)
                    rscomprobante_M.Update
                    rsdiario.MoveFirst
                    rsdiario.Filter = adFilterNone
                    rsdiario.Filter = "cod_comp=" & i
                    'rsdiario.Find "cod_comp=" & i
                    If rspago.State = 1 Then rspago.Close
                    rspago.Open "SELECT * FROM pagos WHERE (ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
                    If rspago_detalle.State = 1 Then rspago_detalle.Close
                    rspago_detalle.Open "SELECT * FROM pago_detalle WHERE (Ges_gestion = '9999')", db, adOpenKeyset, adLockOptimistic
                  '*********ADICION A LA TABLA PAGO
                    rspago.AddNew
                    rspago!Ges_gestion = rscomprobante_M!Ges_gestion
                    rspago!org_codigo = "999"
                    rspago!codigo_pago = rscomprobante_M!Cod_Comp
                    '.rspago!nro_comprobante_anterior = .rscomprobante!Cod_Comp
                    rspago!tipo_comp = "PCE"
                    rspago!codigo_orden = rscomprobante_M!Num_Respaldo
                    rspago!codigo_documento = rscomprobante_M!codigo_documento
                    rspago!fecha_egreso = CDate(Format(rscomprobante_M!Fecha_A, "dd/mm/yyyy"))
                    rspago!monto_Bolivianos = rsdiario!D_MontoBs
                    rspago!monto_Dolares = rsdiario!D_MontoDl
                    rspago!liquido_pagar = rsdiario!D_MontoBs
                    'celia rspago!estado_aprobacion = "N" o "A"
                    rspago!estado_contabilidad = "P"
                    rspago!justificacion = rscomprobante_M!Glosa
                    rspago!estado_pagado = "S"
                    rspago!Usr_Usuario = rscomprobante_M!Usr_Usuario
                    rspago!fecha_aprueba = CDate(Format(Date, "dd/mm/yyyy"))
                    rspago!hora_aprueba = (Format(Time, "hh:mm:ss"))
                    rspago!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
                    rspago!hora_registro = (Format(Time, "hh:mm:ss"))
                    '********ADICION A LA TABLA PAGO DETALLE
                    rspago_detalle.AddNew
                    rspago_detalle!Ges_gestion = rscomprobante_M!Ges_gestion
                    rspago_detalle!org_codigo = "999"
                    rspago_detalle!codigo_pago = Str(rscomprobante_M!Cod_Comp)
                    rspago_detalle!codigo_pago_detalle = "1"
                    rspago_detalle!Codigo_Beneficiario = rscomprobante_M!Codigo_Beneficiario
                    rspago_detalle!tipo_cambio = rsdiario!D_Cambio
                    rspago_detalle!monto_total = rsdiario!D_MontoBs
                    rspago_detalle!departamento = "La Paz"
                    rspago_detalle!honorarios = "N"
                    ''''''''''''
                    rspago_detalle!cta_codigo_destino = rsdiario!D_Cta_Larga
                    rscta_corriente.Filter = "cta_codigo='" & rsdiario!D_Cta_Larga & "'"
                    rspago_detalle!banco_destino = rscta_corriente!bco_descripcion_larga
                    rscta_corriente.Filter = adFilterNone
                    rspago_detalle!Cta_codigo = rsdiario!H_Cta_Larga
                    rspago_detalle!cheque_o_trf = "T"
                    rspago_detalle!tipo_cambio = rsdiario!D_Cambio
                    rspago_detalle!estado_aprobacion = "N"
                    rspago_detalle!monto_Bolivianos = rsdiario!D_MontoBs
                    rspago_detalle!monto_Dolares = rsdiario!D_MontoDl
                    rspago_detalle!fecha_pago = CDate(Format(Date, "dd/mm/yyyy"))
                    'rspago_detalle!departamento=
                    'rspago_detalle!beneficiario_destino=
                    
                    rspago_detalle!Usr_Usuario = rscomprobante_M!Usr_Usuario
                    rspago_detalle!fecha_registro = Format(Date, "dd/mm/yyyy")
                    rspago_detalle!hora_registro = Format(Time, "hh:mm:ss")
                   '********ACTUALIZACION MOVIMIENTOS DEBE Y HABER EN LA CUENTA BANCARIA
                    Set rsfc_cuenta_bancaria = New ADODB.Recordset
                    'CTA BANCARIA DEL DEBE, LA QUE RECIBE
                    If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
                    rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!D_Cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
                    rsfc_cuenta_bancaria.MoveFirst
                    rsfc_cuenta_bancaria!Cta_Saldo_Debe = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Debe), 0, rsfc_cuenta_bancaria!Cta_Saldo_Debe) + rsdiario!D_MontoBs
                    rsfc_cuenta_bancaria.Update
                    ' CTA BANCARIA DEL HABER, LA QUE DA
                    If rsfc_cuenta_bancaria.State = 1 Then rsfc_cuenta_bancaria.Close
                    rsfc_cuenta_bancaria.Open " select * from fc_cuenta_bancaria where cta_codigo= '" & Trim(rsdiario!H_Cta_Larga) & "'", db, adOpenKeyset, adLockOptimistic
                    rsfc_cuenta_bancaria!Cta_Saldo_Haber = IIf(IsNull(rsfc_cuenta_bancaria!Cta_Saldo_Haber), 0, rsfc_cuenta_bancaria!Cta_Saldo_Haber) + Val(rsdiario!H_MontoBs)
                    rsfc_cuenta_bancaria.Update
                    rspago.Update
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
        Me.Frame_aprobacion.Visible = False
        Me.FraGlobal.Visible = False
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
Me.FraGlobal.Visible = False
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
    rsComprobante.Open
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
        rsComprobante.Filter = "status='" & Trim(Me.Text_Valor) & "'"
     Case Else
        rsComprobante.Filter = "status='" & Trim(Me.Text_Valor) & "'"
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
If rsComprobante!Status = "S" Then Me.Cmd_Modificar.Enabled = False
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
    rscomprobante_M!cod_trans_detalle = "1"
    rscomprobante_M!org_codigo = "999"
    rscomprobante_M!Ges_gestion = Trim(Me.Txt_ges)
    rscomprobante_M!Num_Respaldo = Trim(Me.Txt_Respaldo)
    rscomprobante_M!Fecha_A = CDate(Format(Date, "dd/mm/yyyy"))
    rscomprobante_M!Codigo_Beneficiario = Trim(Me.d1beneficiario.Text)
    rscomprobante_M!codigo_documento = Trim(Me.D1documento.Text)
    rscomprobante_M!Glosa = Trim(Me.Txt_glosa)
    rscomprobante_M!Status = "N"
    '****'''''revisar codigo de usuario
    rscomprobante_M!Usr_Usuario = GlUsuario ' variable global de usuario glusuario
    rscomprobante_M!fecha_registro = CDate(Format(Date, "mm/dd/yyyy"))
    rscomprobante_M!hora_registro = Format(Time, "hh:mm:ss")
    '********ADICION AL DIARIO
    rsdiario.AddNew
    rsdiario!tipo_comp = "TRP"
    rsdiario!cod_comp_C = 0
    rsdiario!d_cuenta = "1111"
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
    rsdiario!d_subcta1 = "02"
    rsdiario!d_subcta2 = Trim(Me.cbod_sub2.Text)
    rsdiario!d_Aux1 = "02"
    rsdiario!d_Aux2 = "00"
    rsdiario!d_Aux3 = "00"
    rsdiario!D_Cta_Larga = Trim(Me.cbod_aux1.Text)
    rsdiario!D_Des_Larga = Trim(Me.cbod_aux1_denom.Text)
    rsdiario!D_MontoBs = Val(Me.TxtD_Bs)
    rsdiario!D_MontoDl = Val(Me.TxtD_Dls)
    rsdiario!D_Cambio = Val(Me.Txt_Cambio)
  '*************
    rsdiario!h_cuenta = "1111"
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
    rsdiario!h_subcta1 = "02"
    rsdiario!h_subcta2 = Trim(Me.cboH_sub2.Text)
    rsdiario!h_Aux1 = "02"
    rsdiario!h_Aux2 = "00"
    rsdiario!h_Aux3 = "00"
    rsdiario!H_Cta_Larga = Trim(Me.cboH_aux1.Text)
    rsdiario!H_Des_Larga = Trim(Me.cboh_aux1_denom.Text)
    rsdiario!H_MontoBs = Val(Me.Txth_Bs)
    rsdiario!H_MontoDl = Val(Me.Txth_dls)
    rsdiario!H_Cambio = Val(Me.Txt_Cambio)
    '''''revisar codigo de usuario
    rsdiario!Usr_Usuario = GlUsuario ' variable global de usuario
    rsdiario!fecha_registro = CDate(Format(Date, "dd/mm/yyyy"))
    rsdiario!hora_registro = Format(Time, "hh:mm:ss")
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
Set rsRepCab = New ADODB.Recordset
Set rsRepDet = New ADODB.Recordset

'With dtetraspasos
' ******** Imprime
'    Set rsRepCab = New ADODB.Recordset
    If rsRepCab.State = 1 Then rsRepCab.Close
     rsRepCab.Open "SELECT Co_comprobante_rep.* FROM Co_comprobante_rep", db, adOpenKeyset, adLockOptimistic
    '"select * from co_comprobante_rep ", db, adOpenKeyset, adLockOptimistic
        While Not rsRepCab.EOF And rsRepCab.RecordCount > 0
                rsRepCab.Delete
                rsRepCab.MoveNext
        Wend
        '"select * from co_comprobante_rep ", db, adOpenKeyset, adLockOptimistic
        rsRepCab.AddNew
        rsRepCab("cod_comp") = rscomprobante1!Cod_Comp
        rsRepCab("tipo_comp") = rscomprobante1!tipo_comp
        rsRepCab("ges_gestion") = "2000"  ' ERRRRRRRRRRRRRRRRRRRRRRRRRR
        If Not IsNull(rsComprobante!Cod_Trans) Then
            rsRepCab("cod_trans") = rscomprobante1!Cod_Trans
          Else
            rsRepCab("cod_trans") = "-"
        End If
        rsRepCab("Num_respaldo") = rscomprobante1!Num_Respaldo
        rsRepCab("fecha_A") = CDate(rscomprobante1!Fecha_A)
        If rscomprobante1!Codigo_Beneficiario <> "" Then
            rsRepCab("codigo_beneficiario") = rscomprobante1!Codigo_Beneficiario
        Else
            rsRepCab("codigo_beneficiario") = "-"
        End If
        rsRepCab("codigo_documento") = rscomprobante1!codigo_documento
        rsRepCab("glosa") = rscomprobante1!Glosa
        If rsComprobante("status") = "S" Then
            rsComprobante("status") = "A"
           Else
            rsRepCab("status") = "S"
        End If
        rsRepCab.Update
       'Set rsRepDet = New ADODB.Recordset
       If rsRepDet.State = 1 Then rsRepDet.Close
       rsRepDet.Open "SELECT co_diario_rep.* FROM co_diario_rep ", db, adOpenKeyset, adLockOptimistic
       'rsRepDet.Open "select * from co_diario_rep ", db, adOpenKeyset, adLockOptimistic
       While Not rsRepDet.EOF And rsRepDet.RecordCount > 0
                    rsRepDet.Delete
                    rsRepDet.MoveNext
       Wend
       
        If rsRepDet.State = 1 Then rsRepDet.Close
        rsRepDet.Open
        '"select * from co_diario_rep ", db, adOpenKeyset, adLockOptimistic
        If Not IsNull(rsRepCab("cod_comp")) And Not IsNull(rsRepDet("tipo_comp")) Then
'           Set rsPlanCta = New ADODB.Recordset
'           rsPlanCta.Open "select * from co_diario where cod_comp=" & rsRepCab("cod_comp") & " and tipo_comp='" & rsRepCab("tipo_comp") & "'", db, adOpenKeyset, adLockOptimistic
           
'           Set rsDetalle = New ADODB.Recordset
'
 '          .rscomprobante.Open
           '"select * from co_diario where cod_comp=" & rsRepCab("cod_comp") & " and tipo_comp='" & rsRepCab("tipo_comp") & "'", db, adOpenKeyset, adLockOptimistic
           
            'Set DtGDetalle.DataSource = rsDetalle
'            If .rscomprobante.State = 1 Then .rscomprobante.Close
'            .rscomprobante.Open
          'If .rscomprobante.RecordCount > 0 Then
           Set rsnombre_cta = New ADODB.Recordset
           
          'While Not .rscomprobante.EOF
            rsRepDet.AddNew
            rsRepDet("cod_comp") = rscomprobante1!Cod_Comp
            rsRepDet("tipo_comp") = rscomprobante1!tipo_comp
            rsRepDet("d_cuenta") = rscomprobante1!d_cuenta
            rsRepDet("d_subcta1") = rscomprobante1!d_subcta1
            rsRepDet("D_subcta2") = rscomprobante1!d_subcta2
            If rsnombre_cta.State = 1 Then rsnombre_cta.Close
            rsnombre_cta.Open "SELECT NombreCta From CC_Plan_Cuentas WHERE (SubCta1 = '" & Trim(rscomprobante1!d_subcta1) & _
            "') AND (SubCta2 = '" & Trim(rscomprobante1!d_subcta2) & "') AND  (Aux1 = '" & Trim(rscomprobante1!d_Aux1) & _
            "') AND (Aux2 = '" & Trim(rscomprobante1!d_Aux2) & "') AND (Aux3 = '" & Trim(rscomprobante1!d_Aux3) & "') AND  (Cuenta = '" & _
            Trim(rscomprobante1!d_cuenta) & "')", db, adOpenKeyset, adLockReadOnly
''            rsnombre_cta.Open "SELECT CC_Plan_Cuentas.NombreCta From CC_Plan_Cuentas  where " & _
'            "(CC_Plan_Cuentas.Cuenta)= '" & rscomprobante1!d_cuenta & "' AND (CC_Plan_Cuentas.SubCta1)= '" & _
'            rscomprobante1!d_subcta1 & "' and  CC_Plan_Cuentas.SubCta2 ='" & rscomprobante1!d_subcta2 & _
'            " ' AND (CC_Plan_Cuentas.Aux1)= '" & rscomprobante1!d_aux1 & "' AND (CC_Plan_Cuentas.Aux2)= ' " & rscomprobante1!d_aux2 & "' AND (CC_Plan_Cuentas.Aux3)= '" & _
'            rscomprobante1!d_aux3 & "'", db, adOpenKeyset, adLockReadOnly
            MsgBox rsnombre_cta.RecordCount
            '.nombre_cta .rscomprobante1!d_cuenta, .rscomprobante1!d_subcta1, .rscomprobante1!d_subcta2, .rscomprobante1!d_aux1, .rscomprobante1!d_aux2, .rscomprobante1!d_aux3
            
            rsRepDet("d_nombre") = rsnombre_cta!NombreCta
            rsRepDet("d_Aux1") = rscomprobante1("d_Aux1")
            rsRepDet("d_Aux2") = rscomprobante1("d_Aux2")
            rsRepDet("d_Aux3") = rscomprobante1("d_Aux3")
            rsRepDet("D_Cta_larga") = rscomprobante1("D_Cta_larga")
            rsRepDet("D_MontoBs") = rscomprobante1("D_MontoBs")
            rsRepDet("D_MontoDl") = rscomprobante1("D_MontoDl")
            rsRepDet("D_Cambio") = rscomprobante1("D_Cambio")
            rsRepDet("D_Des_Larga") = rscomprobante1("D_Des_Larga")
            rsRepDet("H_cuenta") = rscomprobante1("H_cuenta")
            rsRepDet("H_subcta1") = rscomprobante1("H_subcta1")
            rsRepDet("H_subcta2") = rscomprobante1("H_subcta2")
            If rsnombre_cta.State = 1 Then rsnombre_cta.Close
            
             rsnombre_cta.Open "SELECT NombreCta From CC_Plan_Cuentas WHERE (SubCta1 = '" & Trim(rscomprobante1!h_subcta1) & _
            "') AND (SubCta2 = '" & Trim(rscomprobante1!h_subcta2) & "') AND  (Aux1 = '" & Trim(rscomprobante1!h_Aux1) & _
            "') AND (Aux2 = '" & Trim(rscomprobante1!h_Aux2) & "') AND (Aux3 = '" & Trim(rscomprobante1!h_Aux3) & "') AND  (Cuenta = '" & _
            Trim(rscomprobante1!h_cuenta) & "')", db, adOpenKeyset, adLockReadOnly
            rsRepDet("h_nombre") = rsnombre_cta!NombreCta
         
'
            rsRepDet("H_Aux1") = rscomprobante1("H_Aux1")
            rsRepDet("H_Aux2") = rscomprobante1("H_Aux2")
            rsRepDet("H_Aux3") = rscomprobante1("H_Aux3")
            If rscomprobante1("H_Cta_larga") <> "" Then
                rsRepDet("H_Cta_larga") = rscomprobante1("H_Cta_larga")
              Else
                rsRepDet("H_Cta_larga") = "-"
            End If
            rsRepDet("H_MontoBs") = rscomprobante1("H_MontoBs")
            LiteralCry = Str(Int(rscomprobante1("H_MontoBs")))
            Decimal2 = Str(Round((rscomprobante1("H_MontoBs") - Val(LiteralCry)), 2) * 100)
            rsRepDet("H_MontoDl") = rscomprobante1("H_MontoDl")
            rsRepDet("H_Cambio") = rscomprobante1("H_Cambio")
            rsRepDet("H_Des_Larga") = rscomprobante1("H_Des_Larga")
            rsRepDet("literal") = Literal(LiteralCry) + " " + Decimal2 + "/100  Bolivianos"
            'rsRepDet("literal") = Literal(LiteralCry) + " Bolivianos"
            
            rsRepDet.Update
            '.rs.MoveNext
            
          'Wend
          'End If
        End If
    
 RepComprob_Conta.Show
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
  Me.Txt_Fecha = Date
  Me.Txt_ges = Year(Date)
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
If rscomprobante1.State = 1 Then rsComprobante.Close
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
    Me.D1documento.Text = rscomprobante1!codigo_documento
    Me.Txt_Respaldo = rscomprobante1!Num_Respaldo
  '''Me.cbobeneficiario.Text = .rscomprobante!codigo_beneficiario
    Me.d1beneficiario.Text = rscomprobante1!Codigo_Beneficiario
    Me.Txt_glosa = rscomprobante1!Glosa
    Me.lblh_cuenta = rscomprobante1!h_cuenta
    Me.lblh_sub1 = rscomprobante1!h_subcta1
    Me.cboH_sub2 = rscomprobante1!h_subcta2
    Me.Txt_Cambio = Val(rscomprobante1!H_Cambio)
    Me.Txth_Bs = Val(rscomprobante1!D_MontoBs)
    Me.Txth_dls = Val(rscomprobante1!H_MontoDl)
    Me.cboH_aux1.Text = rscomprobante1!H_Cta_Larga
    Me.cboh_aux1_denom.Text = rscomprobante1!H_Des_Larga
    Me.lbld_cuenta = rscomprobante1!d_cuenta
    Me.lbld_sub1 = rscomprobante1!d_subcta1
    Me.cbod_sub2 = rscomprobante1!d_subcta2
    Me.TxtD_Bs = Val(rscomprobante1!D_MontoBs)
    Me.TxtD_Dls = Val(rscomprobante1!D_MontoDl)
    Me.cbod_aux1.Text = rscomprobante1!D_Cta_Larga
    Me.cbod_aux1_denom.Text = rscomprobante1!D_Des_Larga
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
'*************recordset para el grid
Set rsComprobante = New ADODB.Recordset
If rsComprobante.State = 1 Then rsComprobante.Close
rsComprobante.Open "SELECT Co_Comprobante_M.Cod_Comp, " & _
"Co_Comprobante_M.Tipo_Comp, Co_Comprobante_M.codigo_beneficiario," & _
"Co_Comprobante_M.status,Co_Comprobante_M.cod_trans,Co_Comprobante_M.org_codigo " & _
"FROM CO_Diario INNER JOIN  Co_Comprobante_M ON CO_Diario.Cod_Comp = Co_Comprobante_M.Cod_Comp " & _
"AND  CO_Diario.Tipo_Comp = Co_Comprobante_M.Tipo_Comp " & _
"WHERE (CO_Diario.D_Cuenta = '1111') AND (CO_Diario.D_Subcta1 = '02') AND " & _
"(CO_Diario.H_Cuenta = '1111') AND (CO_Diario.H_SubCta1 = '02') AND " & _
"(Co_Comprobante_M.Tipo_Comp = 'PCE' OR Co_Comprobante_M.Tipo_Comp = 'TRP' ) AND (Co_Comprobante_M.status <> 'E')" & _
"ORDER BY Co_Comprobante_M.Cod_Comp", db, adOpenDynamic, adLockOptimistic
Set Me.DtGrid_comprobante.DataSource = rsComprobante

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

'rscomprobante.Open "SELECT  Co_Comprobante_M.Cod_Comp, Co_Comprobante_M.cod_trans," & _
'"Co_Comprobante_M.cod_trans_detalle, Co_Comprobante_M.codigo_beneficiario, Co_Comprobante_M.codigo_documento," & _
'"Co_Comprobante_M.Fecha_A, Co_Comprobante_M.Fecha_Registro, Co_Comprobante_M.ges_gestion," & _
'"Co_Comprobante_M.Glosa, Co_Comprobante_M.Hora_Registro, Co_Comprobante_M.Num_Respaldo," & _
'"Co_Comprobante_M.org_codigo, Co_Comprobante_M.status, Co_Comprobante_M.Tipo_Comp," & _
'"Co_Comprobante_M.Usr_Usuario, CO_Diario.D_Aux1, CO_Diario.D_Aux2, CO_Diario.D_Aux3," & _
'"CO_Diario.D_Cambio, CO_Diario.D_Cta_Larga, CO_Diario.D_Cuenta, CO_Diario.D_Des_Larga," & _
'"CO_Diario.D_MontoBs, CO_Diario.D_MontoDl, CO_Diario.D_Nombre, CO_Diario.D_Subcta1," & _
'"CO_Diario.D_SubCta2, CO_Diario.H_Aux1, CO_Diario.H_Aux2, CO_Diario.H_Aux3," & _
'"CO_Diario.H_Cambio, CO_Diario.H_Cta_Larga, CO_Diario.H_Cuenta, CO_Diario.H_Des_Larga," & _
'"CO_Diario.H_MontoBs, CO_Diario.H_MontoDl, CO_Diario.H_Nombre, CO_Diario.H_SubCta1," & _
'"CO_Diario.H_SubCta2 FROM Co_Comprobante_M, CO_Diario WHERE Co_Comprobante_M.Cod_Comp = CO_Diario.Cod_Comp " & _
'"AND Co_Comprobante_M.Tipo_Comp = CO_Diario.Tipo_Comp AND (Co_Comprobante_M.Tipo_Comp = 'PCE')" & _
'"AND (CO_Diario.D_Cuenta = '1111') AND (CO_Diario.D_Subcta1 = '02') AND " & _
'"CO_Diario.H_Cuenta = '1111') AND (CO_Diario.H_SubCta1 = '02') AND " & _
'"(Co_Comprobante_M.Tipo_Comp = 'PCE') AND (Co_Comprobante_M.status <> 'E') ORDER BY Co_Comprobante_M.Cod_Comp", db, adOpenKeyset, adLockOptimistic


'**********
'cod_usr = "gab001"
	Call SeguridadSet(Me)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
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
Set rscorrelativo = New ADODB.Recordset
If rscorrelativo.State = 1 Then rscorrelativo.Close
  rscorrelativo.Open "SELECT numero_correlativo, tipo_tramite FROM fc_correl WHERE (tipo_tramite = 'cmbte')", db, adOpenKeyset, adLockOptimistic
  rscorrelativo.MoveFirst
  num_comprobante = rscorrelativo!numero_correlativo + 1
  rscorrelativo!numero_correlativo = rscorrelativo!numero_correlativo + 1
  rscorrelativo.Update
'End With
End Sub


VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSoesMain 
   Caption         =   "Solicitud de Desembolso"
   ClientHeight    =   7335
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frmSoesMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Impresión de Listados de Comprobantes "
      Height          =   540
      Left            =   135
      TabIndex        =   67
      Top             =   -15
      Width           =   3240
      Begin VB.CommandButton cmd_prn_lista_comp 
         Caption         =   "Listar"
         Height          =   285
         Left            =   2535
         TabIndex        =   70
         Top             =   195
         Width           =   630
      End
      Begin VB.OptionButton op_comp_env 
         Caption         =   "Enviados"
         Height          =   195
         Left            =   105
         TabIndex        =   69
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton op_comp_no_env 
         Caption         =   "No Enviados"
         Height          =   195
         Left            =   1260
         TabIndex        =   68
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.TextBox txtsoc_nro_sol 
      BackColor       =   &H80000018&
      DataField       =   "soc_nro_sol"
      Enabled         =   0   'False
      Height          =   285
      Left            =   8385
      TabIndex        =   66
      Top             =   240
      Width           =   660
   End
   Begin VB.Frame fra_ver_lista 
      Caption         =   "Ver Lista de Solicitudes ..."
      Height          =   810
      Left            =   135
      TabIndex        =   57
      Top             =   540
      Width           =   3240
      Begin VB.OptionButton opt_no_confirmados 
         Caption         =   "No Confirmados"
         Height          =   195
         Left            =   105
         TabIndex        =   59
         Top             =   495
         Width           =   1440
      End
      Begin VB.OptionButton opt_confirmados 
         Caption         =   "Confirmados"
         Height          =   195
         Left            =   105
         TabIndex        =   58
         Top             =   240
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   3540
      TabIndex        =   52
      Top             =   6540
      Width           =   7200
      Begin VB.CommandButton cmd_Insert_Cab 
         Caption         =   "Insertar Item"
         Height          =   390
         Left            =   3330
         TabIndex        =   55
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmd_update_cab 
         Caption         =   "Ver Detalle"
         Height          =   390
         Left            =   4575
         TabIndex        =   54
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmd_del_cab 
         Caption         =   "Borrar Item"
         Height          =   390
         Left            =   5835
         TabIndex        =   53
         Top             =   240
         Width           =   1245
      End
      Begin MSAdodcLib.Adodc adoSoes 
         Height          =   330
         Left            =   135
         Top             =   225
         Width           =   2220
         _ExtentX        =   3916
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
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   3540
      TabIndex        =   38
      Top             =   3135
      Width           =   7200
      Begin MSDataListLib.DataCombo dcmCtas 
         Bindings        =   "frmSoesMain.frx":324A
         Height          =   315
         Left            =   2535
         TabIndex        =   64
         Top             =   1410
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "cta"
         BoundColumn     =   "cta"
         Text            =   "DataCombo1"
      End
      Begin VB.ComboBox txtsoc_tipo_mone_sol 
         Height          =   315
         ItemData        =   "frmSoesMain.frx":3261
         Left            =   2535
         List            =   "frmSoesMain.frx":326B
         TabIndex        =   43
         Top             =   180
         Width           =   750
      End
      Begin VB.CommandButton cmd_banco 
         Caption         =   "..."
         Height          =   285
         Left            =   6735
         TabIndex        =   41
         Top             =   780
         Width           =   360
      End
      Begin VB.CommandButton cmd_banco2 
         Caption         =   "..."
         Height          =   330
         Left            =   6750
         TabIndex        =   39
         Top             =   1095
         Width           =   360
      End
      Begin MSAdodcLib.Adodc ado_bancos1 
         Height          =   330
         Left            =   3480
         Top             =   750
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin MSDataListLib.DataCombo dcmBancos1 
         Bindings        =   "frmSoesMain.frx":3279
         Height          =   315
         Left            =   2535
         TabIndex        =   40
         Top             =   780
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "bco_descripcion_larga"
         BoundColumn     =   "bco_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSMask.MaskEdBox txtsoc_mon_mone_sol 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   5745
         TabIndex        =   42
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc ado_bancos2 
         Height          =   330
         Left            =   3255
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
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
      Begin MSDataListLib.DataCombo dcmBancos2 
         Bindings        =   "frmSoesMain.frx":3293
         Height          =   315
         Left            =   2535
         TabIndex        =   44
         Top             =   1095
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "bco_descripcion_larga"
         BoundColumn     =   "bco_codigo"
         Text            =   "Todos"
      End
      Begin MSMask.MaskEdBox txtsoc_monto_us 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   5745
         TabIndex        =   45
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc ado_Ctas 
         Height          =   330
         Left            =   5850
         Top             =   1455
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Moneda Solicitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   525
         TabIndex        =   51
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto Solicitado en BOB:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3630
         TabIndex        =   50
         Top             =   210
         Width           =   2085
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto Solicitado en USD:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3615
         TabIndex        =   49
         Top             =   510
         Width           =   2100
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco Intermediario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   735
         TabIndex        =   48
         Top             =   840
         Width           =   1740
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco Depositario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   945
         TabIndex        =   47
         Top             =   1125
         Width           =   1530
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. Cta. Bancaria Beneficiario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   75
         TabIndex        =   46
         Top             =   1455
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   3540
      TabIndex        =   30
      Top             =   1050
      Width           =   7200
      Begin VB.TextBox txtsoc_nro_ref 
         DataField       =   "soc_nro_ref"
         Height          =   285
         Left            =   2265
         TabIndex        =   33
         Top             =   180
         Width           =   2010
      End
      Begin VB.TextBox txtsoc_cod_ent_eje 
         DataField       =   "soc_cod_ent_eje"
         Height          =   285
         Left            =   300
         TabIndex        =   32
         Text            =   "MECYD"
         Top             =   450
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txt_Unidad 
         DataField       =   "soc_cod_ent_eje"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2265
         TabIndex        =   31
         Text            =   "MINISTERIO DE EDUCACION, CULTURA Y DEPORTES"
         Top             =   480
         Width           =   4710
      End
      Begin MSComCtl2.DTPicker dtp_Fecha 
         Height          =   255
         Left            =   5745
         TabIndex        =   34
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23461889
         CurrentDate     =   36656
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5100
         TabIndex        =   37
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No de Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   36
         Top             =   225
         Width           =   1575
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Entidad Ejecutora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   645
         TabIndex        =   35
         Top             =   510
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   3540
      TabIndex        =   27
      Top             =   480
      Width           =   7200
      Begin VB.TextBox txtsoc_nro_sol_def 
         BackColor       =   &H80000018&
         DataField       =   "soc_nro_sol"
         Height          =   285
         Left            =   6375
         TabIndex        =   65
         Top             =   195
         Width           =   660
      End
      Begin MSDataListLib.DataCombo cb_codigo_convenio 
         Bindings        =   "frmSoesMain.frx":32AD
         Height          =   315
         Left            =   2475
         TabIndex        =   63
         Top             =   165
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo_convenio"
         BoundColumn     =   "codigo_convenio"
         Text            =   "DataCombo1"
      End
      Begin MSAdodcLib.Adodc ado_fc_convenios 
         Height          =   330
         Left            =   3525
         Top             =   165
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N. de Solicitud:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5055
         TabIndex        =   29
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. de operacion del BID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   28
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   225
      Left            =   255
      TabIndex        =   25
      Top             =   7050
      Visible         =   0   'False
      Width           =   930
   End
   Begin MSDataGridLib.DataGrid dgTodo 
      Bindings        =   "frmSoesMain.frx":32CC
      Height          =   5325
      Left            =   1350
      TabIndex        =   24
      Top             =   1380
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   9393
      _Version        =   393216
      BackColor       =   -2147483624
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "soc_codigo_convenio"
         Caption         =   "Convenio"
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
         DataField       =   "soc_nro_sol"
         Caption         =   "Nro. Sol."
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
         DataField       =   "soc_fec_elab"
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
      BeginProperty Column03 
         DataField       =   "soc_monto_us"
         Caption         =   "Monto US."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
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
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgSoes 
      Bindings        =   "frmSoesMain.frx":32E2
      Height          =   1515
      Left            =   3555
      TabIndex        =   19
      Top             =   5025
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   2672
      _Version        =   393216
      BackColor       =   -2147483624
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
         DataField       =   "soe_nro_sec"
         Caption         =   "Sec."
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
         DataField       =   "soe_codigo_categoria"
         Caption         =   "Nro. COI"
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
         DataField       =   "denominacion_categoria"
         Caption         =   "Descripción"
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
         DataField       =   "soe_pais_origen"
         Caption         =   "Pais Origen"
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
         DataField       =   "soe_monto_sol_us"
         Caption         =   "Monto Us."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "soe_monto_sol_bs"
         Caption         =   "Monto Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollGroup     =   2
         BeginProperty Column00 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
   Begin VB.Frame fr_Presentacion 
      Caption         =   "SOLICITAMOS/PRESENTAMOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   3540
      TabIndex        =   14
      Top             =   2010
      Width           =   7200
      Begin VB.OptionButton opt_pre_justificacion 
         Caption         =   "Justificación del Fondo Rotatorio"
         Height          =   225
         Left            =   3390
         TabIndex        =   26
         Top             =   585
         Width           =   2760
      End
      Begin VB.OptionButton opt_sol_reposicion 
         Caption         =   "Reposición del Fondo Rotatorio"
         Height          =   225
         Left            =   3405
         TabIndex        =   23
         Top             =   285
         Width           =   2925
      End
      Begin VB.OptionButton opt_sol_desembolso 
         Caption         =   "Desembolos del Fondo Rotatorio"
         Height          =   225
         Left            =   135
         TabIndex        =   22
         Top             =   765
         Width           =   3270
      End
      Begin VB.OptionButton opt_sol_pago 
         Caption         =   "Pago directo al proveedor o contratista"
         Height          =   225
         Left            =   135
         TabIndex        =   21
         Top             =   525
         Width           =   3270
      End
      Begin VB.OptionButton opt_sol_reembolso 
         Caption         =   "Reembolso de pagos Efectuados"
         Height          =   225
         Left            =   150
         TabIndex        =   20
         Top             =   255
         Width           =   3270
      End
   End
   Begin MSAdodcLib.Adodc adoTodo 
      Height          =   330
      Left            =   1380
      Top             =   6750
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.TextBox txtsoc_usuario 
      DataField       =   "soc_usuario"
      Height          =   285
      Left            =   9870
      TabIndex        =   13
      Top             =   6255
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtsoc_fecha_reg 
      DataField       =   "soc_fecha_reg"
      Height          =   285
      Left            =   9870
      TabIndex        =   11
      Top             =   5865
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtsoc_pre_justificacion 
      DataField       =   "soc_pre_justificacion"
      Height          =   285
      Left            =   10995
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtsoc_sol_reposicion 
      DataField       =   "soc_sol_reposicion"
      Height          =   285
      Left            =   10995
      TabIndex        =   7
      Top             =   4305
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtsoc_sol_desembolso 
      DataField       =   "soc_sol_desembolso"
      Height          =   285
      Left            =   10995
      TabIndex        =   6
      Top             =   3915
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtsoc_sol_pago 
      DataField       =   "soc_sol_pago"
      Height          =   285
      Left            =   10995
      TabIndex        =   5
      Top             =   3540
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox txtsoc_sol_reembolso 
      DataField       =   "soc_sol_reembolso"
      Height          =   285
      Left            =   10995
      TabIndex        =   4
      Top             =   3165
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Frame FraOpciones2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5400
      Left            =   60
      TabIndex        =   1
      Top             =   1350
      Width           =   1215
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   720
         Left            =   105
         Picture         =   "frmSoesMain.frx":32F8
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Permite Confirmar el Registro Acutal"
         Top             =   4575
         Width           =   1005
      End
      Begin VB.CommandButton cdmAnular 
         Caption         =   "Elimina"
         Height          =   720
         Left            =   120
         Picture         =   "frmSoesMain.frx":3602
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Elimina el Registro Actual"
         Top             =   3825
         Width           =   1005
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   135
         Picture         =   "frmSoesMain.frx":3CEC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprime el Formulario correspondiente al Registro Actual"
         Top             =   3090
         Width           =   1005
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "Adicionar"
         Height          =   720
         Left            =   120
         Picture         =   "frmSoesMain.frx":43D6
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Adiciona un Nuevo Registro"
         Top             =   2355
         Width           =   1005
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   720
         Left            =   135
         Picture         =   "frmSoesMain.frx":46E0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Modifica el Registro Actual"
         Top             =   1605
         Width           =   1005
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "frmSoesMain.frx":48EA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Graba las Modificaciones Realizadas al Registro Actual"
         Top             =   150
         Width           =   1005
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   720
         Left            =   120
         MousePointer    =   4  'Icon
         Picture         =   "frmSoesMain.frx":4BF4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancela las Modificaciones realizadas al Registro Actual"
         Top             =   885
         Width           =   1005
      End
   End
   Begin Crystal.CrystalReport CryReporte 
      Left            =   3720
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      Caption         =   "Estado de Registro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3555
      TabIndex        =   61
      Top             =   45
      Width           =   1785
   End
   Begin VB.Label lb_estado_registro 
      Caption         =   "Normal"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   5325
      TabIndex        =   60
      Top             =   60
      Width           =   810
   End
   Begin VB.Label Es_vigente 
      Alignment       =   1  'Right Justify
      Caption         =   "Es_vigente"
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
      Height          =   240
      Left            =   8280
      TabIndex        =   56
      Top             =   45
      Width           =   2430
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "soc_usuario:"
      Height          =   255
      Index           =   17
      Left            =   7920
      TabIndex        =   12
      Top             =   6270
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "soc_fecha_reg:"
      Height          =   255
      Index           =   16
      Left            =   7920
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "soc_pre_justificacion:"
      Height          =   255
      Index           =   9
      Left            =   9165
      TabIndex        =   8
      Top             =   5250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SOLICITUD DE DESEMBOLSO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5460
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmSoesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscomprobante As ADODB.Recordset
Dim rsDetComp     As ADODB.Recordset
Dim accion As String, Cancelar As Boolean
Dim TipoListaComp As String
Public frmSoesMain_ret As String

Public Sub frmSoesMain_procesar(proceso As String)
  Cancelar = True
  accion = Trim(proceso)
  frmSoesMain_ret = ""
  llenaAdoBancos1
  llenaAdoBancos2
  Set tFc_convenios = New ADODB.Recordset
  If tFc_convenios.State = 1 Then tFc_convenios.Close
    tFc_convenios.Open "SELECT codigo_convenio, codigo_convenio FROM fc_convenios WHERE emision_soes = 'S' ", db, adOpenDynamic, adLockReadOnly
  Set ado_fc_convenios.Recordset = tFc_convenios
  habilitaCasillas False
  TipoListaComp = "SELECT_CONF"
  soa_cab_refresca TipoListaComp
  lb_estado_registro.Caption = "Normal"
  If accion = "ABM_SOES" Then
    Caption = "Solicitud de Desembolso"
    frmDetalleSoes.cmd_cap_rechazado.Enabled = False
  ElseIf accion = "DEDUCCIONES" Then
    Caption = "Planilla de Deducciones"
    frmDetalleSoes.cmd_cap_rechazado.Enabled = True
  End If
  Show vbModal
End Sub

Private Sub cb_codigo_convenio_Change()
  llenaCtas (cb_codigo_convenio.Text)
End Sub

Private Sub cdmAnular_Click()
  If adoTodo.Recordset!soc_vigente = "T" Then
    If MsgBox("Esta seguro de Borrar este Registro?", vbYesNo) = vbYes Then
      If Not (adoTodo.Recordset.EOF Or adoTodo.Recordset.BOF) Then
        'borra el registro actual
        delete_soes_cab Me.adoTodo.Recordset!soc_nro_sol, Me.adoTodo.Recordset!soc_codigo_convenio
        'refresca el dbgrid del SOES_CAB y el de SOES
        soa_cab_refresca TipoListaComp
        soa_refresca False
      End If
    End If
  Else
    MsgBox "Este Formulario ya fue confirmado o tiene Comprobantes rechazados"
  End If
End Sub

Private Sub cmd_banco_Click()
  If opt_sol_pago.Value = True Then
    'llama al form BANCOS en modo INSERT que devuelve el codigo del banco que se inserto
    frmFc_bancos.fc_bancos_procesar "INSERT", ""
    If frmFc_bancos.bco_codigo_ret <> "" Then
      llenaAdoBancos1
      llenaAdoBancos2
      dcmBancos1.BoundText = frmFc_bancos.bco_codigo_ret
    End If
  End If
End Sub

Private Sub cmd_banco2_Click()
  If opt_sol_pago.Value = True Then
    'llama al form BANCOS en modo INSERT que devuelve el codigo del banco que se inserto
    frmFc_bancos.fc_bancos_procesar "INSERT", ""
    If frmFc_bancos.bco_codigo_ret <> "" Then
      llenaAdoBancos1
      llenaAdoBancos2
      dcmBancos2.BoundText = frmFc_bancos.bco_codigo_ret
    End If
  End If
End Sub

Private Sub cmd_prn_lista_comp_Click()
Dim monto As String, iResult As Integer
  If Me.op_comp_env.Value = True Then
    Call frmDetalleSoes.Rep002("COMP_ENVIADOS_SOES", "\rep_lista_comp.rpt", "Lista de Comprobantes enviados")
  Else
    Call frmDetalleSoes.Rep002("COMP_NO_ENV_SOES", "\rep_lista_comp_soes.rpt", "Lista de Comprobantes que faltan enviar")
  End If
End Sub

Private Sub cmdAdicionar_Click()
  lb_estado_registro.Caption = "Insertado"
  Call set_habilita_boton_todo(0, opt_confirmados.Value, True, lb_estado_registro.Caption)
  frmSoesMain.Caption = "Nueva Solicitud de Reemboloso"
  ResetSoes_cab 0 'limpia casillas de SOES_CAB
  soa_refresca True
  habilitaCasillas True
  'Datos.conexion.BeginTrans
End Sub

Private Sub CmdCancelar_Click()
  If MsgBox("Esta seguro de Cancelar?", vbYesNo) = vbYes Then
    lb_estado_registro.Caption = "Normal"
    habilitaCasillas False
    frmSoesMain.Caption = "Solicitud de Reemboloso"
    'Datos.conexion.RollbackTrans
    soa_cab_refresca TipoListaComp
  End If
End Sub

Private Sub cmdConfirmar_Click()
  If adoTodo.Recordset!soc_vigente = "T" Then
    If MsgBox("Atencion. Una vez Confirmado este registro no podra ser modificado. Desea Confirmarlo?", vbYesNo) = vbYes Then
      'Datos.conexion.BeginTrans
      db.Execute "update soes_cab set soc_vigente= 'N' where soc_nro_sol = " & Me.adoTodo.Recordset!soc_nro_sol & " and soc_codigo_convenio = '" & Me.adoTodo.Recordset!soc_codigo_convenio & "'"
      'Datos.conexion.CommitTrans
    End If
  Else
    MsgBox "Este Formulario ya fue confirmado o tiene Comprobantes rechazados"
  End If
End Sub

Private Sub cmdImprimir_Click()
Dim monto As String, iResult As Integer
  If Me.txtsoc_tipo_mone_sol.Text = "USD" Then
    monto = Me.txtsoc_monto_us.Text
  Else
    monto = Me.txtsoc_mon_mone_sol.Text
  End If
  '415 BANCO MUNDIAL 411 BID
  If GetValorGeneral("select org_codigo as retorno from fc_convenios where codigo_convenio = '" & Me.adoTodo.Recordset!soc_codigo_convenio & "'") = "411" Then
    frmDetalleSoes.Rep001 "REP_CATEGORIA", "\rep_bid_ctrl_desembolso.rpt", "", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, Me.txtsoc_nro_sol.Text, literal(monto)
    frmDetalleSoes.Rep001 "REP_CATEGORIA", "\rep_bid_form2_soes_v1.rpt", "", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, 0, ""
    frmDetalleSoes.Rep001 "REP_SOES_FORM", "\rep_bid_form1_soes_v1.rpt", "FORMULARIO DE ....", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, Me.txtsoc_nro_sol.Text, literal(monto)
  Else
    frmDetalleSoes.Rep001 "REP_CATEGORIA", "\rep_bm_ctrl_desembolso.rpt", "", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, Me.txtsoc_nro_sol.Text, literal(monto)
    frmDetalleSoes.Rep001 "REP_CATEGORIA", "\rep_bm_form2_soes_v1.rpt", "", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, 0, ""
    frmDetalleSoes.Rep001 "REP_SOES_FORM", "\rep_bm_form1_soes_v1.rpt", "FORMULARIO DE ....", Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, Me.txtsoc_nro_sol.Text, literal(monto)
  End If
End Sub

Private Sub cmdModificar_Click()
  If adoTodo.Recordset!soc_vigente = "T" Then
    lb_estado_registro.Caption = "Modificado"
    Call set_habilita_boton_todo(0, opt_confirmados.Value, True, lb_estado_registro.Caption)
    Caption = "Solicitud de Reemboloso"
    habilitaCasillas True
    'Datos.conexion.BeginTrans
  Else
    MsgBox "Este Formulario ya fue confirmado o tiene Comprobantes rechazados"
  End If
End Sub

Private Sub Command1_Click()
Dim tc As Double, cod As String
'MsgBox GetValor("fc_convenios", "denominacion_convenio", "codigo_convenio", "2650-BO")
'MsgBox GetValor("fc_convenios", "count(*)", "codigo_convenio", "2650-BO")
'MsgBox GetValor("fc_convenios", "cta_codigo_bcb", "codigo_convenio", cb_codigo_convenio.Text)
'GetTc2 "111", 140, "1-297809", tc, cod
'MsgBox tc & " " & cod
'MsgBox adoSoes.Recordset.EOF
'Set dgSoes.DataSource = Nothing
'  adoSoes.Recordset.Open
  adoSoes.Recordset.Close
End Sub

Private Sub llenaAdoBancos1()
Dim fecha As Date
  Datos.dbo_so_fc_bancos "SELECT", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", fecha, "", "", ""
  With Datos.rsdbo_so_fc_bancos
    Set ado_bancos1.Recordset = .Clone
    .Close
  End With
End Sub

Private Sub llenaAdoBancos2()
Dim fecha As Date
  Datos.dbo_so_fc_bancos "SELECT", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", fecha, "", "", ""
  With Datos.rsdbo_so_fc_bancos
    Set ado_bancos2.Recordset = .Clone
    .Close
  End With
End Sub

Public Sub adoSoes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'MsgBox "fila " + Str(adoSoes.Recordset.AbsolutePosition) ' + " nro " + adoTodo.Recordset!soc_nro_sol
  If Not (adoSoes.Recordset.EOF Or adoSoes.Recordset.BOF) Then
    adoSoes.Caption = CStr(adoSoes.Recordset.Bookmark) & " de " & CStr(adoSoes.Recordset.RecordCount)
  Else
    adoSoes.Caption = "0 de 0"
  End If
End Sub

Private Sub adoTodo_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'MsgBox "fila " + Str(adoTodo.Recordset.AbsolutePosition) ' + " nro " + adoTodo.Recordset!soc_nro_sol
  If Not (adoTodo.Recordset.BOF Or adoTodo.Recordset.EOF) Then
    adoTodo.Caption = CStr(adoTodo.Recordset.Bookmark) & " de " & CStr(adoTodo.Recordset.RecordCount)
    llenaSoes_cab
    soa_refresca False
    Call set_habilita_boton_todo(adoTodo.Recordset.RecordCount, opt_confirmados.Value, True, lb_estado_registro.Caption)
  Else
    ResetSoes_cab 0
    adoTodo.Caption = " 0 de 0"
    soa_refresca True
    Call set_habilita_boton_todo(0, opt_confirmados.Value, True, lb_estado_registro.Caption)
  End If
End Sub

Private Sub cmd_del_cab_Click()
  If MsgBox("Esta seguro de Borrar el Item?", vbYesNo) = vbYes Then
    If Not (adoSoes.Recordset.EOF Or adoSoes.Recordset.BOF) Then
      frmSoesMain.txtsoc_monto_us.Text = Str(Val(frmSoesMain.txtsoc_monto_us.Text) - Me.adoSoes.Recordset!soe_monto_sol_us)
      frmSoesMain.txtsoc_mon_mone_sol.Text = Str(Val(frmSoesMain.txtsoc_mon_mone_sol.Text) - Me.adoSoes.Recordset!soe_monto_sol_bs)
      delete_soes Val(txtsoc_nro_sol.Text), cb_codigo_convenio.Text, Me.adoSoes.Recordset!soe_nro_sec
      soa_refresca False
    End If
  End If
End Sub

Private Sub cmd_Insert_Cab_Click()
Dim convenio As String
  convenio = cb_codigo_convenio.Text
  If valida_reg_soes_cab(False) Then
    'Graba temporalmente hasta que confirme el soes_cab
    If frmSoesMain.Caption = "Nueva Solicitud de Reemboloso" Then
      frmSoesMain.Caption = "Solicitud de Reemboloso. Confirme o Cancele"
      soa_cab_get_max_nro_sol
      insert_soes_cab "INSERT"
      soa_cab_refresca TipoListaComp
      frmSoesMain.adoTodo.Recordset.MoveFirst
      frmSoesMain.adoTodo.Recordset.Find "soc_codigo_convenio = '" & convenio & "'", , adSearchForward
      frmSoesMain.adoTodo.Recordset.Find "soc_nro_sol= '" & Str(gl_nro_sol) & "'", , adSearchForward
      'MsgBox "se inserto la cabecera"
    End If
    frmDetalleSoes.frmDetalleSoes_procesar "INSERT", Me.cb_codigo_convenio.Text, frmSoesMain.adoTodo.Recordset!soc_nro_sol
  End If
  txtsoc_tipo_mone_sol.Enabled = False
End Sub

Private Sub cmd_update_cab_Click()
  If accion = "ABM_SOES" Then
    frmDetalleSoes.cmd_cap_rechazado.Enabled = False
  Else
    frmDetalleSoes.cmd_cap_rechazado.Enabled = True
  End If
  
  If cmd_update_cab.Caption = "Ver Detalle" And accion = "DEDUCCIONES" Then
    frmDetalleSoes.cmd_InsertItem.Enabled = False
    frmDetalleSoes.cdm_borrar_item.Enabled = False
    frmDetalleSoes.CmdGrabar.Enabled = False
  Else
    frmDetalleSoes.cmd_InsertItem.Enabled = True
    frmDetalleSoes.cdm_borrar_item.Enabled = True
    frmDetalleSoes.CmdGrabar.Enabled = True
'    CmdCancelar.Enabled = False
  End If
     
  If Not (adoSoes.Recordset.EOF Or adoSoes.Recordset.BOF) Then
    habilitaCasillas (CmdGrabar.Enabled)
    If cmd_update_cab.Caption = "Ver Detalle" Then
      frmDetalleSoes.frmDetalleSoes_procesar "SELECT", cb_codigo_convenio.Text, Val(Me.txtsoc_nro_sol.Text)
    ElseIf cmd_update_cab.Caption = "Modificar" Then
'      Datos.conexion.BeginTrans  'para deshacer lo modificado
      frmDetalleSoes.frmDetalleSoes_procesar "UPDATE", cb_codigo_convenio.Text, Val(Me.txtsoc_nro_sol.Text)
    End If
  End If
End Sub

Private Sub CmdGrabar_Click()
  If valida_reg_soes_cab(True) Then
    If MsgBox("Esta seguro de Grabar?", vbYesNo) = vbYes Then
      lb_estado_registro.Caption = "Normal"
      habilitaCasillas False
      frmSoesMain.Caption = "Solicitud de Reemboloso"
      insert_soes_cab "UPDATE"
      soa_cab_refresca TipoListaComp
      'Datos.conexion.CommitTrans
    End If
  End If
End Sub

Private Sub CmdSalir_Click()
  Unload frmSoesMain
End Sub
'------------------------------funciones de SOES_CAB ---------------------------------

Public Sub soa_cab_refresca_detalle(nro_sol As Integer, convenio As String)
Dim fecha As Date
  Datos.dbo_so_soes_cab "SELECT_UNO", nro_sol, convenio, fecha, "", "", "", 0, 0, 0, "", "", "", fecha, "", "", ""
  With Datos.rsdbo_so_soes_cab
    Call llenaSoes_cab
    Set adoTodo.Recordset = .Clone
    Me.dgTodo.Refresh
   .Close
  End With
End Sub

Public Sub update_soes_cab()
 Datos.dbo_so_soes_cab "UPDATE", _
  Me.txtsoc_nro_sol, _
  cb_codigo_convenio.Text, _
  Me.dtp_Fecha.Value, _
  Me.txtsoc_nro_ref, _
  Me.txtsoc_cod_ent_eje, _
  Me.txtsoc_sol_reembolso, _
  IIf(Me.txtsoc_tipo_mone_sol.Text = "USD", "Us", "Bs"), _
  Val(Me.txtsoc_mon_mone_sol), _
  Val(Me.txtsoc_monto_us), _
  Me.dcmBancos1.BoundText, _
  Me.dcmBancos2.BoundText, _
  Me.dcmCtas.BoundText, _
  Date, _
  Me.txtsoc_usuario, _
  "T", _
  txtsoc_nro_sol_def.Text
End Sub

Public Sub delete_soes_cab(nro_sol As Integer, convenio As String)
Dim fecha As Date
  Datos.dbo_so_soes_cab "DELETE", nro_sol, convenio, fecha, "", "", "", 0, 0, 0, "", "", "", fecha, "", "", ""
End Sub

Public Sub soa_refresca(nuevo As Boolean)
Dim nro_sol As Integer, codigo_convenio As String
  If Not (Me.adoTodo.Recordset.EOF Or Me.adoTodo.Recordset.BOF) Then
    nro_sol = Me.adoTodo.Recordset!soc_nro_sol
    codigo_convenio = Me.adoTodo.Recordset!soc_codigo_convenio
    If nuevo Then
      nro_sol = -1
    End If
  End If
  Datos.dbo_so_soes "SELECT_HIJOS", nro_sol, codigo_convenio, 0, "", "", 0, "", 0, 0, "", ""
  With Datos.rsdbo_so_soes
    Set adoSoes.Recordset = .Clone
'    Me.dgSoes.Refresh
    .Close
  End With
End Sub

Public Sub soa_cab_refresca(opcion As String)
Dim fecha As Date
 Datos.dbo_so_soes_cab opcion, 0, "", fecha, "", "", "", 0, 0, 0, "", "", "", fecha, "", "", ""
 With Datos.rsdbo_so_soes_cab
   Set adoTodo.Recordset = .Clone
   Me.dgTodo.Refresh
  .Close
 End With
 'soa_refresca False
End Sub

Public Sub delete_soes(nro_sol As Integer, convenio As String, soe_nro_sec As Integer)
Dim fecha As Date
 Datos.dbo_so_soes "DELETE", nro_sol, convenio, soe_nro_sec, "", "", 0, "", 0, 0, "", ""
End Sub

Public Sub soa_cab_get_max_nro_sol()
Dim fecha As Date
 Datos.dbo_so_soes_cab "GET_NRO_SOL", 0, cb_codigo_convenio.Text, fecha, "", "", "", 0, 0, 0, "", "", "", fecha, "", "", ""
 With Datos.rsdbo_so_soes_cab
   gl_nro_sol = Datos.rsdbo_so_soes_cab!max_nro_sol
   .Close
 End With
End Sub

Public Sub soa_get_max_nro_sec(nro_sol As Integer, cod_convenio As String)
Dim fecha As Date
 Datos.dbo_so_soes "GET_NRO_SEC", nro_sol, cod_convenio, 0, "", "", 0, "", 0, 0, "", ""
 With Datos.rsdbo_so_soes
   gl_nro_sol = Datos.rsdbo_so_soes!max_nro_sec
   .Close
 End With
End Sub

Private Sub ResetSoes_cab(nro_sol As Integer)
  Me.txtsoc_nro_sol = nro_sol
  'cb_codigo_convenio.ListIndex = 0
  dtp_Fecha.Value = CStr(Date)
  Me.txtsoc_nro_ref = ""
  Me.txtsoc_cod_ent_eje = "MECYD"
  Me.txtsoc_sol_reembolso = "R"
  Me.txtsoc_sol_pago = ""
  Me.txtsoc_sol_desembolso = ""
  Me.txtsoc_sol_reposicion = ""
  Me.txtsoc_pre_justificacion = ""
  Me.txtsoc_tipo_mone_sol.ListIndex = 0
  Me.txtsoc_mon_mone_sol = "0"
  Me.txtsoc_monto_us = "0"
  Me.dcmBancos1.BoundText = ""
  Me.dcmBancos2.BoundText = ""
'  Me.txtsoc_cta_banco = ""
'  Me.txtsoc_fecha_reg = CStr(Date)
  Me.txtsoc_usuario = glusuario
'  Me.dgSoes.ClearFields
End Sub

Private Sub insert_soes_cab(tipo As String)
Dim fecha As Date, tipo_sol, tipo_just, vigente, usuario As String

If tipo = "INSERT" Then
  vigente = "T"
  usuario = ""   'si esta nulo lo borra
  Me.txtsoc_nro_sol.Text = CStr(gl_nro_sol)
ElseIf tipo = "UPDATE" Then    'Si es Update
  vigente = "T"
  usuario = glusuario  'en un UPDATE
End If

If opt_sol_reembolso.Value = True Then
  tipo_sol = "R"
ElseIf opt_sol_pago.Value = True Then
  tipo_sol = "P"
ElseIf opt_sol_desembolso.Value = True Then
  tipo_sol = "D"
ElseIf opt_sol_reposicion.Value = True Then
  tipo_sol = "E"
ElseIf opt_pre_justificacion.Value = True Then
  tipo_sol = "J"
End If
 
'MsgBox tipo_sol & " nro_sol: " & txtsoc_nro_sol.Text & " convenio: " & cb_codigo_convenio.Text
 
 Datos.dbo_so_soes_cab tipo, _
  Val(Me.txtsoc_nro_sol.Text), _
  cb_codigo_convenio.Text, _
  dtp_Fecha.Value, _
  Me.txtsoc_nro_ref.Text, _
  Me.txtsoc_cod_ent_eje.Text, _
  tipo_sol, _
  IIf(Me.txtsoc_tipo_mone_sol.Text = "USD", "Us", "Bs"), _
  Val(Me.txtsoc_mon_mone_sol.Text), _
  Val(Me.txtsoc_monto_us.Text), _
  Me.dcmBancos1.BoundText, _
  Me.dcmBancos2.BoundText, _
  Me.dcmCtas.BoundText, _
  Date, _
  usuario, _
  vigente, _
  Me.txtsoc_nro_sol_def.Text

End Sub

Private Sub llenaSoes_cab()
  
  opt_sol_reembolso.Value = False
  opt_sol_pago.Value = False
  opt_sol_desembolso.Value = False
  opt_sol_reposicion.Value = False
  opt_pre_justificacion.Value = False

  Me.txtsoc_nro_sol = Me.adoTodo.Recordset!soc_nro_sol
'  If Me.adoTodo.Recordset!soc_codigo_convenio = "931/SF-BO" Then
'    cb_codigo_convenio.ListIndex = 0
'  End If
  cb_codigo_convenio.BoundText = Me.adoTodo.Recordset!soc_codigo_convenio
  
  dtp_Fecha.Value = Me.adoTodo.Recordset!soc_fec_elab
  Me.txtsoc_nro_ref = IIf(IsNull(Me.adoTodo.Recordset!soc_nro_ref), "", Me.adoTodo.Recordset!soc_nro_ref)
  Me.txtsoc_cod_ent_eje = Me.adoTodo.Recordset!soc_cod_ent_eje
  If Me.adoTodo.Recordset!soc_sol_reembolso = "R" Then
    opt_sol_reembolso.Value = True
  ElseIf Me.adoTodo.Recordset!soc_sol_reembolso = "P" Then
    opt_sol_pago.Value = True
  ElseIf Me.adoTodo.Recordset!soc_sol_reembolso = "D" Then
    opt_sol_desembolso.Value = True
  ElseIf Me.adoTodo.Recordset!soc_sol_reembolso = "E" Then
    opt_sol_reposicion.Value = True
  ElseIf Me.adoTodo.Recordset!soc_sol_reembolso = "J" Then
    opt_pre_justificacion.Value = True
  End If
  Me.txtsoc_tipo_mone_sol.ListIndex = IIf(Me.adoTodo.Recordset!soc_tipo_mone_sol = "Us", 0, 1)
  Me.txtsoc_mon_mone_sol = Me.adoTodo.Recordset!soc_mon_mone_sol
  Me.txtsoc_monto_us = Me.adoTodo.Recordset!soc_monto_us
  Me.dcmBancos1.BoundText = Me.adoTodo.Recordset!soc_bco_intermedio
  Me.dcmBancos2.BoundText = Me.adoTodo.Recordset!soc_bco_deposito
  Me.dcmCtas.BoundText = Me.adoTodo.Recordset!soc_cta_banco
  Me.txtsoc_usuario = Me.adoTodo.Recordset!soc_usuario
  If Me.adoTodo.Recordset!soc_vigente = "T" Then
    Es_vigente.Caption = "No confirmado"
  Else
    Es_vigente.Caption = "Confirmado o Rechazado"
  End If
  Me.txtsoc_nro_sol_def = IIf(IsNull(Me.adoTodo.Recordset!soc_nro_sol_def), "", Me.adoTodo.Recordset!soc_nro_sol_def)
  adoTodo.Caption = CStr(adoTodo.Recordset.Bookmark) & " de " & CStr(adoTodo.Recordset.RecordCount)
End Sub

Public Sub GetTc2(org_codigo As String, codigo_pago As Integer, cta_codigo As String, tc As Double, cod_beneficiario, par_codigo, pro_proyecto As String)
Dim ok As Boolean, consulta, cta, cta_bcb As String
  cta = GetValor("fc_convenios", "cta_codigo", "codigo_convenio", cb_codigo_convenio.Text)
  cta_bcb = GetValor("fc_convenios", "cta_codigo_bcb", "codigo_convenio", cb_codigo_convenio.Text)
  tc = 0
  consulta = "SELECT top 1 codigo_beneficiario, tipo_cambio, par_codigo, pro_proyecto " _
    & " FROM pago_detalle " _
    & " WHERE ges_gestion = '2002' " _
    & " AND org_codigo = '" & org_codigo & "' " _
    & " AND codigo_pago = " & codigo_pago _
    & " AND ( cta_codigo = '" & cta & "' or cta_codigo = '" & cta_bcb & "' )"
'  MsgBox consulta
  cod_beneficiario = ""
  Datos.dbo_apGeneralSearching consulta
  With Datos.rsdbo_apGeneralSearching
    If Not Datos.rsdbo_apGeneralSearching.EOF Then
      tc = Datos.rsdbo_apGeneralSearching!tipo_cambio
      cod_beneficiario = Datos.rsdbo_apGeneralSearching!codigo_beneficiario
      par_codigo = Datos.rsdbo_apGeneralSearching!par_codigo
      pro_proyecto = Datos.rsdbo_apGeneralSearching!pro_proyecto
    End If
   .Close
  End With

End Sub

Private Sub Form_Load()
Dim fecha As Date
  Datos.dbo_so_soes_cab "DEL_CANCELADOS", 0, "", fecha, "", "", "", 0, 0, 0, "", "", "", fecha, "", "", ""
	Call SeguridadSet(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Cancelar Then
    If MsgBox("¿Desea Salir de esta ventana y cancelar los cambios ingresados?", vbQuestion + vbYesNo, "Diálogo Cerrar") = vbNo Then
      Cancel = -1
    Else
      If Caption = "Solicitud de Reemboloso. Confirme o Cancele" Then
        delete_soes_cab Me.adoTodo.Recordset!soc_nro_sol, Me.adoTodo.Recordset!soc_codigo_convenio
        'soa_cab_refresca "SELECT_TODO"
      End If
    End If
  End If
End Sub

Private Sub opt_confirmados_Click()
'  titulo1.Caption = "Solicitudes Confirmadas"
  TipoListaComp = "SELECT_CONF"
  Me.soa_cab_refresca TipoListaComp
  cmdAdicionar.Enabled = False
End Sub

Private Sub opt_no_confirmados_Click()
  'titulo1.Caption = "Solicitudes No Confirmadas"
  TipoListaComp = "SELECT_NO_CONF"
  Me.soa_cab_refresca TipoListaComp
'  cmdAdicionar.Enabled = True
End Sub

Function valida_reg_soes_cab(definitivo As Boolean) As Boolean
Dim ok As Boolean
  ok = True
  If ok And dcmBancos1.Text = "" Then
    MsgBox "Ingrese banco Intermediario"
    ok = False
  End If
  If ok And dcmBancos2.Text = "" Then
    MsgBox "Ingrese banco Depositario"
    ok = False
  End If
  If ok And (dcmBancos1.Text = dcmBancos2.Text) Then
    MsgBox "Los Bancos Intermediario y Depositario no deben ser iguales"
    ok = False
  End If
  If ok And (dcmBancos1.Text = dcmBancos2.Text) Then
    MsgBox "Los Bancos Intermediario y Depositario no deben ser iguales"
    ok = False
  End If
  If frmSoesMain.Caption = "Nueva Solicitud de Reemboloso" Then  'solo valida cuando es nuevo registro
    If ok And Not (dtp_Fecha.Value >= Date And dtp_Fecha.Value < Date + 15) Then
      MsgBox "Ingrese fecha de elaboración entre hoy y quince dias despues"
      ok = False
    End If
  End If
  If ok And dcmCtas.Text = "" Then
    MsgBox "Ingrese el Nro. de Cuenta Bancaria del Beneficiario"
    ok = False
  End If
  If definitivo Then
    If txtsoc_mon_mone_sol.Text = "0" Then
      MsgBox "Montos de Solicitud estan en cero"
      ok = False
    End If
  End If
  valida_reg_soes_cab = ok
End Function

Public Sub habilitaCasillas(Habilitado As Boolean)
  txtsoc_nro_ref.Enabled = Habilitado
  fr_Presentacion.Enabled = Habilitado
  dcmBancos1.Enabled = Habilitado
  dcmBancos2.Enabled = Habilitado
  cmd_Insert_Cab.Enabled = Habilitado
  cmd_del_cab.Enabled = Habilitado
End Sub

Private Sub set_habilita_boton_todo(cant_regs As Integer, confirmado, autorizado As Boolean, estado_registro As String)
Dim ok, abmSoes As Boolean
  abmSoes = IIf(accion = "ABM_SOES", True, False)
  If estado_registro = "Normal" Then
    dgTodo.Enabled = True
    cb_codigo_convenio.Enabled = False
    cmd_del_cab.Enabled = False
    cmd_update_cab.Caption = "Ver Detalle"
    cmd_update_cab.Enabled = True
    cmd_banco.Enabled = False
    cmd_banco2.Enabled = False
    dtp_Fecha.Enabled = False
    txtsoc_tipo_mone_sol.Enabled = False
    dcmCtas.Enabled = False
    fra_ver_lista.Enabled = True
    txtsoc_nro_sol_def.Enabled = False
  ElseIf estado_registro = "Modificado" Then
    dgTodo.Enabled = False
    cb_codigo_convenio.Enabled = False
    cmd_del_cab.Enabled = True
    cmd_update_cab.Caption = "Modificar"
    cmd_update_cab.Enabled = True  'para modificar
    cmd_banco.Enabled = True
    cmd_banco2.Enabled = True
    dtp_Fecha.Enabled = False
    txtsoc_tipo_mone_sol.Enabled = False
    dcmCtas.Enabled = False
    fra_ver_lista.Enabled = False
    txtsoc_nro_sol_def.Enabled = True
  ElseIf estado_registro = "Insertado" Then
    dgTodo.Enabled = False
    cb_codigo_convenio.Enabled = True
    cmd_del_cab.Enabled = True
    cmd_update_cab.Caption = "Modificar"
    cmd_update_cab.Enabled = True
    cmd_banco.Enabled = True
    cmd_banco2.Enabled = True
    dtp_Fecha.Enabled = True
    txtsoc_tipo_mone_sol.Enabled = True
    dcmCtas.Enabled = True
    fra_ver_lista.Enabled = False
    txtsoc_nro_sol_def.Enabled = True
  End If
  If (Not confirmado And autorizado And (estado_registro = "Insertado" Or estado_registro = "Modificado")) Then
    ok = True
  Else
    ok = False
  End If
  CmdGrabar.Enabled = ok
  CmdCancelar.Enabled = ok
  dgTodo.Enabled = Not ok
  If cant_regs > 0 And Not confirmado And autorizado And estado_registro = "Normal" Then
    ok = True
  Else
    ok = False
  End If
  cmdModificar.Enabled = ok
  If abmSoes And (Not confirmado And autorizado And estado_registro = "Normal") Then
    ok = True
  Else
    ok = False
  End If
  cmdAdicionar.Enabled = ok
  If abmSoes And cant_regs > 0 And autorizado And estado_registro = "Normal" Then
    ok = True
  Else
    ok = False
  End If
  cmdImprimir.Enabled = ok
  If abmSoes And cant_regs > 0 And autorizado And estado_registro = "Normal" And Not confirmado Then
    ok = True
  Else
    ok = False
  End If
  cdmAnular.Enabled = ok
  If abmSoes And cant_regs > 0 And Not confirmado And autorizado And estado_registro = "Normal" Then
    ok = True
  Else
    ok = False
  End If
  cmdConfirmar.Enabled = ok
End Sub

Private Sub llenaCtas(cod_convenio As String)
  Set tCtas = New ADODB.Recordset
  If tCtas.State = 1 Then tCtas.Close
    tCtas.Open "select cta_codigo_bcb as cta From fc_convenios where codigo_convenio = '" & cod_convenio & "' Union " & _
      " select cta_codigo as cta From fc_convenios where codigo_convenio = '" & cod_convenio & "' ", db, adOpenDynamic, adLockReadOnly
  Set ado_Ctas.Recordset = tCtas
  dcmCtas.BoundText = ado_Ctas.Recordset!cta
End Sub

'Private Sub txtsoc_cta_banco_Validate(Cancel As Boolean)
'  If txtsoc_cta_banco.Text = "" Then
'    Cancel = False
'  Else
'    If Not (GetValor("fc_convenios", "cta_codigo_bcb", "codigo_convenio", cb_codigo_convenio.Text) = txtsoc_cta_banco.Text _
'    Or GetValor("fc_convenios", "cta_codigo", "codigo_convenio", cb_codigo_convenio.Text) = txtsoc_cta_banco.Text) Then
'      Cancel = True
'      MsgBox "Nro. de Cuenta invalido"
'    End If
'  End If
'End Sub

'  consulta = "SELECT top 1 codigo_beneficiario, tipo_cambio, par_codigo, pro_proyecto " _
'    & " FROM pago_detalle " _
'    & " WHERE ges_gestion = '2000' " _
'    & " AND org_codigo = '" & org_codigo & "' " _
'    & " AND codigo_pago = " & codigo_pago _
'    & " AND cta_codigo in " _
'    & " ( ( SELECT cta_codigo as cta " _
'    & " FROM fc_convenios " _
'    & " WHERE ges_gestion = '2000' " _
'    & " AND org_codigo = '" & org_codigo & "' ) " _
'    & " UNION " _
'    & " ( SELECT cta_codigo_bcb as cta " _
'    & " WHERE ges_gestion = '2000' " _
'    & " FROM fc_convenios " _
'    & " AND org_codigo = '" & org_codigo & "' " _
'    & " ) ) "


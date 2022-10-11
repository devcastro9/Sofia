VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmDetalleSoes 
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "frmDetalleSoes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   165
      TabIndex        =   28
      Top             =   1860
      Width           =   7965
      Begin MSAdodcLib.Adodc ado_categoria 
         Height          =   330
         Left            =   6165
         Top             =   495
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
      Begin MSDataListLib.DataCombo dcmCategoria 
         Bindings        =   "frmDetalleSoes.frx":324A
         Height          =   315
         Left            =   2280
         TabIndex        =   29
         Top             =   495
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cod_desc_categoria"
         BoundColumn     =   "codigo_categoria"
         Text            =   "Todos"
      End
      Begin MSAdodcLib.Adodc ado_paises 
         Height          =   330
         Left            =   6165
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
      Begin MSDataListLib.DataCombo dcmPaises 
         Bindings        =   "frmDetalleSoes.frx":3266
         Height          =   315
         Left            =   2280
         TabIndex        =   30
         Top             =   180
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "denominacion_pais"
         BoundColumn     =   "codigo_pais"
         Text            =   "Todos"
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pais Origen:"
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
         Index           =   7
         Left            =   1230
         TabIndex        =   32
         Top             =   225
         Width           =   1050
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Categoria de la Inversion:"
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
         Index           =   14
         Left            =   90
         TabIndex        =   31
         Top             =   525
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   165
      TabIndex        =   10
      Top             =   4620
      Width           =   7995
      Begin VB.TextBox txtsoe_cta_bancaria 
         DataField       =   "soe_cta_bancaria"
         Height          =   285
         Left            =   2505
         TabIndex        =   20
         Top             =   945
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtsoe_cant_comp 
         DataField       =   "soe_cant_comp"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4860
         TabIndex        =   19
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox txtsoe_es_contable 
         DataField       =   "soe_es_contable"
         Height          =   285
         Left            =   3075
         TabIndex        =   18
         Top             =   1185
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   5850
         MousePointer    =   4  'Icon
         Picture         =   "frmDetalleSoes.frx":327F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   930
         Width           =   1005
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   6840
         MousePointer    =   4  'Icon
         Picture         =   "frmDetalleSoes.frx":3589
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   930
         Width           =   1005
      End
      Begin VB.CommandButton cdm_borrar_item 
         Caption         =   "Borrar Item"
         Height          =   390
         Left            =   1305
         TabIndex        =   15
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Modificar Item"
         Height          =   390
         Left            =   360
         TabIndex        =   14
         Top             =   1215
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmd_InsertItem 
         Caption         =   "Insertar Item"
         Height          =   390
         Left            =   105
         TabIndex        =   13
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4845
         Picture         =   "frmDetalleSoes.frx":3893
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime el comprobante de Ingreso"
         Top             =   930
         Width           =   1005
      End
      Begin VB.CommandButton cmd_cap_rechazado 
         Caption         =   "Registrar como Rechazado"
         Height          =   525
         Left            =   2520
         TabIndex        =   11
         Top             =   195
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc adoDetalleSoes 
         Height          =   330
         Left            =   120
         Top             =   645
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
      Begin MSMask.MaskEdBox txtsoe_monto_sol_us 
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
         Left            =   6525
         TabIndex        =   21
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtsoe_monto_sol_bs 
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
         Left            =   6510
         TabIndex        =   22
         Top             =   525
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "soe_cta_bancaria:"
         Height          =   255
         Index           =   6
         Left            =   660
         TabIndex        =   27
         Top             =   990
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Comp."
         Height          =   195
         Index           =   10
         Left            =   3930
         TabIndex        =   26
         Top             =   300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "soe_es_contable:"
         Height          =   255
         Index           =   11
         Left            =   1230
         TabIndex        =   25
         Top             =   1230
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto (Us.)"
         Height          =   195
         Index           =   12
         Left            =   5640
         TabIndex        =   24
         Top             =   285
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Monto (Bs.)"
         Height          =   195
         Index           =   13
         Left            =   5640
         TabIndex        =   23
         Top             =   540
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   165
      TabIndex        =   3
      Top             =   705
      Width           =   7965
      Begin VB.TextBox txtsoc_nro_sol 
         BackColor       =   &H80000018&
         DataField       =   "soc_nro_sol"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   165
         Width           =   660
      End
      Begin VB.TextBox txtsoe_cod_convenio 
         BackColor       =   &H80000018&
         DataField       =   "soe_cod_convenio"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   450
         Width           =   3300
      End
      Begin VB.TextBox txtsoe_nro_sec 
         BackColor       =   &H80000018&
         DataField       =   "soe_nro_sec"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   780
         Width           =   660
      End
      Begin Crystal.CrystalReport CryReporte 
         Left            =   6495
         Top             =   630
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   210
         Width           =   1440
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Convenio:"
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
         Left            =   990
         TabIndex        =   8
         Top             =   510
         Width           =   870
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Secuencia:"
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
         Left            =   885
         TabIndex        =   7
         Top             =   810
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid dgDetalleSoes 
      Bindings        =   "frmDetalleSoes.frx":3F7D
      Height          =   1770
      Left            =   165
      TabIndex        =   2
      Top             =   2835
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   3122
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
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "dso_rechazado"
         Caption         =   "Rech."
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
         DataField       =   "denominacion_beneficiario"
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
      BeginProperty Column02 
         DataField       =   "dso_nro_prism"
         Caption         =   "PRISM"
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
         DataField       =   "codigo_pago"
         Caption         =   "Nro. Comp."
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
         DataField       =   "dso_fecha_pago"
         Caption         =   "Fecha P."
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
         DataField       =   "dso_monto_pago"
         Caption         =   "Monto P."
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
      BeginProperty Column06 
         DataField       =   "dso_monto_pago_tot"
         Caption         =   "Monto P. total"
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
      BeginProperty Column07 
         DataField       =   "dso_tc_pago"
         Caption         =   "T.C."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "dso_monto_equi"
         Caption         =   "Monto Equi."
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
      BeginProperty Column09 
         DataField       =   "dso_fin_bid"
         Caption         =   "BID"
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
      BeginProperty Column10 
         DataField       =   "dso_por_fin"
         Caption         =   "%FIN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """%"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "dso_otras_fuentes"
         Caption         =   "Otr. Fuen."
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
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1035.213
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PRESTAMOS DE INVERSION/COOPERACION TECNICA/PEQUEÑOS PROYECTOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   945
      TabIndex        =   1
      Top             =   255
      Width           =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ESTADO DE GASTOS O PAGOS "
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
      Left            =   2190
      TabIndex        =   0
      Top             =   15
      Width           =   3735
   End
End
Attribute VB_Name = "frmDetalleSoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fila_actual As Integer
Dim accion As String, Cancelar As Boolean
Public frmDetalleSoes_ret As String

Public Sub frmDetalleSoes_procesar(proceso, convenio As String, nro_sol As Integer)
  Cancelar = True
  accion = proceso
  bco_codigo_ret = ""
  If proceso = "INSERT" Then
    Caption = "Nuevo Estado de Gastos o Pagos"
    Detalle_soa_refresca True, convenio, nro_sol
    dcmPaises.Enabled = True
    dcmCategoria.Enabled = True
  ElseIf proceso = "SELECT" Then
    Caption = "Estado de Gastos o Pagos"
    Detalle_soa_refresca False, convenio, nro_sol
    dcmPaises.Enabled = False
    dcmCategoria.Enabled = False
  ElseIf proceso = "UPDATE" Then   'este es select...
    Caption = "Estado de Gastos o Pagos"
    Detalle_soa_refresca False, convenio, nro_sol
    dcmPaises.Enabled = False
    dcmCategoria.Enabled = False
  End If
  cmd_InsertItem.Enabled = Not frmSoesMain.cmdModificar.Enabled
  cdm_borrar_item.Enabled = Not frmSoesMain.cmdModificar.Enabled
  Show vbModal
End Sub

Public Sub Detalle_soa_refresca(nuevo As Boolean, convenio As String, nro_sol As Integer)
Dim fecha As Date
  Llena_lista_categoria (convenio)
  If nuevo Then
    'nro_sol = -1
    ResetSoes nro_sol, convenio
  Else
    llenaSoes
  End If
  Datos.dbo_so_detalle_soes "SELECT_HIJOS", nro_sol, convenio, Val(frmDetalleSoes.txtsoe_nro_sec.Text), "", "", 0, 0, "", "", fecha, 0, 0, 0, 0, 0, 0, 0, ""
  With Datos.rsdbo_so_detalle_soes
    Set frmDetalleSoes.adoDetalleSoes.Recordset = .Clone
    frmDetalleSoes.dgDetalleSoes.Refresh
    .Close
  End With
End Sub

Private Sub adoDetalleSoes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If Not (adoDetalleSoes.Recordset.EOF Or adoDetalleSoes.Recordset.BOF) Then
    adoDetalleSoes.Caption = CStr(adoDetalleSoes.Recordset.Bookmark) & " de " & CStr(adoDetalleSoes.Recordset.RecordCount)
  Else
    adoDetalleSoes.Caption = " 0 de 0"
  End If
End Sub

Private Sub cdm_borrar_item_Click()
  If Not (adoDetalleSoes.Recordset.EOF Or adoDetalleSoes.Recordset.BOF) Then
    delete_detalle_soes Me.adoDetalleSoes.Recordset!soc_nro_sol, Me.adoDetalleSoes.Recordset!soe_cod_convenio, Me.adoDetalleSoes.Recordset!soe_nro_sec, Me.adoDetalleSoes.Recordset!ges_gestion, Me.adoDetalleSoes.Recordset!org_codigo, Me.adoDetalleSoes.Recordset!codigo_pago
    Detalle_soa_refresca False, Me.txtsoe_cod_convenio, Val(Me.txtsoc_nro_sol)
    LlenaTotalMonto
  End If
End Sub

Private Sub cmd_cap_rechazado_Click()
  If frmDetalleSoes.adoDetalleSoes.Recordset!dso_rechazado = "Si" Then
    MsgBox "Este comprobante fue rechazado"
  Else
    frmCompRechazo.frmCompRechazo_procesar ""
  End If
End Sub

Private Sub cmd_InsertItem_Click()
  If validaRegSoes Then
    If Caption = "Nuevo Estado de Gastos o Pagos" Then
      Caption = "Estado de Gastos o Pagos. Confirme o Cancele"
      frmSoesMain.soa_get_max_nro_sec Val(Me.txtsoc_nro_sol.Text), Me.txtsoe_cod_convenio.Text
      dcmCategoria.Enabled = False
      dcmPaises.Enabled = False
      insert_soes
      frmSoesMain.soa_refresca False
      frmSoesMain.adoSoes.Recordset.MoveLast
'    frmSoesMain.adoSoes.Recordset.Find "soc_nro_sol= '" & txtsoc_nro_sol.Text & "'", , adSearchForward
'    frmSoesMain.adoSoes.Recordset.Find "soe_cod_convenio = '" & txtsoe_cod_convenio.Text & "'", , adSearchForward
'    frmSoesMain.adoSoes.Recordset.Find "soe_nro_sec = '" & txtsoe_nro_sec.Text & "'", , adSearchForward
    End If
    comprobantes_refresca 1, 3600, Me.txtsoe_cod_convenio.Text, dcmCategoria.BoundText
    frmComprobante.Show vbModal
  End If
End Sub

Private Sub cmd_print_Click()
  If GetValorGeneral("select org_codigo as retorno from fc_convenios where codigo_convenio = '" & frmSoesMain.adoTodo.Recordset!soc_codigo_convenio & "'") = "411" Then
    Rep001 "REP_SOES_ESTADO", "\rep_bid_estado_pago_v1.rpt", "ESTADO DE GASTOS O PAGOS", Val(txtsoc_nro_sol.Text), txtsoe_cod_convenio.Text, Val(txtsoe_nro_sec.Text), ""
  Else
    Rep001 "REP_SOES_ESTADO", "\rep_bm_estado_pago_v1.rpt", "ESTADO DE GASTOS O PAGOS", Val(txtsoc_nro_sol.Text), txtsoe_cod_convenio.Text, Val(txtsoe_nro_sec.Text), ""
  End If
End Sub

Private Sub CmdCancelar_Click()
  Cancelar = True
  Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim monto_sol_us, monto_sol_bs As Double
  If MsgBox("Esta seguro de Grabar?", vbYesNo) = vbYes Then
    If validaListaComp Then
    
      Datos.dbo_so_soes "GET_SOES_US_BOB", Val(txtsoc_nro_sol.Text), txtsoe_cod_convenio.Text, 0, "", "", 0, "", 0, 0, "", ""
      With Datos.rsdbo_so_soes
        monto_sol_us = Datos.rsdbo_so_soes!monto_sol_us
        monto_sol_bs = Datos.rsdbo_so_soes!monto_sol_bs
        .Close
      End With
      If accion = "INSERT" Then 'si es nuevo se adiciona
        frmSoesMain.txtsoc_monto_us.Text = Str(Val(frmSoesMain.txtsoc_monto_us.Text) + Val(frmDetalleSoes.txtsoe_monto_sol_us.Text))
        frmSoesMain.txtsoc_mon_mone_sol.Text = Str(Val(frmSoesMain.txtsoc_mon_mone_sol.Text) + Val(frmDetalleSoes.txtsoe_monto_sol_bs.Text))
      ElseIf accion = "UPDATE" Then
        frmSoesMain.txtsoc_monto_us.Text = Str(Val(frmSoesMain.txtsoc_monto_us.Text) - frmSoesMain.adoSoes.Recordset!soe_monto_sol_us + Val(frmDetalleSoes.txtsoe_monto_sol_us.Text))
        frmSoesMain.txtsoc_mon_mone_sol.Text = Str(Val(frmSoesMain.txtsoc_mon_mone_sol.Text) - frmSoesMain.adoSoes.Recordset!soe_monto_sol_bs + Val(frmDetalleSoes.txtsoe_monto_sol_bs.Text))
      End If
      
      If frmDetalleSoes.Caption = "Estado de Gastos o Pagos" Then
        update_soes False
      Else
        update_soes True
      End If
      frmDetalleSoes.Caption = "Estado de Gastos o Pagos"
      frmSoesMain.soa_refresca False
      Cancelar = False
      Unload Me
    End If
  End If
End Sub

Public Sub comprobantes_refresca(del As Integer, al As Integer, cod_convenio, codigo_categoria As String)
Dim cod_cta, cod_cta_bcb, convenio As String
  convenio = cod_convenio
  cod_cta = mod_librerias.GetValor("fc_convenios", "cta_codigo", "codigo_convenio", convenio)
  cod_cta_bcb = mod_librerias.GetValor("fc_convenios", "cta_codigo_bcb", "codigo_convenio", convenio)
  'MsgBox "del:" + Str(del) + " al:" + Str(al) + " convenio:" + cod_convenio
  Datos.dbo_so_comprobantes "SELECT", del, al, cod_convenio, codigo_categoria, cod_cta, cod_cta_bcb
  With Datos.rsdbo_so_comprobantes
    Set frmComprobante.adoComprobantes.Recordset = .Clone
    frmComprobante.dgComprobantes.Refresh
    .Close
  End With
End Sub

Private Sub delete_detalle_soes(nro_sol As Integer, cod_convenio As String, nro_sec As Integer, ges_gestion As String, org_codigo As String, codigo_pago As Integer)
Dim fecha As Date
 Datos.dbo_so_detalle_soes "DELETE", nro_sol, cod_convenio, nro_sec, ges_gestion, org_codigo, codigo_pago, 0, "", "", fecha, 0, 0, 0, 0, 0, 0, 0, ""
End Sub

Private Sub update_soes(nuevo As Boolean)
  If nuevo Then
    frmDetalleSoes.txtsoe_nro_sec.Text = CStr(gl_nro_sol)
  End If
  Datos.dbo_so_soes "UPDATE", Val(frmDetalleSoes.txtsoc_nro_sol.Text), frmDetalleSoes.txtsoe_cod_convenio.Text, Val(frmDetalleSoes.txtsoe_nro_sec.Text), frmDetalleSoes.txtsoe_cta_bancaria.Text, frmDetalleSoes.dcmPaises.BoundText, Val(txtsoe_cant_comp.Text), frmDetalleSoes.txtsoe_es_contable.Text, Val(frmDetalleSoes.txtsoe_monto_sol_us.Text), Val(frmDetalleSoes.txtsoe_monto_sol_bs.Text), frmDetalleSoes.dcmCategoria.BoundText, "T"
End Sub

Private Sub delete_soes()
  frmDetalleSoes.txtsoe_nro_sec.Text = CStr(gl_nro_sol)
  Datos.dbo_so_soes "DELETE", Val(frmDetalleSoes.txtsoc_nro_sol.Text), frmDetalleSoes.txtsoe_cod_convenio.Text, Val(frmDetalleSoes.txtsoe_nro_sec.Text), frmDetalleSoes.txtsoe_cta_bancaria.Text, frmDetalleSoes.dcmPaises.BoundText, 0, frmDetalleSoes.txtsoe_es_contable.Text, Val(frmDetalleSoes.txtsoe_monto_sol_us.Text), Val(frmDetalleSoes.txtsoe_monto_sol_bs.Text), frmDetalleSoes.dcmCategoria.Text, ""
End Sub

Public Sub Rep001(tipoRep As String, ArchRep As String, titulo1 As String, nro_sol As Integer, cod_convenio As String, nro_sec As Integer, literal As String)
Dim fecha As Date, tot_acu_otr As Double, secuencia As Integer
  secuencia = 2
  CryReporte.ReportFileName = App.Path & ArchRep
  CryReporte.StoredProcParam(0) = CDate("01/01/1900")
  CryReporte.StoredProcParam(1) = CDate("01/01/1900")
  CryReporte.StoredProcParam(2) = tipoRep
  CryReporte.StoredProcParam(3) = nro_sol
  CryReporte.StoredProcParam(4) = cod_convenio
  CryReporte.StoredProcParam(5) = nro_sec
  CryReporte.Formulas(0) = "fFecha1 ='" & CDate("01/01/1900") & "'"
  CryReporte.Formulas(1) = "fFecha2 ='" & CDate("01/01/1900") & "'"
  
  If titulo1 <> "" Then
    CryReporte.Formulas(secuencia) = "Titulo1 = '" & titulo1 & "'"
    secuencia = secuencia + 1
  End If
  
  If ArchRep = "\rep_bid_estado_pago_v1.rpt" Or _
     ArchRep = "\rep_bm_estado_pago_v1.rpt" Then 'nro real de solicitud, le pasamos porque no se puede en el reporte
    CryReporte.Formulas(secuencia) = "NRO_SOL_DEF = '" & frmSoesMain.txtsoc_nro_sol_def.Text & "'"
    secuencia = secuencia + 1
  End If
  
  If ArchRep = "\rep_bid_form1_soes_v1.rpt" Or _
     ArchRep = "\rep_bm_form1_soes_v1.rpt" Then
    CryReporte.Formulas(secuencia) = "literal ='" & literal & "'"
    secuencia = secuencia + 1
  End If
  
  If ArchRep = "\rep_bid_form2_soes_v1.rpt" Or _
     ArchRep = "\rep_bm_form2_soes_v1.rpt" Then
    Datos.dbo_so_soes "GET_ACU_US_OTR", nro_sol, cod_convenio, 0, "", "", 0, "", 0, 0, "", ""
    With Datos.rsdbo_so_soes
      tot_acu_otr = Datos.rsdbo_so_soes!acu_us_otr
      .Close
    End With
    CryReporte.Formulas(secuencia) = "tot_acu_us_otr = " & tot_acu_otr
    secuencia = secuencia + 1
  End If
  
  iResult = CryReporte.PrintReport
  If iResult <> 0 Then
    MsgBox CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Public Sub Rep002(tipoRep, ArchRep, titulo1 As String)
  frmSoesMain.CryReporte.ReportFileName = App.Path & ArchRep
  frmSoesMain.CryReporte.StoredProcParam(0) = CDate("01/01/2000")
  frmSoesMain.CryReporte.StoredProcParam(1) = CDate("30/12/2000")
  frmSoesMain.CryReporte.StoredProcParam(2) = tipoRep
  frmSoesMain.CryReporte.StoredProcParam(3) = "%"
  frmSoesMain.CryReporte.StoredProcParam(4) = "%"
  frmSoesMain.CryReporte.StoredProcParam(5) = "%"
  frmSoesMain.CryReporte.StoredProcParam(6) = "%"
  frmSoesMain.CryReporte.StoredProcParam(7) = "%"
  frmSoesMain.CryReporte.StoredProcParam(8) = "%"
  frmSoesMain.CryReporte.StoredProcParam(9) = "%"
  frmSoesMain.CryReporte.StoredProcParam(10) = "%"
  'CryReporte.Formulas(0) = "Titulo1 ='" & titulo1 & "'"
  iResult = frmSoesMain.CryReporte.PrintReport
  If iResult <> 0 Then
    MsgBox frmSoesMain.CryReporte.LastErrorNumber & " : " & CryReporte.LastErrorString, vbCritical + vbOKOnly, "Error..."
  End If
End Sub

Private Sub insert_soes()
 frmDetalleSoes.txtsoe_nro_sec.Text = CStr(gl_nro_sol)
 Datos.dbo_so_soes "INSERT", Val(frmDetalleSoes.txtsoc_nro_sol.Text), frmDetalleSoes.txtsoe_cod_convenio.Text, Val(frmDetalleSoes.txtsoe_nro_sec.Text), frmDetalleSoes.txtsoe_cta_bancaria.Text, frmDetalleSoes.dcmPaises.BoundText, 0, frmDetalleSoes.txtsoe_es_contable.Text, Val(frmDetalleSoes.txtsoe_monto_sol_us.Text), Val(frmDetalleSoes.txtsoe_monto_sol_bs.Text), frmDetalleSoes.dcmCategoria.BoundText, "N"
End Sub

Private Sub dgDetalleSoes_AfterColUpdate(ByVal ColIndex As Integer)
Dim Monto_100_bs, Monto_en_mon_ope, monto_bid_bs As Double
Dim monto_otr_ftes_bs, por_fin As Double
Dim ok As Boolean
  ok = True
  If adoDetalleSoes.Recordset!dso_por_fin <= 0 Or _
     adoDetalleSoes.Recordset!dso_por_fin > 100 Then
     MsgBox "Ingrese un valor mayor a cero y menor o igual a 100"
     adoDetalleSoes.Recordset!dso_por_fin = 0
     ok = False
  End If
  
  If ok Then
    por_fin = adoDetalleSoes.Recordset!dso_por_fin
    monto_otr_ftes_bs = (adoDetalleSoes.Recordset!dso_monto_pago / por_fin) * (100 - por_fin)
    Monto_100_bs = adoDetalleSoes.Recordset!dso_monto_pago + monto_otr_ftes_bs
    monto_bid_bs = Monto_100_bs - monto_otr_ftes_bs
    adoDetalleSoes.Recordset!dso_monto_pago_tot = Monto_100_bs
    If frmSoesMain.txtsoc_tipo_mone_sol.Text = "USD" Then
      monto_bid_bs = monto_bid_bs / adoDetalleSoes.Recordset!dso_tc_pago
      monto_otr_ftes_bs = monto_otr_ftes_bs / adoDetalleSoes.Recordset!dso_tc_pago
      Monto_100_bs = Monto_100_bs / adoDetalleSoes.Recordset!dso_tc_pago
    End If
    adoDetalleSoes.Recordset!dso_otras_fuentes = monto_otr_ftes_bs
    adoDetalleSoes.Recordset!dso_fin_bid = monto_bid_bs
    adoDetalleSoes.Recordset!dso_monto_equi = Monto_100_bs
    fila_actual = adoDetalleSoes.Recordset.Bookmark
    LlenaTotalMonto
  End If
End Sub

Function validaListaComp() As Boolean
Dim ok, unNulo As Boolean
  ok = False
  unNulo = False
  If adoDetalleSoes.Recordset.RecordCount > 0 Then
    adoDetalleSoes.Recordset.MoveFirst
    While Not adoDetalleSoes.Recordset.EOF
      If adoDetalleSoes.Recordset!dso_por_fin = 0 Or _
         IsNull(adoDetalleSoes.Recordset!dso_por_fin) Then
         unNulo = True
      End If
      adoDetalleSoes.Recordset.MoveNext
    Wend
    If unNulo Then
      MsgBox "Todos los valores de %FIN deben ser distintos de cero"
    Else
      ok = True
    End If
  Else
    MsgBox "Se debe tener por lo menos un comprobante"
  End If
  validaListaComp = ok
End Function

Function seguir() As Boolean
Dim ret As Boolean
  ret = True
  If MsgBox("Esta seguro de Cancelar?", vbYesNo) = vbYes Then
    If frmDetalleSoes.Caption = "Estado de Gastos o Pagos. Confirme o Cancele" Then
      delete_soes
      frmSoesMain.soa_refresca False
'      MsgBox "se borro soes"
    End If
    frmDetalleSoes.Caption = "Estado de Gastos o Pagos"
    ret = False
    'Unload Me
  End If
  Cancelar = ret
End Function

Private Sub Form_Unload(Cancel As Integer)
  If Cancelar Then
    If MsgBox("¿Desea Salir de esta ventana y cancelar los cambios ingresados?", vbQuestion + vbYesNo, "Diálogo Cerrar") = vbNo Then
      Cancel = -1
    Else
      If Caption = "Estado de Gastos o Pagos. Confirme o Cancele" Then
        frmSoesMain.delete_soes Val(Me.txtsoc_nro_sol), Me.txtsoe_cod_convenio, Val(Me.txtsoe_nro_sec)
        frmSoesMain.soa_refresca False
      End If
    End If
  End If
End Sub

Function validaRegSoes() As Boolean
Dim ok As Boolean
ok = True
  If ok And Me.dcmCategoria.Text = "" Then
    ok = False
    MsgBox "Ingrese Categoria"
  End If
  If ok And Me.dcmPaises.Text = "" Then
    ok = False
    MsgBox "Ingrese Pais"
  End If
  validaRegSoes = ok
End Function

Public Sub LlenaTotalMonto()
Dim total_us, total_bs, cant, monto_pago, tc, aux As Double
  total_us = 0
  total_bs = 0
  cant = 0
  If adoDetalleSoes.Recordset.RecordCount > 0 Then
    adoDetalleSoes.Recordset.MoveFirst
    While Not adoDetalleSoes.Recordset.EOF
      tc = IIf(IsNull(adoDetalleSoes.Recordset!dso_tc_pago), 0, adoDetalleSoes.Recordset!dso_tc_pago)
      monto_pago = IIf(IsNull(adoDetalleSoes.Recordset!dso_fin_bid), 0, adoDetalleSoes.Recordset!dso_fin_bid)
      If frmSoesMain.txtsoc_tipo_mone_sol.Text = "USD" Then
        total_us = total_us + monto_pago
        total_bs = total_bs + monto_pago * tc
      Else
        total_bs = total_bs + monto_pago
        total_us = total_us + monto_pago / tc
      End If
      cant = cant + 1
      adoDetalleSoes.Recordset.MoveNext
    Wend
    Me.txtsoe_cant_comp.Text = Str(cant)
    Me.txtsoe_monto_sol_us.Text = Str(total_us)
    Me.txtsoe_monto_sol_bs.Text = Str(total_bs)
  End If
End Sub

Public Sub Llena_lista_categoria(cod_convenio As String)
'MsgBox cod_convenio
  Set tFc_categoria_financiador = New ADODB.Recordset
  If tFc_categoria_financiador.State = 1 Then tFc_categoria_financiador.Close
    tFc_categoria_financiador.Open "SELECT codigo_categoria, codigo_categoria + ' - ' + denominacion_categoria as cod_desc_categoria FROM fc_categoria_financiador WHERE codigo_convenio = '" & cod_convenio & "' ", db, adOpenDynamic, adLockReadOnly
  Set frmDetalleSoes.ado_categoria.Recordset = tFc_categoria_financiador
  
  Set tpaises = New ADODB.Recordset
  If tpaises.State = 1 Then tpaises.Close
    tpaises.Open "SELECT codigo_pais, denominacion_pais FROM paises ", db, adOpenDynamic, adLockReadOnly
  Set frmDetalleSoes.ado_paises.Recordset = tpaises
  
End Sub

Private Sub ResetSoes(nro_sol As Integer, convenio As String)
  txtsoc_nro_sol = Str(nro_sol)
  txtsoe_cod_convenio = convenio
  txtsoe_nro_sec = ""
  txtsoe_cta_bancaria = ""
  dcmPaises.BoundText = ""
  txtsoe_cant_comp = "0"
  txtsoe_es_contable = ""
  txtsoe_monto_sol_bs = "0"
  txtsoe_monto_sol_us = "0"
  dcmCategoria.BoundText = ""
End Sub

Private Sub llenaSoes()
'  Llena_lista_categoria (cb_codigo_convenio.Text)
  txtsoc_nro_sol = frmSoesMain.adoSoes.Recordset!soc_nro_sol
  txtsoe_cod_convenio = frmSoesMain.adoSoes.Recordset!soe_cod_convenio
  txtsoe_nro_sec = frmSoesMain.adoSoes.Recordset!soe_nro_sec
  txtsoe_monto_sol_us = frmSoesMain.adoSoes.Recordset!soe_monto_sol_us
  txtsoe_monto_sol_bs = frmSoesMain.adoSoes.Recordset!soe_monto_sol_bs
  dcmPaises.BoundText = frmSoesMain.adoSoes.Recordset!soe_pais_origen
  dcmCategoria.BoundText = frmSoesMain.adoSoes.Recordset!soe_codigo_categoria
End Sub


VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmIngresosabm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     Registro de Ingresos..."
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   12015
   Icon            =   "ingresos.frx":0000
   Moveable        =   0   'False
   Picture         =   "ingresos.frx":038A
   ScaleHeight     =   8730
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frmmensaje 
      Height          =   2475
      Left            =   4320
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   5115
      Begin VB.Label LblMensaje 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   660
         TabIndex        =   64
         Top             =   780
         Width           =   3255
      End
      Begin VB.Label LblTitMensaje 
         BackColor       =   &H8000000D&
         Caption         =   "  Espere un momento por favor ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   60
         TabIndex        =   63
         Top             =   120
         Width           =   5015
      End
   End
   Begin VB.Frame FraIngresosDat 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7896
      Left            =   4000
      TabIndex        =   33
      Top             =   840
      Width           =   8010
      Begin VB.TextBox TxtMonto_bolivianos 
         DataField       =   "monto_bolivianos"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1635
         TabIndex        =   30
         ToolTipText     =   "Formato con Punto Decimal"
         Top             =   7290
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo DtCrbr_codigo 
         Bindings        =   "ingresos.frx":35D4
         DataField       =   "rbr_codigo"
         DataSource      =   "AdoFc_Rubro_ingresos"
         Height          =   315
         Left            =   40
         TabIndex        =   24
         ToolTipText     =   "Elije el Código del Rubro"
         Top             =   5475
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rbr_codigo"
         BoundColumn     =   "rbr_descripcion"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCrbr_descripcion 
         Bindings        =   "ingresos.frx":35F7
         DataField       =   "rbr_descripcion"
         DataSource      =   "AdoFc_Rubro_ingresos"
         Height          =   315
         Left            =   1460
         TabIndex        =   25
         ToolTipText     =   "Elije la Descripción del Rubro"
         Top             =   5475
         Width           =   6400
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rbr_descripcion"
         BoundColumn     =   "rbr_codigo"
         Text            =   ""
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1460
         TabIndex        =   58
         Text            =   "VICEMINISTERIO DE EDUCACION INICAL, PRIMARIA Y SECUNDARIA"
         Top             =   2640
         Width           =   6400
      End
      Begin VB.TextBox TxtUNI_CODIGO 
         Enabled         =   0   'False
         Height          =   285
         Left            =   40
         TabIndex        =   57
         Text            =   "VEIPS"
         Top             =   2640
         Width           =   1410
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_documento 
         Bindings        =   "ingresos.frx":361A
         DataField       =   "Denominacion_documento"
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         ToolTipText     =   "Elije la descripción del Documento de Respaldo"
         Top             =   1920
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_documento"
         BoundColumn     =   "Codigo_documento"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCodigo_documento 
         Bindings        =   "ingresos.frx":3641
         DataField       =   "Codigo_documento"
         Height          =   312
         Left            =   36
         TabIndex        =   15
         ToolTipText     =   "Elije el Código del Documento de Respaldo"
         Top             =   1920
         Width           =   1416
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Codigo_documento"
         BoundColumn     =   "Denominacion_documento"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_moneda 
         Bindings        =   "ingresos.frx":3668
         DataField       =   "Denominacion_moneda"
         Height          =   315
         Left            =   1620
         TabIndex        =   27
         Top             =   6720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_moneda"
         BoundColumn     =   "Tipo_moneda"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_tipo_solicitud 
         Bindings        =   "ingresos.frx":3685
         DataField       =   "Denominacion_tipo_solicitud"
         Height          =   315
         Left            =   1460
         TabIndex        =   12
         ToolTipText     =   "Elije el Tipo de Solicitud"
         Top             =   795
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_tipo_solicitud"
         BoundColumn     =   "Codigo_tipo_solicitud"
         Text            =   ""
      End
      Begin VB.TextBox Txtmonto_dolares 
         DataField       =   "monto_dolares"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6180
         TabIndex        =   32
         ToolTipText     =   "Formato con Punto Decimal"
         Top             =   7290
         Width           =   1215
      End
      Begin VB.TextBox TxtNumero_documento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "AdoIngresos"
         Height          =   285
         Left            =   6320
         MaxLength       =   20
         TabIndex        =   17
         ToolTipText     =   "Número de Documento de Respaldo (hasta 20 caracteres)"
         Top             =   1920
         Width           =   1510
      End
      Begin MSDataListLib.DataCombo DtCFte_codigo 
         Bindings        =   "ingresos.frx":36A5
         DataField       =   "fte_codigo"
         Height          =   315
         Left            =   45
         TabIndex        =   18
         ToolTipText     =   "Elije el Código de Fuente de Financiamiento"
         Top             =   3360
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "fte_codigo"
         BoundColumn     =   "Fte_descripcion_larga"
         Text            =   ""
      End
      Begin VB.TextBox TxtConcepto 
         DataField       =   "concepto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   525
         Left            =   40
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         ToolTipText     =   "Acepta hasta 100 caracteres"
         Top             =   6120
         Width           =   7770
      End
      Begin VB.TextBox TxtTipo_cambio 
         DataField       =   "tipo_cambio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16394
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6180
         TabIndex        =   29
         Top             =   6720
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker DTPFecha_Ingreso 
         DataField       =   "fecha_ingreso"
         Height          =   285
         Left            =   6300
         TabIndex        =   11
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   25034753
         CurrentDate     =   36541
      End
      Begin VB.TextBox TxtCodigo_solicitud 
         DataField       =   "codigo_solicitud"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   13
         ToolTipText     =   "Acepta hasta 6 caracteres"
         Top             =   820
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo DtCCta_descripcion_larga 
         Bindings        =   "ingresos.frx":36C3
         Height          =   315
         Left            =   1460
         TabIndex        =   23
         ToolTipText     =   "Elije el Nombre de la Cuenta Bancaria"
         Top             =   4755
         Visible         =   0   'False
         Width           =   6400
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Cta_descripcion_larga"
         BoundColumn     =   "Cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCta_codigo 
         Bindings        =   "ingresos.frx":36E7
         Height          =   315
         Left            =   40
         TabIndex        =   22
         ToolTipText     =   "Elije el Código de la Cuenta Bancaria"
         Top             =   4755
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Cta_codigo"
         BoundColumn     =   "Cta_descripcion_larga"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCOrg_descripcion 
         Bindings        =   "ingresos.frx":370B
         DataField       =   "Org_descripcion"
         Height          =   315
         Left            =   1460
         TabIndex        =   21
         ToolTipText     =   "Elije el Nombre del Organismo de Financiamiento"
         Top             =   4065
         Width           =   6400
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "Org_descripcion"
         BoundColumn     =   "Org_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCOrg_codigo 
         Bindings        =   "ingresos.frx":372D
         DataField       =   "Org_codigo"
         Height          =   315
         Left            =   40
         TabIndex        =   20
         ToolTipText     =   "Elije el Código de Organismo de Financiamiento"
         Top             =   4065
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Org_codigo"
         BoundColumn     =   "Org_descripcion"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc AdoFc_cuenta_bancaria 
         Height          =   330
         Left            =   1860
         Top             =   4890
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
         Caption         =   "Adodc2"
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
      Begin MSDataListLib.DataCombo DtCFte_descripcion_larga 
         Bindings        =   "ingresos.frx":375E
         DataField       =   "Fte_descripcion_larga"
         Height          =   315
         Left            =   1460
         TabIndex        =   19
         ToolTipText     =   "Elije el Nombre de la Fuente de Financiamiento"
         Top             =   3345
         Width           =   6400
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "Fte_descripcion_larga"
         BoundColumn     =   "fte_codigo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc AdoFte_financia 
         Height          =   450
         Left            =   2400
         Top             =   3255
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   794
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
      Begin MSAdodcLib.Adodc AdoOrganismo_finan 
         Height          =   330
         Left            =   4320
         Top             =   3960
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
      Begin MSDataListLib.DataCombo DtCDenominacion_tipo 
         Bindings        =   "ingresos.frx":377D
         DataField       =   "Denominacion_tipo"
         Height          =   315
         Left            =   1455
         TabIndex        =   14
         ToolTipText     =   "Elije el Tipo de Comprobante"
         Top             =   1260
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_tipo"
         BoundColumn     =   "codigo_tipo"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adoac_documento_respaldo 
         Height          =   330
         Left            =   1560
         Top             =   2040
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSAdodcLib.Adodc AdoTipo_comprobante 
         Height          =   330
         Left            =   2220
         Top             =   1380
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSAdodcLib.Adodc AdoTipo_solicitud 
         Height          =   330
         Left            =   2640
         Top             =   840
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
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
      Begin MSAdodcLib.Adodc AdoTipo_moneda 
         Height          =   330
         Left            =   2040
         Top             =   6840
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSAdodcLib.Adodc AdoFc_Rubro_ingresos 
         Height          =   330
         Left            =   1860
         Top             =   5640
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
         Caption         =   "rubro"
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
      Begin MSDataListLib.DataCombo DtCdenominacion_beneficiario 
         Bindings        =   "ingresos.frx":379F
         Height          =   312
         Left            =   1460
         TabIndex        =   66
         ToolTipText     =   "Elije el Nombre del Beneficiario"
         Top             =   4848
         Visible         =   0   'False
         Width           =   6396
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "denominacion_beneficiario"
         BoundColumn     =   "codigo_beneficiario"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCcodigo_beneficiario 
         Bindings        =   "ingresos.frx":37C0
         Height          =   312
         Left            =   36
         TabIndex        =   67
         ToolTipText     =   "Elije el Código del Beneficiario"
         Top             =   4860
         Visible         =   0   'False
         Width           =   1416
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_beneficiario"
         BoundColumn     =   "denominacion_beneficiario"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc AdoFc_beneficiario 
         Height          =   336
         Left            =   3720
         Top             =   4980
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
         Caption         =   "Adodc2"
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
      Begin VB.Label LblmontoRecaudado 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto Recaudado: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         TabIndex        =   65
         Top             =   1260
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Rubro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   60
         Top             =   5250
         Width           =   585
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   7880
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   10
         X2              =   7880
         Y1              =   590
         Y2              =   590
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Registro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   45
         TabIndex        =   59
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Técnica :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   56
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Documento Respaldo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   55
         Top             =   1695
         Width           =   1845
      End
      Begin VB.Label LblGes_Gestion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ges_gestion"
         Height          =   255
         Left            =   3645
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblCorrelativo_ingreso 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "correlativo_ingreso"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Bolivianos :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   50
         Top             =   7320
         Width           =   1500
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Dólares :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4695
         TabIndex        =   49
         Top             =   7320
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6600
         TabIndex        =   48
         Top             =   1695
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Fuente Financiamiento :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   44
         Top             =   3120
         Width           =   1950
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Concepto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   43
         Top             =   5880
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   42
         Top             =   6795
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Moneda :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   41
         Top             =   6780
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4755
         TabIndex        =   40
         Top             =   285
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Solicitud :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   45
         TabIndex        =   39
         Top             =   885
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro Solicitud Desembolso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4500
         TabIndex        =   38
         Top             =   885
         Width           =   2175
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gestión :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2880
         TabIndex        =   37
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Lblcuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   36
         Top             =   4530
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label LblCod_Poa 
         AutoSize        =   -1  'True
         Caption         =   "Organismo Financiamiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   40
         TabIndex        =   35
         Top             =   3840
         Width           =   2250
      End
      Begin VB.Label LblCod_Sol 
         AutoSize        =   -1  'True
         Caption         =   "Nro Comprobante :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   34
         Top             =   285
         Width           =   1560
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         Caption         =   "Beneficiario :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   60
         TabIndex        =   68
         Top             =   4560
         Visible         =   0   'False
         Width           =   1116
      End
   End
   Begin VB.Frame FraIngresosNav 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7896
      Left            =   900
      TabIndex        =   31
      Top             =   840
      Width           =   3240
      Begin VB.OptionButton OptFilGral2 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   120
         Width           =   795
      End
      Begin VB.OptionButton OptFilGral1 
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
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc AdoIngresos 
         Height          =   336
         Left            =   60
         Top             =   6420
         Width           =   3048
         _ExtentX        =   5371
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
      Begin MSDataGridLib.DataGrid DtGIngresos 
         Bindings        =   "ingresos.frx":37E1
         Height          =   5955
         Left            =   30
         TabIndex        =   72
         Top             =   420
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   10504
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
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
         ColumnCount     =   34
         BeginProperty Column00 
            DataField       =   "correlativo_ingreso"
            Caption         =   "Correl"
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
            Caption         =   "Org"
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
            DataField       =   "codigo_tipo"
            Caption         =   "Tipo"
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
            DataField       =   "correlativo_anterior"
            Caption         =   "Anter"
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
         BeginProperty Column05 
            DataField       =   "estado_recaudado"
            Caption         =   "R"
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
            DataField       =   "estado_desafectado"
            Caption         =   "F"
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
            DataField       =   "estado_anulado"
            Caption         =   "L"
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
            DataField       =   "ges_gestion"
            Caption         =   "ges_gestion"
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
            DataField       =   "codigo_solicitud"
            Caption         =   "codigo_solicitud"
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
            DataField       =   "codigo_documento"
            Caption         =   "codigo_documento"
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
            DataField       =   "codigo_tipo_solicitud"
            Caption         =   "codigo_tipo_solicitud"
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
            DataField       =   "codigo_beneficiario"
            Caption         =   "codigo_beneficiario"
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
            DataField       =   "fecha_ingreso"
            Caption         =   "fecha_ingreso"
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
            DataField       =   "tipo_cambio"
            Caption         =   "tipo_cambio"
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
         BeginProperty Column16 
            DataField       =   "concepto"
            Caption         =   "concepto"
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
            DataField       =   "tipo_comp"
            Caption         =   "tipo_comp"
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
         BeginProperty Column19 
            DataField       =   "fte_codigo"
            Caption         =   "fte_codigo"
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
            DataField       =   "rbr_codigo"
            Caption         =   "rbr_codigo"
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
            DataField       =   "cheque_o_trf"
            Caption         =   "cheque_o_trf"
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
            DataField       =   "Bco_codigo"
            Caption         =   "Bco_codigo"
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
            DataField       =   "cta_codigo"
            Caption         =   "cta_codigo"
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
            DataField       =   "numero_documento"
            Caption         =   "numero_documento"
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
         BeginProperty Column26 
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
         BeginProperty Column27 
            DataField       =   "monto_recaudado_dolares"
            Caption         =   "monto_recaudado_dolares"
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
            DataField       =   "estado_Aprobacion"
            Caption         =   "estado_Aprobacion"
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
            DataField       =   "ultimo"
            Caption         =   "ultimo"
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
         BeginProperty Column31 
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
         BeginProperty Column32 
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
         BeginProperty Column33 
            DataField       =   "estado_conciliacion"
            Caption         =   "estado_conciliacion"
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
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   374.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   195.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   180.283
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   180.283
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   195.024
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column26 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column27 
               ColumnWidth     =   1964.976
            EndProperty
            BeginProperty Column28 
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column29 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column30 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column31 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column32 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column33 
               ColumnWidth     =   1484.787
            EndProperty
         EndProperty
      End
      Begin VB.Label Label16 
         Caption         =   "Donde:            D = Devengado    R = Recaudado     L = Anulado      F = Desafectado"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1260
         TabIndex        =   61
         Top             =   6780
         Width           =   1815
      End
   End
   Begin VB.Frame Fra 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   0.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2
      TabIndex        =   0
      Top             =   0
      Width           =   12020
      Begin VB.Label LblAccion 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   240
         TabIndex        =   69
         Top             =   540
         Width           =   45
      End
      Begin VB.Label Lblusuario 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "USUARIO: "
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
         Left            =   9000
         TabIndex        =   51
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label LblCF301 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE INGRESOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3555
         TabIndex        =   28
         Top             =   210
         Width           =   3765
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE INGRESOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3590
         TabIndex        =   71
         Top             =   225
         Width           =   3765
      End
      Begin VB.Image Image3 
         Height          =   750
         Left            =   45
         Picture         =   "ingresos.frx":37FB
         Stretch         =   -1  'True
         Top             =   45
         Width           =   11925
      End
   End
   Begin VB.Frame FraOpciones 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7870
      Left            =   1
      TabIndex        =   54
      Top             =   840
      Width           =   900
      Begin VB.CommandButton CmdActualTeso 
         Caption         =   "Actualiza Tesorerí"
         Height          =   720
         Left            =   75
         Picture         =   "ingresos.frx":5E05
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Sale del Formulario de Ingresos"
         Top             =   6210
         Width           =   770
      End
      Begin VB.CommandButton CmdCopiar 
         Caption         =   "Copiar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   70
         Picture         =   "ingresos.frx":600F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copia el comprobante de Ingreso a uno nuevo"
         Top             =   1900
         Width           =   770
      End
      Begin VB.CommandButton CmdAprueba 
         Caption         =   "Aprobar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   60
         Picture         =   "ingresos.frx":6219
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Aprueba el comprobante de Ingreso"
         Top             =   5400
         Width           =   770
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   720
         Left            =   70
         MaskColor       =   &H8000000F&
         Picture         =   "ingresos.frx":6423
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Busca un Comprobante de Ingreso"
         Top             =   3630
         Width           =   770
      End
      Begin VB.CommandButton CmdModificar 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   60
         Picture         =   "ingresos.frx":662D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modifica el comprobante de Ingreso"
         Top             =   1050
         Width           =   770
      End
      Begin VB.CommandButton CmdAñadir 
         Caption         =   "Adicionar"
         Height          =   720
         Left            =   60
         MousePointer    =   4  'Icon
         Picture         =   "ingresos.frx":6837
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Adiciona un comprobante de Ingreso"
         Top             =   180
         Width           =   770
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "Errado"
         Enabled         =   0   'False
         Height          =   720
         Left            =   60
         Picture         =   "ingresos.frx":6B41
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Anula el comprobante de Ingreso"
         Top             =   2760
         Width           =   770
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   720
         Left            =   70
         Picture         =   "ingresos.frx":722B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Sale del Formulario de Ingresos"
         Top             =   7080
         Width           =   770
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   60
         Picture         =   "ingresos.frx":7435
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprime el comprobante de Ingreso"
         Top             =   4500
         Width           =   770
      End
      Begin Crystal.CrystalReport Cry 
         Left            =   360
         Top             =   4860
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Image Image4 
         Height          =   7710
         Left            =   30
         Picture         =   "ingresos.frx":7B1F
         Stretch         =   -1  'True
         Top             =   75
         Width           =   825
      End
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
      Height          =   7896
      Left            =   1
      TabIndex        =   45
      Top             =   840
      Visible         =   0   'False
      Width           =   900
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   720
         Left            =   70
         MousePointer    =   4  'Icon
         Picture         =   "ingresos.frx":A138
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1050
         Width           =   770
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Grabar"
         Height          =   720
         Left            =   70
         MousePointer    =   4  'Icon
         Picture         =   "ingresos.frx":A342
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   180
         Width           =   770
      End
      Begin VB.Image Image1 
         Height          =   7785
         Left            =   45
         Picture         =   "ingresos.frx":A54C
         Stretch         =   -1  'True
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "mnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAccion 
         Caption         =   "Recaudado"
         Index           =   0
      End
      Begin VB.Menu mnuAccion 
         Caption         =   "Desafectado"
         Index           =   1
      End
      Begin VB.Menu mnuAccion 
         Caption         =   "Anular Recaudado"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FrmIngresosabm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  SAF-2001 / FE
' Módulo:                   Ejecución Presupuestaria de Ingresos
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmIngresosabm.frm
' Descipción :              Registro de Ingresos Presupuestarios
' Formularios relacionados: MainMenu.frm (Padre)
'                           ComprobIngreso.rpt (Crystal Reports ver. 7.0)
' Autor:                    Greco Viscarra Iturri
' Versión:                  2.0
' cd now 19930209
'========================================================================================

' ULTIMAS MOD GVI 28/06/001

Option Explicit

Dim sino As String
Dim v_añadir As Integer

Dim v_añadirstat As Integer
Dim v_cod_solicitud As Integer
Dim rstIngresos As New ADODB.Recordset
Dim rstOrganismo_finan As New ADODB.Recordset
Dim rstFte_financia As New ADODB.Recordset
Dim rstFc_bancos As New ADODB.Recordset
Dim rstFc_cuenta_bancaria As New ADODB.Recordset
Dim rstac_documento_respaldo As New ADODB.Recordset
Dim swgraba As Integer
Dim marca1 As BookmarkEnum
Dim correlativo1 As Integer
Dim correlativo_ingreso1 As String
Dim Org_Codigo1 As String
Dim ges_gestion1 As String
Dim rstTipo_comprobante As New ADODB.Recordset
Dim rstTipo_solicitud As New ADODB.Recordset
Dim rstTipo_moneda As New ADODB.Recordset
Dim rstFc_Rubro_ingresos As New ADODB.Recordset
Dim rstCodComp As New ADODB.Recordset
Dim rstfc_beneficiario As New ADODB.Recordset
Dim Cont_Comp As Integer
Dim rstdestino As New ADODB.Recordset
Dim buscasi As Integer
Dim operadorbus As String
Dim campobus As String
Dim V_accion As String
Dim fte_codigo1 As String
Dim swcopiar As Integer
Dim swmodificar As Integer

Dim ClBuscaGrid As CompBusquedas.ClBuscaEnGridExterno
Dim EntrarAdo As Boolean 'Para que al aprobar no muestre uno por uno
Dim queryinicial As String
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String

Private Sub CmdCF306_Click()
'===== Salida del Módulo
  sino = MsgBox("¿Está seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
  If sino = vbYes Then
    Unload Me
  End If
End Sub

Private Sub AdoIngresos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'===== Actualización de Despliegue de Datos
   If (Not AdoIngresos.Recordset.EOF) And (Not AdoIngresos.Recordset.BOF) Then
        If Not IsNull(AdoIngresos.Recordset("Correlativo_ingreso")) Then
            LblCorrelativo_ingreso = IIf(IsNull(AdoIngresos.Recordset("Correlativo_ingreso")) = True, " ", AdoIngresos.Recordset("Correlativo_ingreso"))
            lblges_gestion = IIf(IsNull(AdoIngresos.Recordset("Ges_Gestion")) = True, " ", AdoIngresos.Recordset("Ges_Gestion"))
            txtCodigo_solicitud = IIf(IsNull(AdoIngresos.Recordset("Codigo_solicitud")) = True, " ", AdoIngresos.Recordset("Codigo_solicitud"))
            DTPFecha_Ingreso = IIf(IsNull(AdoIngresos.Recordset("Fecha_Ingreso")) = True, " ", AdoIngresos.Recordset("Fecha_Ingreso"))
            txtTipo_Cambio = IIf(IsNull(AdoIngresos.Recordset("Tipo_cambio")) = True, 0, AdoIngresos.Recordset("Tipo_cambio"))
            txtConcepto = IIf(IsNull(AdoIngresos.Recordset("Concepto")) = True, " ", AdoIngresos.Recordset("Concepto"))
            Txtmonto_dolares = IIf(IsNull(AdoIngresos.Recordset("monto_dolares")) = True, 0, AdoIngresos.Recordset("monto_dolares"))
            txtMonto_bolivianos = IIf(IsNull(AdoIngresos.Recordset("Monto_bolivianos")) = True, 0, AdoIngresos.Recordset("Monto_bolivianos"))
            txtNumero_documento.Text = IIf(IsNull(AdoIngresos.Recordset("numero_documento")) = True, 0, AdoIngresos.Recordset("numero_documento"))
            
            DtCrbr_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("rbr_codigo")) = True, " ", AdoIngresos.Recordset("rbr_codigo"))
            DtCrbr_descripcion.Text = DtCrbr_codigo.BoundText
            
            DtCdenominacion_moneda.BoundText = IIf(IsNull(AdoIngresos.Recordset("tipo_moneda")) = True, "", AdoIngresos.Recordset("tipo_moneda"))
            
            Select Case AdoIngresos.Recordset("Codigo_tipo")
              Case "DYR"
                Set rstTipo_comprobante = New ADODB.Recordset
                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
                AdoTipo_comprobante.Refresh
                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
                DtCDenominacion_tipo.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo")) = True, " ", AdoIngresos.Recordset("Codigo_tipo"))
                LblmontoRecaudado.Visible = False
                Call activar_Obj
              Case "REC"
                Set rstTipo_comprobante = New ADODB.Recordset
                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'REC' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
                AdoTipo_comprobante.Refresh
                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
                DtCDenominacion_tipo.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo")) = True, " ", AdoIngresos.Recordset("Codigo_tipo"))
                LblmontoRecaudado.Visible = False
                Call desactivar_Obj
'                DtCDenominacion_tipo.Enabled = False
'                CmdCopiar.Enabled = False
              Case "DEV"
                Set rstTipo_comprobante = New ADODB.Recordset
                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
                AdoTipo_comprobante.Refresh
                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
                DtCDenominacion_tipo.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo")) = True, " ", AdoIngresos.Recordset("Codigo_tipo"))
                LblmontoRecaudado.Caption = " Monto Recaudado: " & CStr(AdoIngresos.Recordset("monto_recaudado_dolares"))
                LblmontoRecaudado.Visible = True
                Call activar_Obj
'                DtCDenominacion_tipo.Enabled = True
'                CmdCopiar.Enabled = True
              Case "DES"
                Set rstTipo_comprobante = New ADODB.Recordset
                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'DES' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
                AdoTipo_comprobante.Refresh
                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
                DtCDenominacion_tipo.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo")) = True, " ", AdoIngresos.Recordset("Codigo_tipo"))
                LblmontoRecaudado.Visible = False
                Call desactivar_Obj
                Me.cmdCopiar.Enabled = False
'                DtCDenominacion_tipo.Enabled = False
              Case "ANI"
'verificar con tia que es anu nuevo tipo de compro.
                Set rstTipo_comprobante = New ADODB.Recordset
                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'ANI' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
                AdoTipo_comprobante.Refresh
                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
                DtCDenominacion_tipo.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo")) = True, " ", AdoIngresos.Recordset("Codigo_tipo"))
                LblmontoRecaudado.Visible = False
                Call desactivar_Obj
                Me.cmdCopiar.Enabled = False
'                DtCDenominacion_tipo.Enabled = False
            End Select


            '0000
            DtCDenominacion_tipo_solicitud.BoundText = IIf(IsNull(AdoIngresos.Recordset("Codigo_tipo_solicitud")) = True, " ", AdoIngresos.Recordset("Codigo_tipo_solicitud"))
            
            DtCCodigo_documento.Text = IIf(IsNull(AdoIngresos.Recordset("Codigo_documento")) = True, " ", AdoIngresos.Recordset("Codigo_documento"))
            DtCDenominacion_documento.Text = DtCCodigo_documento.BoundText
            
            DtCFte_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("fte_codigo")) = True, " ", AdoIngresos.Recordset("fte_codigo"))
            DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
            
            DtCOrg_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("org_codigo")) = True, " ", AdoIngresos.Recordset("org_codigo"))
            DtCOrg_descripcion.Text = DtCOrg_codigo.BoundText
            
            If AdoIngresos.Recordset("Codigo_tipo") = "DEV" Then
              lblBeneficiario.Visible = True
              DtCcodigo_beneficiario.Text = IIf(IsNull(AdoIngresos.Recordset("Codigo_beneficiario")) = True, " ", AdoIngresos.Recordset("Codigo_beneficiario"))
              DtCdenominacion_beneficiario.Text = DtCcodigo_beneficiario.BoundText
              DtCcodigo_beneficiario.Visible = True
              DtCdenominacion_beneficiario.Visible = True
              lblcuenta.Visible = False
              DtCCta_codigo.Visible = False
              DtCCta_descripcion_larga.Visible = False
            End If

            If AdoIngresos.Recordset("Codigo_tipo") = "REC" Or AdoIngresos.Recordset("Codigo_tipo") = "DYR" Or AdoIngresos.Recordset("Codigo_tipo") = "ANI" Then
              DtCCta_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("Cta_Codigo")) = True, "", AdoIngresos.Recordset("Cta_Codigo"))
              DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
              lblcuenta.Visible = True
              DtCCta_codigo.Visible = True
              DtCCta_descripcion_larga.Visible = True
              lblBeneficiario.Visible = False
              DtCcodigo_beneficiario.Visible = False
              DtCdenominacion_beneficiario.Visible = False
            End If
            
            CmdModificar.Enabled = True
            CmdBorrar.Enabled = True
            'AQUI VERIFICAR QUIEN TIENE ACCESO A APROBAR
            'If ((AdoIngresos.Recordset("estado_aprobacion") = "N") Or (IsNull(AdoIngresos.Recordset("estado_aprobacion")))) And (UCase(GlUsuario) = "F_FLORES" Or UCase(GlUsuario) = "A_ITURRI" Or UCase(GlUsuario) = "M_YAÑEZ" Or UCase(GlUsuario) = "J_CRUZ" Or UCase(GlUsuario) = "SAF" Or UCase(GlUsuario) = "ISM001" Or UCase(GlUsuario) = "MEC002" Or UCase(GlUsuario) = "MEY001" Or UCase(GlUsuario) = "MYB159" Or UCase(GlUsuario) = "FFL001") Then
            If ((AdoIngresos.Recordset("estado_aprobacion") = "N") Or (IsNull(AdoIngresos.Recordset("estado_aprobacion")))) Then
                CmdAprueba.Enabled = True
            Else
                CmdAprueba.Enabled = False
            End If
            If (AdoIngresos.Recordset!estado_aprobacion = "E") Or (AdoIngresos.Recordset!Codigo_Tipo = "DES" Or AdoIngresos.Recordset!Codigo_Tipo = "ANI") Then
              cmdCopiar.Enabled = False
            Else
              cmdCopiar.Enabled = True
            End If
            If (AdoIngresos.Recordset("estado_aprobacion") = "S") Or (AdoIngresos.Recordset("estado_aprobacion") = "E") Then
              CmdBorrar.Enabled = False
              CmdModificar.Enabled = False
            Else
              CmdBorrar.Enabled = True
              CmdModificar.Enabled = True
            End If
            
          mnuAccion(0).Enabled = False
          mnuAccion(1).Enabled = False
          mnuAccion(2).Enabled = False
          With AdoIngresos
            If (.Recordset!estado_devengado = "S") And (Trim(.Recordset!estado_recaudado) = "" Or IsNull(.Recordset!estado_recaudado)) And (Trim(.Recordset!estado_desafectado) = "" Or IsNull(.Recordset!estado_desafectado)) Then
                'mnuAccion(0).Enabled = True
                If .Recordset!monto_dolares > .Recordset!monto_recaudado_dolares Then
                  mnuAccion(0).Enabled = True
                Else
                  mnuAccion(0).Enabled = False
                End If
                If .Recordset!monto_recaudado_dolares <= 0 Then
                  mnuAccion(1).Enabled = True
                Else
                  mnuAccion(1).Enabled = False
                End If
'                sw = 1
            End If
            If (.Recordset!estado_recaudado = "S") And (Trim(.Recordset!estado_devengado) = "" Or IsNull(.Recordset!estado_devengado)) And (Trim(.Recordset!estado_desafectado) = "" Or IsNull(.Recordset!estado_desafectado)) Then
                mnuAccion(0).Enabled = False
                mnuAccion(1).Enabled = True
                'mnuAccion(2).Enabled = True
'                sw = 1
            End If
            If (.Recordset!estado_devengado = "S") And (.Recordset!estado_recaudado = "S") And (Trim(.Recordset!estado_desafectado) = "" Or IsNull(.Recordset!estado_desafectado)) Then
                mnuAccion(0).Enabled = False
                mnuAccion(2).Enabled = True
' que se hace cuando se anula un recaudado, se anulan toDOS LOS REGISTROS RECUADADOS?
'                sw = 2
            End If
          End With
' FIN AHORA ***************************

        Else
          mnuAccion(0).Enabled = False
          mnuAccion(1).Enabled = False
          mnuAccion(2).Enabled = False
'          mnuAccion(3).Enabled = False

          LblCorrelativo_ingreso = ""
          lblges_gestion = ""
          txtCodigo_solicitud = ""
          DTPFecha_Ingreso = Format(Date, "dd/mm/yyy")
          txtTipo_Cambio = 0
          txtConcepto = ""
          Txtmonto_dolares = 0
          txtMonto_bolivianos = 0
          DtCFte_codigo.Text = ""
          DtCFte_descripcion_larga.Text = ""
          DtCOrg_codigo.Text = ""
          DtCOrg_descripcion.Text = ""
          DtCCta_codigo.Text = ""
          DtCCta_descripcion_larga.Text = ""
      End If
   End If
End Sub

Private Sub CmdActualTeso_Click()
  Dim rstacum As New ADODB.Recordset
  Dim rstdestino As New ADODB.Recordset
  Set rstacum = New ADODB.Recordset
  Set rstdestino = New ADODB.Recordset
  If rstdestino.State = 1 Then rstdestino.Close
  rstdestino.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
  While Not rstdestino.EOF
    If rstacum.State = 1 Then rstacum.Close
    rstacum.Open "select sum (monto_bolivianos) as cta_acum from fo_ingresos where cta_codigo = '" & rstdestino("cta_codigo") & "'and estado_aprobacion = 'S' and estado_recaudado = 'S' and( estado_anulado <> 'S' and estado_desafectado <> 'S')", db, adOpenStatic, adLockReadOnly
    If IsNull(rstacum("cta_acum")) Then
    Else
      Print rstacum("cta_acum")
      rstdestino("cta_ingresos") = rstacum("cta_acum")
      rstdestino.Update
    End If
    If rstacum.State = 1 Then rstacum.Close
    rstdestino.MoveNext
  Wend
End Sub

Private Sub CmdAñadir_Click()
'===== Proceso para Añadir y/o modificar Datos
    v_añadir = 1
    FraIngresosNav.Enabled = False
    FraIngresosDat.Enabled = True
    FraOpciones.Visible = False
    FraOpciones2.Visible = True
    DTPFecha_Ingreso.Value = Date
    lblges_gestion = Year(DTPFecha_Ingreso.Value)
    If swcopiar = 1 Then
      LblAccion = "Copiando registros..."
      DtCOrg_codigo.Enabled = False
      If V_accion = "ANI" Or V_accion = "DES" Then
        'Me.txtTipo_Cambio = GlTipoCambioOficial GVI 28/06/001
        Print rstIngresos!monto_dolares
        Print rstIngresos!tipo_cambio
        Txtmonto_dolares = Round(((txtMonto_bolivianos / GlTipoCambioOficial) * -1), 2)
        txtMonto_bolivianos = Round((txtMonto_bolivianos * -1), 2)
'        DtCDenominacion_tipo.BoundText = "ANI"
      End If
    Else
      LblAccion = "Añadiendo registros..."
    End If
    If v_añadir = 1 Then
        If Not (AdoIngresos.Recordset.BOF) Or Not (AdoIngresos.Recordset.EOF) Then
          If swcopiar = 0 Then 'ultimo
            LblCorrelativo_ingreso = ""
            lblges_gestion = ""
            txtCodigo_solicitud = ""
            DTPFecha_Ingreso = Format(Date, "dd/mm/yyyy")
            'TxtTipo_cambio = 0
            txtConcepto = ""
            Txtmonto_dolares = 0
            txtMonto_bolivianos = 0
            DtCDenominacion_tipo_solicitud = ""
            txtCodigo_solicitud.Text = ""
            DtCDenominacion_tipo.Text = ""
            DtCCodigo_documento = ""
            DtCDenominacion_documento = ""
            txtNumero_documento = ""
            DtCrbr_codigo = ""
            DtCrbr_descripcion = ""
            DtCFte_codigo.Text = ""
            DtCFte_descripcion_larga.Text = ""
            DtCOrg_codigo.Text = ""
            DtCOrg_descripcion.Text = ""
            DtCCta_codigo.Text = ""
            DtCCta_descripcion_larga.Text = ""
            lblges_gestion.Caption = Year(Date) '"2001"
            DTPFecha_Ingreso.Value = Format(Date, "dd/mm/yyyy")
            txtTipo_Cambio = GlTipoCambioOficial
            Set rstTipo_comprobante = New ADODB.Recordset
            If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
            rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
            Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
            AdoTipo_comprobante.Refresh

            'Public GlTipoCambioMercado As Currency

          End If 'ultimo
            Call activar_Obj
            DtCFte_codigo.Enabled = True
'            DtCOrg_codigo.Enabled = True
          Select Case V_accion
            Case "REC"
              Call desactivar_Obj
              DtCCta_codigo.Enabled = True
              DtCCta_descripcion_larga.Enabled = True
              DtCdenominacion_moneda.Enabled = True
              Txtmonto_dolares.Enabled = True
              txtMonto_bolivianos.Enabled = True
            Case "ANI"
              Call desactivar_Obj
'              DtCCta_codigo.Enabled = True
'              DtCCta_descripcion_larga.Enabled = True
'              DtCDenominacion_moneda.Enabled = True
'LAST
'              Txtmonto_dolares.Text = (Txtmonto_dolares.Text * -1)
'              TxtMonto_bolivianos.Enabled = (TxtMonto_bolivianos * -1)
              Txtmonto_dolares.Enabled = False
              txtMonto_bolivianos.Enabled = False
          End Select
         
      ' DtCDenominacion_tipo.BoundText
            
        End If
    End If
End Sub

'Private Sub CmdAñadir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''  CmdAñadir.BackColor = &HC0FFFF
'End Sub

Private Sub CmdAprueba_Click()
'===== Proceso para generar Asientos Contables Automáticos "CAD" y "CAR"

  sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
  If sino = vbYes Then
    If AdoIngresos.Recordset("codigo_tipo") = "REC" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("monto_dolares") < rstdestino("monto_recaudado_dolares") + AdoIngresos.Recordset("monto_dolares") Then
          MsgBox "El monto que está intentando recaudar en dolares es mayor al DEVENGADO, por fsavor corrija Monto: Devengado: " & CStr(rstdestino("monto_dolares")) & " Solo puede recaudar :" & CStr(rstdestino("monto_dolares") - rstdestino("monto_recaudado_dolares")), vbOKOnly + vbCritical, "ERROR en el minto de Recaudo"
          Exit Sub
        End If
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If

'**** aqui consultar tia que hacer ***************
    If AdoIngresos.Recordset("codigo_tipo") = "DES" Then
      
    End If
    
    If AdoIngresos.Recordset("codigo_tipo") = "ANI" Then
      
    End If
'**** aqui consultar tia que hacer ***************
    
    
    Dim cta_deb1 As String
    Dim Subcta_deb11 As String
    Dim Subcta_deb21 As String
    
    Dim cta_credito1 As String
    Dim Subcta_cred11 As String
    Dim Subcta_cred21 As String
    
    Dim cod_ant As Integer
    Dim org_ant As String
    
    If DtCCta_codigo.Text <> "01" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
      If Not rstFc_cuenta_bancaria.EOF Then
        fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
      Else
      End If
    Else
        fte_codigo1 = Me.DtCFte_codigo.Text
    End If
    
    If AdoIngresos.Recordset!Codigo_Tipo = "DEV" Or AdoIngresos.Recordset!Codigo_Tipo = "DES" Then
      fte_codigo1 = Me.DtCFte_codigo.Text
    End If
    
    Dim I As Integer
    Dim j As Integer
    Dim v_Tipo_Comp(1, 2)
    If AdoIngresos.Recordset("codigo_tipo") = "DYR" Then
      j = 2
      v_Tipo_Comp(1, 1) = "CAD"
      v_Tipo_Comp(1, 2) = "CAR"
    Else
      j = 1
      v_Tipo_Comp(1, 1) = IIf(AdoIngresos.Recordset("codigo_tipo") = "DEV", "CAD", IIf(AdoIngresos.Recordset("codigo_tipo") = "REC", "CAR", IIf(AdoIngresos.Recordset("codigo_tipo") = "DES", "DES", IIf(AdoIngresos.Recordset("codigo_tipo") = "ANI", "ANI", ""))))
    End If
    
    
    For I = 1 To j
      If rstdestino.State = 1 Then rstdestino.Close
      If v_Tipo_Comp(1, I) = "CAD" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
      End If
      If v_Tipo_Comp(1, I) = "CAR" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "41", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
      End If
      If v_Tipo_Comp(1, I) = "ANI" Or v_Tipo_Comp(1, I) = "REC" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "41", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
      End If
      
      If v_Tipo_Comp(1, I) = "DEV" Or v_Tipo_Comp(1, I) = "DES" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
      End If

      If v_Tipo_Comp(1, I) = "" Then
        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
        Exit Sub
      End If

      If rstdestino.RecordCount < 1 Then
        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
        Exit Sub
      End If
    Next
    
    If rstdestino.State = 1 Then rstdestino.Close
    db.BeginTrans
    Frmmensaje.Visible = True
    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
    Dim d_cta_nombre_1 As String
    Dim d_aux1_1 As String
    Dim d_aux2_1 As String
    Dim d_aux3_1 As String
    Dim h_cta_nombre_1 As String
    Dim h_aux1_1 As String
    Dim h_aux2_1 As String
    Dim h_aux3_1 As String
    If rstdestino.State = 1 Then rstdestino.Close
    '===== ini registro de co_comprobante_M =====
    
    For I = 1 To j
' nuevo ini
      If v_Tipo_Comp(1, I) = "CAD" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
      End If
      If v_Tipo_Comp(1, I) = "CAR" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "41", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
      End If
      If v_Tipo_Comp(1, I) = "ANI" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "41", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
      End If
      If v_Tipo_Comp(1, I) = "DES" Then
        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
      End If
      If rstdestino.RecordCount > 0 Then
        cta_deb1 = rstdestino("cta_deb")
        Subcta_deb11 = rstdestino("Subcta_deb1")
        Subcta_deb21 = rstdestino("Subcta_deb2")
        cta_credito1 = rstdestino("cta_credito")
        Subcta_cred11 = rstdestino("Subcta_cred1")
        Subcta_cred21 = rstdestino("Subcta_cred2")
      Else
        MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
'        Exit Sub
      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rstdestino.RecordCount > 0 Then
        d_cta_nombre_1 = rstdestino("NombreCta")
        d_aux1_1 = rstdestino("aux1")
        d_aux2_1 = rstdestino("aux2")
        d_aux3_1 = rstdestino("aux3")
      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
      If rstdestino.RecordCount > 0 Then
        h_cta_nombre_1 = rstdestino("NombreCta")
        h_aux1_1 = rstdestino("aux1")
        h_aux2_1 = rstdestino("aux2")
        h_aux3_1 = rstdestino("aux3")
      End If

' nuevo fin
      '===== ini GENERA EL CODIGO DE COMPROBANTE ====
      Set rstCodComp = New ADODB.Recordset
      rstCodComp.CursorLocation = adUseClient
      If rstCodComp.State = 1 Then rstCodComp.Close
      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
      If rstCodComp.RecordCount > 0 Then
        Cont_Comp = Val(rstCodComp!numero_correlativo)
        Cont_Comp = Cont_Comp + 1
        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
        rstCodComp.Update
      End If
      If rstCodComp.State = 1 Then rstCodComp.Close
      '===== fin TERMINA GENERACION DE COMPROBANTE =====

      '==== ini registro co_comprobantre_m
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
      End If
      rstdestino.AddNew
      rstdestino("cod_comp") = Cont_Comp
      rstdestino("cod_trans") = AdoIngresos.Recordset!correlativo_anterior  'AdoIngresos.Recordset("correlativo_ingreso")
      rstdestino("org_codigo") = AdoIngresos.Recordset("org_codigo")
      rstdestino("cod_trans_detalle") = 1
      rstdestino("Num_Respaldo") = AdoIngresos.Recordset("numero_documento")
      rstdestino("Fecha_A") = Date
      rstdestino("codigo_beneficiario") = "-"
      rstdestino("glosa") = AdoIngresos.Recordset("Concepto")
      rstdestino("status") = "S"
      rstdestino("ges_gestion") = AdoIngresos.Recordset("ges_gestion")
      rstdestino("codigo_documento") = AdoIngresos.Recordset("codigo_documento")
      rstdestino("Tipo_Comp") = v_Tipo_Comp(1, I) 'IIf(AdoIngresos.Recordset("codigo_tipo") = "DEV", "CAD", IIf(AdoIngresos.Recordset("codigo_tipo") = "REC", "CAR", v_Tipo_Comp(i)))
      rstdestino("Usr_Usuario") = GlUsuario
      rstdestino("Fecha_registro") = Format(Date, "dd/mm/yyyy")
      rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
      rstdestino.Update
      '==== fin registro co_comprobantre_m
      
      '===== ini registra CO_diaRIO =========
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
      If rstdestino.RecordCount > 0 Then
          
      Else
        rstdestino.AddNew
        rstdestino("Cod_Comp") = Cont_Comp
      End If
      
      rstdestino("Tipo_Comp") = v_Tipo_Comp(1, I)
      rstdestino("Cod_Comp_C") = Cont_Comp
      'If v_Tipo_Comp(1, i) = "DEV" Or v_Tipo_Comp(1, i) = "REC" Then
      If (AdoIngresos.Recordset("codigo_tipo") = "DEV") Or (AdoIngresos.Recordset("codigo_tipo") = "REC") Or (AdoIngresos.Recordset("codigo_tipo") = "DYR") Then
        rstdestino("D_Cuenta") = cta_deb1
        rstdestino("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino("D_Subcta1") = Subcta_deb11
        rstdestino("D_SubCta2") = Subcta_deb21
        rstdestino("D_Aux1") = d_aux1_1
        rstdestino("D_Aux2") = d_aux2_1
        rstdestino("D_Aux3") = d_aux3_1
        If d_aux1_1 = "01" Then
          rstdestino("D_Cta_Larga") = IIf(Len(Trim(AdoIngresos.Recordset("codigo_beneficiario"))) > 0, AdoIngresos.Recordset("codigo_beneficiario"), "-")
        End If
        If d_aux1_1 = "02" Then
          rstdestino("D_Cta_Larga") = AdoIngresos.Recordset("cta_codigo")
        End If
        rstdestino("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino("D_MontoBs") = IIf(AdoIngresos.Recordset("monto_bolivianos") < 0, (AdoIngresos.Recordset("monto_bolivianos") * -1), AdoIngresos.Recordset("monto_bolivianos"))
        rstdestino("D_MontoDl") = IIf(AdoIngresos.Recordset("monto_dolares") < 0, (AdoIngresos.Recordset("monto_dolares") * -1), AdoIngresos.Recordset("monto_dolares"))
        rstdestino("D_Cambio") = AdoIngresos.Recordset("tipo_cambio")
        rstdestino("H_Cuenta") = cta_credito1
        rstdestino("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino("H_SubCta1") = Subcta_cred11
        rstdestino("H_SubCta2") = Subcta_cred21
        rstdestino("H_Aux1") = h_aux1_1
        rstdestino("H_Aux2") = h_aux2_1
        rstdestino("H_Aux3") = h_aux3_1
        'rstdestino("H_Cta_Larga") = "VEIPS"
        If h_aux1_1 = "01" Then
          rstdestino("h_Cta_Larga") = IIf(Len(Trim(AdoIngresos.Recordset("codigo_beneficiario"))) > 0, AdoIngresos.Recordset("codigo_beneficiario"), "-")
          'DtCCta_descripcion_larga
        End If
        If h_aux1_1 = "02" Then
          rstdestino("h_Cta_Larga") = AdoIngresos.Recordset("cta_codigo")
        End If
        rstdestino("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino("H_MontoBs") = IIf(AdoIngresos.Recordset("monto_bolivianos") < 0, (AdoIngresos.Recordset("monto_bolivianos") * -1), AdoIngresos.Recordset("monto_bolivianos"))
        rstdestino("H_MontoDl") = IIf(AdoIngresos.Recordset("monto_dolares") < 0, (AdoIngresos.Recordset("monto_dolares") * -1), AdoIngresos.Recordset("monto_dolares"))
        rstdestino("H_Cambio") = AdoIngresos.Recordset("tipo_cambio")
      End If
      
      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
      If (AdoIngresos.Recordset("codigo_tipo") = "DES") Or (AdoIngresos.Recordset("codigo_tipo") = "ANI") Then
        'desafecta un devengado
        rstdestino("D_Cuenta") = cta_credito1
        rstdestino("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino("D_Subcta1") = Subcta_cred11
        rstdestino("D_SubCta2") = Subcta_cred21
        rstdestino("D_Aux1") = h_aux1_1
        rstdestino("D_Aux2") = h_aux2_1
        rstdestino("D_Aux3") = h_aux3_1
'        rstdestino("D_Cta_Larga") = "VEIPS"
        If h_aux1_1 = "01" Then
          rstdestino("D_Cta_Larga") = IIf(Len(Trim(AdoIngresos.Recordset("codigo_beneficiario"))) > 0, AdoIngresos.Recordset("codigo_beneficiario"), "-")
        End If
        If h_aux1_1 = "02" Then
          rstdestino("D_Cta_Larga") = AdoIngresos.Recordset("cta_codigo")
        End If
        rstdestino("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
        rstdestino("D_MontoBs") = IIf(AdoIngresos.Recordset("monto_bolivianos") < 0, (AdoIngresos.Recordset("monto_bolivianos") * -1), AdoIngresos.Recordset("monto_bolivianos"))
        rstdestino("D_MontoDl") = IIf(AdoIngresos.Recordset("monto_dolares") < 0, (AdoIngresos.Recordset("monto_dolares") * -1), AdoIngresos.Recordset("monto_dolares"))
        rstdestino("D_Cambio") = AdoIngresos.Recordset("tipo_cambio")
        
        rstdestino("H_Cuenta") = cta_deb1
        rstdestino("H_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
        rstdestino("H_SubCta1") = Subcta_deb11
        rstdestino("H_SubCta2") = Subcta_deb21
        rstdestino("H_Aux1") = d_aux1_1
        rstdestino("H_Aux2") = d_aux2_1
        rstdestino("H_Aux3") = d_aux3_1
'        rstdestino("H_Cta_Larga") = "VEIPS"
        If d_aux1_1 = "01" Then
          rstdestino("h_Cta_Larga") = IIf(Len(Trim(AdoIngresos.Recordset("codigo_beneficiario"))) > 0, AdoIngresos.Recordset("codigo_beneficiario"), "-")
          'DtCCta_descripcion_larga
        End If
        If d_aux1_1 = "02" Then
          rstdestino("h_Cta_Larga") = AdoIngresos.Recordset("cta_codigo")
        End If

        rstdestino("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
        rstdestino("H_MontoBs") = IIf(AdoIngresos.Recordset("monto_bolivianos") < 0, (AdoIngresos.Recordset("monto_bolivianos") * -1), AdoIngresos.Recordset("monto_bolivianos"))
        rstdestino("H_MontoDl") = IIf(AdoIngresos.Recordset("monto_dolares") < 0, (AdoIngresos.Recordset("monto_dolares") * -1), AdoIngresos.Recordset("monto_dolares"))
        rstdestino("H_Cambio") = AdoIngresos.Recordset("tipo_cambio")
      End If
      rstdestino("Usr_Usuario") = GlUsuario
      rstdestino("Fecha_registro") = Date
      rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
      '======= fin registra co_diario ==========
    Next I
    '======= inI Actualiza campos de estatus de ingresos ==========
    If rstdestino.State = 1 Then rstdestino.Close
    rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = '" & AdoIngresos.Recordset("correlativo_ingreso") & "' and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' and ges_gestion = '" & AdoIngresos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
    rstdestino.MoveFirst
    If Not (rstdestino.EOF) Then
      rstdestino("estado_aprobacion") = "S"
        If AdoIngresos.Recordset("codigo_tipo") = "DEV" Then
          rstdestino("estado_devengado") = "S"
        End If
        If AdoIngresos.Recordset("codigo_tipo") = "REC" Then
          rstdestino("estado_recaudado") = "S"
        End If
        If AdoIngresos.Recordset("codigo_tipo") = "DYR" Then
          rstdestino("estado_devengado") = "S"
          rstdestino("estado_recaudado") = "S"
        End If
        
        If AdoIngresos.Recordset("codigo_tipo") = "DES" Then
          rstdestino("estado_desafectado") = "S"
        End If
        If AdoIngresos.Recordset("codigo_tipo") = "ANI" Then
          rstdestino("estado_anulado") = "S"
        End If
       rstdestino.Update
       If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza campos de estatus de ingresos ==========
    
    cod_ant = 0
    org_ant = ""
    '======= ini Actualiza el monto recaudado  ==========
    If (AdoIngresos.Recordset("codigo_tipo") = "REC") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = rstdestino("correlativo_anterior")
'        org_ant = rstdestino("org_codigo")
'      End If
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + AdoIngresos.Recordset("monto_dolares")
          rstdestino.Update
      End If
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    
    If (AdoIngresos.Recordset("codigo_tipo") = "DES") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      Print AdoIngresos.Recordset("correlativo_anterior")
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        cod_ant = IIf(IsNull(rstdestino("correlativo_anterior")), 0, rstdestino("correlativo_anterior"))
'        org_ant = rstdestino("org_codigo")
'      End If
      
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DEV" Then 'And AdoIngresos.Recordset("codigo_tipo") = "DES"
          rstdestino!estado_desafectado = "S"
          rstdestino!estado_devengado = "L"
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        Else
          rstdestino("estado_desafectado") = "S"
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - AdoIngresos.Recordset("monto_dolares")
          cod_ant = IIf(IsNull(rstdestino("correlativo_anterior")), 0, rstdestino("correlativo_anterior"))
          org_ant = rstdestino("org_codigo")
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
          rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - AdoIngresos.Recordset("monto_dolares")
          End If
          rstdestino.Update
          If rstdestino.State = 1 Then rstdestino.Close
        End If
      End If
    End If
    
    If (AdoIngresos.Recordset("codigo_tipo") = "ANI") Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("correlativo_anterior") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
        If rstdestino("codigo_tipo") = "DYR" Then
          rstdestino("estado_desafectado") = ""
          rstdestino("estado_recaudado") = "L"
          rstdestino("estado_devengado") = "S"
          rstdestino("estado_anulado") = ""
          rstdestino("codigo_tipo") = "DEV"
          rstdestino("monto_recaudado_dolares") = 0
        End If
      End If
      rstdestino.Update
      If rstdestino.State = 1 Then rstdestino.Close
    End If
    '======= fin Actualiza el monto recaudado  ==========
    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    If AdoIngresos.Recordset("codigo_tipo") = "REC" Or AdoIngresos.Recordset("codigo_tipo") = "DYR" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & AdoIngresos.Recordset("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + AdoIngresos.Recordset("monto_bolivianos")
        rstdestino.Update
      End If
    End If
    If AdoIngresos.Recordset("codigo_tipo") = "ANI" Then
      If rstdestino.State = 1 Then rstdestino.Close
      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & AdoIngresos.Recordset("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
      If Not rstdestino.EOF Then
        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + AdoIngresos.Recordset("monto_bolivianos")
        rstdestino.Update
      End If
    End If
    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
    LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
    Frmmensaje.Visible = False
    db.CommitTrans
  End If
  marca1 = AdoIngresos.Recordset.BookMark
  rstIngresos.Update
  rstIngresos.Requery
  Set AdoIngresos.Recordset = rstIngresos
  If rstIngresos.RecordCount > 0 Then
    AdoIngresos.Recordset.Move marca1 - 1
  End If
  db.Execute "EXEC ts_mf_ActualizaCtaBancaria"
End Sub

Private Sub CmdBorrar_Click()
' ===== Proceso para confirmar el eliminado de registros
  v_añadir = 3
  sino = MsgBox("¿Está seguro de ANULAR este registro?", vbYesNo + vbQuestion, "Atención...")
  If sino = vbYes Then
    Call elimina
    Call errado
  End If
End Sub

Private Sub cmdBuscar_Click()
  Dim ClVBusca As CompBusquedas.ClBuscaEnGridPropio 'Componente de busquedas
  Dim ClBuscaSec As CompBusquedas.ClBuscaSecuencialEnRS
  PosibleApliqueFiltro = False
  Dim RSNADA As ADODB.Recordset
  Dim GrSqlAux As String
  Set ClBuscaGrid = New CompBusquedas.ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = DtGIngresos
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = AdoIngresos.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True
End Sub

'Private Sub CmdBuscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' CmdBuscar.Picture = LoadPicture("d:\Pragma\iconos\busca1.ico")
' CmdBuscar.BackColor = &HC0FFFF
' Image1.Visible = True
 
'End Sub

'Private Sub Cmdbusfin_Click()
'  FrmBuscar.Visible = False
'  FraOpciones.Enabled = True
'End Sub

Private Sub cmdCancelar_Click()
'===== Ini cancela actualizaciones ==========
   FraOpciones2.Visible = False
   FraOpciones.Visible = True
   FraIngresosNav.Enabled = True
   FraIngresosDat.Enabled = False
'   AdoIngresos.Refresh
'  Set AdoIngresos.Recordset = rstIngresos
  rstIngresos.Requery
'  Set DtGIngresos.DataSource = AdoIngresos.Recordset
  LblAccion = ""
End Sub

Private Sub CmdGrabar_Click()
'======= Ini grabado de datos
   swgraba = 0
   Call valida
    
   If swgraba = 1 Then
      FraOpciones2.Visible = False
      FraOpciones.Visible = True
      FraIngresosNav.Enabled = True
      FraIngresosDat.Enabled = False
      
      If v_añadir = 1 Then

         Call add_correl
         Set rstdestino = New ADODB.Recordset
         rstdestino.Open "select * from fo_ingresos order by correlativo_ingreso, org_codigo  ", db, adOpenDynamic, adLockOptimistic
         rstdestino.AddNew
'         lblges_gestion.Caption = "2001"
         rstdestino("Correlativo_ingreso") = correlativo1
         rstdestino("Ges_Gestion") = Trim(lblges_gestion.Caption)
         rstdestino("Codigo_solicitud") = txtCodigo_solicitud.Text
         rstdestino("rbr_codigo") = DtCrbr_codigo.Text
         rstdestino("tipo_moneda") = DtCdenominacion_moneda.BoundText
         rstdestino("UNI_CODIGO") = Txtuni_codigo
         
         Select Case V_accion
            Case "REC"
              rstdestino("Codigo_tipo") = "REC"
              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
              rstdestino("estado_recaudado") = "N"
            Case "DES"
              rstdestino("Codigo_tipo") = "DES"
              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
              rstdestino("estado_desafectado") = "N"
            Case "ANI"
              rstdestino("Codigo_tipo") = "ANI"
              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
              rstdestino("estado_anulado") = "N"
            Case "COPIA"
              rstdestino("Codigo_tipo") = DtCDenominacion_tipo.BoundText
              If DtCDenominacion_tipo.BoundText = "DEV" Then
               rstdestino("estado_devengado") = "N"
               rstdestino("correlativo_anterior") = correlativo1
              End If
              If DtCDenominacion_tipo.BoundText = "REC" Then
               rstdestino("estado_recaudado") = "N"
              End If
              If DtCDenominacion_tipo.BoundText = "DYR" Then
               rstdestino("correlativo_anterior") = correlativo1
               rstdestino("estado_recaudado") = "N"
               rstdestino("estado_devengado") = "N"
              End If

         End Select ' DtCDenominacion_tipo.BoundText
         
         rstdestino("Codigo_tipo_solicitud") = IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
         rstdestino("Codigo_documento") = DtCCodigo_documento.Text
'         DTPFecha_Ingreso.Value = Date
         rstdestino("Fecha_Ingreso") = DTPFecha_Ingreso.Value
         rstdestino("Tipo_Cambio") = txtTipo_Cambio.Text
         rstdestino("Concepto") = (txtConcepto.Text)
         rstdestino("fte_codigo") = DtCFte_codigo.Text
         rstdestino("org_codigo") = DtCOrg_codigo.Text
         
'         rstdestino("cta_codigo") = DtCCta_codigo.Text
         If DtCDenominacion_tipo.BoundText = "DEV" Then
           rstdestino("Codigo_beneficiario") = DtCcodigo_beneficiario.Text
           rstdestino("cta_codigo") = ""
         End If
  
         If DtCDenominacion_tipo.BoundText = "REC" Or DtCDenominacion_tipo.BoundText = "DYR" Or DtCDenominacion_tipo.BoundText = "ANI" Then
           rstdestino("cta_codigo") = DtCCta_codigo.Text
           rstdestino("Codigo_beneficiario") = ""
         End If
         
         rstdestino("numero_documento") = txtNumero_documento.Text
         rstdestino("monto_dolares") = Round(Txtmonto_dolares.Text, 2)
         rstdestino("monto_bolivianos") = Round(txtMonto_bolivianos.Text, 2)
         rstdestino("usr_usuario") = GlUsuario
         rstdestino("fecha_registro") = Date
         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
         
         rstdestino("estado_aprobacion") = "N"
         rstdestino("monto_recaudado_dolares") = 0
         If v_añadir = 1 Then
            rstdestino("ultimo") = "S"
         End If
         rstdestino.Update
         If rstdestino.State = 1 Then rstdestino.Close
          
'          If rstIngresos.State = 1 Then rstIngresos.Close
'          rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'          rstIngresos.Sort = "correlativo_ingreso"
          rstIngresos.Requery
          
'          rstIngresos.Requery
          Set AdoIngresos.Recordset = rstIngresos
          AdoIngresos.Refresh
          AdoIngresos.Recordset.Find "ultimo = 'S'"
          If Not (AdoIngresos.Recordset.EOF) Then
            marca1 = AdoIngresos.Recordset.BookMark
            AdoIngresos.Recordset("ultimo") = "N"
            AdoIngresos.Recordset.Update
          End If
'          rstIngresos.Find "ultimo = 'S'"
'          If Not (rstIngresos.EOF) Then
'            rstIngresos("ultimo") = "N"
'            rstIngresos.Update
'          End If

'          AdoIngresos.Recordset.Move marca1 - 1

'          marca1 = 0
      End If
      
      If v_añadir = 2 Then
        '===== modifica un registro =====
         Set rstdestino = New ADODB.Recordset
         If rstdestino.State = 1 Then rstdestino.Close
         rstdestino.Open "select * from fo_ingresos where correlativo_ingreso = '" & AdoIngresos.Recordset("correlativo_ingreso") & "' and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "' and ges_gestion = '" & AdoIngresos.Recordset("ges_gestion") & "' order by correlativo_ingreso, org_codigo ", db, adOpenDynamic, adLockOptimistic
         rstdestino.MoveFirst
         If Not (rstdestino.EOF) Then
'            If rstdestino("org_codigo") <> DtCOrg_codigo.Text Then
'              Call add_correl
'              rstdestino("Correlativo_ingreso") = correlativo1
'              rstdestino("correlativo_anterior") = correlativo1
'            End If
            rstdestino("Codigo_solicitud") = txtCodigo_solicitud.Text
            rstdestino("rbr_codigo") = DtCrbr_codigo.Text
            rstdestino("tipo_moneda") = DtCdenominacion_moneda.BoundText
            rstdestino("Codigo_tipo_solicitud") = IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
            rstdestino("Codigo_documento") = DtCCodigo_documento.Text
            rstdestino("Codigo_tipo") = IIf(DtCDenominacion_tipo.BoundText = "", "", DtCDenominacion_tipo.BoundText)
            rstdestino("Fecha_Ingreso") = DTPFecha_Ingreso.Value
            rstdestino("Tipo_Cambio") = txtTipo_Cambio.Text
            rstdestino("Concepto") = txtConcepto.Text
            rstdestino("UNI_CODIGO") = Txtuni_codigo
            rstdestino("fte_codigo") = DtCFte_codigo.Text
            rstdestino("org_codigo") = DtCOrg_codigo.Text
            
'            rstdestino("cta_codigo") = DtCCta_codigo.Text
             If DtCDenominacion_tipo.BoundText = "DEV" Then
               rstdestino("Codigo_beneficiario") = DtCcodigo_beneficiario.Text
               rstdestino("cta_codigo") = ""
             End If
      
             If DtCDenominacion_tipo.BoundText = "REC" Or DtCDenominacion_tipo.BoundText = "DYR" Or DtCDenominacion_tipo.BoundText = "ANI" Then
               rstdestino("cta_codigo") = DtCCta_codigo.Text
               rstdestino("Codigo_beneficiario") = ""
             End If

            rstdestino("numero_documento") = txtNumero_documento.Text
            rstdestino("monto_dolares") = Round(Txtmonto_dolares.Text, 2)
            rstdestino("monto_bolivianos") = Round(txtMonto_bolivianos.Text, 2)
            If DtCDenominacion_tipo.BoundText = "DEV" Then
             rstdestino("estado_devengado") = "N"
             rstdestino("estado_recaudado") = ""
             rstdestino("estado_desafectado") = ""
            End If
            If DtCDenominacion_tipo.BoundText = "REC" Then
             rstdestino("estado_recaudado") = "N"
             rstdestino("estado_devengado") = ""
             rstdestino("estado_desafectado") = ""
            End If
            If DtCDenominacion_tipo.BoundText = "DYR" Then
             rstdestino("estado_recaudado") = "N"
             rstdestino("estado_devengado") = "N"
             rstdestino("estado_desafectado") = ""
            End If
            rstdestino("estado_Aprobacion") = "N"
            rstdestino("ultimo") = "N"
            rstdestino("usr_usuario") = GlUsuario
            rstdestino("fecha_registro") = Date
            rstdestino("hora_registro") = Left(CStr(Time()), 8)
            rstdestino.Update
            If rstdestino.State = 1 Then rstdestino.Close
            
            marca1 = AdoIngresos.Recordset.BookMark
            rstIngresos.CancelUpdate
            rstIngresos.Requery
'            rstIngresos.Sort = "correlativo_ingreso"
            Set AdoIngresos.Recordset = rstIngresos
'            AdoIngresos.Refresh
            AdoIngresos.Recordset.Move marca1 - 1
         End If
'         marca1 = 0
      End If
   Else
      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
'      FraOpciones2.Visible = False
'      FraOpciones.Visible = True
'      FraIngresosNav.Enabled = True
'      FraIngresosDat.Enabled = False
'      AdoIngresos.Refresh
   End If
   LblAccion = ""
End Sub

Private Sub cmdImprimir_Click()
If rstIngresos.RecordCount > 0 Then
'===== Ini comando para iniciar impresión
  Dim rstfo_ingresos_rep As New ADODB.Recordset
  Set rstfo_ingresos_rep = New ADODB.Recordset
  Dim iResult As Integer
  '  Cry.Reset
  Cry.ReportFileName = App.path & "\FormsIngresos\ComprobIngreso.rpt"
'  Cry.SelectionFormula = "{fv_comprobante2.Maquina} = '" & GlMaquina & "'"
  If rstfo_ingresos_rep.State = 1 Then rstfo_ingresos_rep.Close
  rstfo_ingresos_rep.Open "select * from fo_ingresos_rep where maquina = '" & GlMaquina & "'", db, adOpenKeyset, adLockOptimistic
  While Not (rstfo_ingresos_rep.EOF)
    rstfo_ingresos_rep.Delete
    rstfo_ingresos_rep.MoveNext
  Wend
  '====== ini cargado de la tabla aux para impresion ====
  rstfo_ingresos_rep.AddNew
  rstfo_ingresos_rep("Correlativo_ingreso") = LblCorrelativo_ingreso.Caption
  rstfo_ingresos_rep("Correlativo_anterior") = AdoIngresos.Recordset("correlativo_anterior")
  rstfo_ingresos_rep("Ges_Gestion") = Trim(lblges_gestion.Caption) ' TxtGes_Gestion.Text
  rstfo_ingresos_rep("Codigo_solicitud") = txtCodigo_solicitud.Text
  rstfo_ingresos_rep("rbr_codigo") = DtCrbr_codigo.Text
  rstfo_ingresos_rep("tipo_moneda") = DtCdenominacion_moneda.BoundText
  rstfo_ingresos_rep("Codigo_tipo") = DtCDenominacion_tipo.BoundText
  rstfo_ingresos_rep("Codigo_tipo_solicitud") = IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
  rstfo_ingresos_rep("Codigo_documento") = DtCCodigo_documento.Text
  rstfo_ingresos_rep("Fecha_Ingreso") = DTPFecha_Ingreso.Value
  rstfo_ingresos_rep("Tipo_Cambio") = txtTipo_Cambio.Text
  rstfo_ingresos_rep("Concepto") = txtConcepto.Text
  rstfo_ingresos_rep("UNI_CODIGO") = Txtuni_codigo
  rstfo_ingresos_rep("fte_codigo") = DtCFte_codigo.Text
  rstfo_ingresos_rep("org_codigo") = DtCOrg_codigo.Text

  If AdoIngresos.Recordset("Codigo_tipo") = "DEV" Then
    rstfo_ingresos_rep("codigo_beneficiario") = DtCcodigo_beneficiario.Text
  End If

  If AdoIngresos.Recordset("Codigo_tipo") = "DYR" Or AdoIngresos.Recordset("Codigo_tipo") = "REC" Or AdoIngresos.Recordset("Codigo_tipo") = "ANI" Then
    rstfo_ingresos_rep("Cta_codigo") = DtCCta_codigo.Text
  End If

'  rstfo_ingresos_rep("cta_codigo") = DtCCta_codigo.Text
  
  rstfo_ingresos_rep("numero_documento") = txtNumero_documento.Text
  rstfo_ingresos_rep("monto_dolares") = Round(Txtmonto_dolares.Text, 2)
  rstfo_ingresos_rep("monto_bolivianos") = Round(txtMonto_bolivianos.Text, 2)
  rstfo_ingresos_rep("usr_usuario") = GlUsuario
  rstfo_ingresos_rep("fecha_registro") = Date
  rstfo_ingresos_rep("hora_registro") = Left(CStr(Time()), 8)
  rstfo_ingresos_rep("estado_recaudado") = IIf(AdoIngresos.Recordset("estado_recaudado") = "S", "A", IIf(AdoIngresos.Recordset("estado_recaudado") = "N", "S", ""))
  rstfo_ingresos_rep("estado_devengado") = IIf(AdoIngresos.Recordset("estado_devengado") = "S", "A", IIf(AdoIngresos.Recordset("estado_devengado") = "N", "S", ""))
  rstfo_ingresos_rep("estado_desafectado") = IIf(AdoIngresos.Recordset("estado_desafectado") = "S", "A", IIf(AdoIngresos.Recordset("estado_desafectado") = "N", "S", ""))
  rstfo_ingresos_rep("estado_anulado") = IIf(AdoIngresos.Recordset("estado_anulado") = "S", "A", IIf(AdoIngresos.Recordset("estado_anulado") = "N", "S", ""))
  rstfo_ingresos_rep("estado_aprobacion") = AdoIngresos.Recordset("estado_aprobacion")
  rstfo_ingresos_rep("maquina") = GlMaquina
  rstfo_ingresos_rep.Update
  If rstfo_ingresos_rep.State = 1 Then rstfo_ingresos_rep.Close
  '====== fin cargado de la tabla aux para impresion ====
  
  Cry.SelectionFormula = "{Vi_Fo_ingresos_rep.Maquina} = '" & GlMaquina & "'"
  Cry.WindowShowPrintBtn = True
  Cry.WindowShowExportBtn = True
  Cry.WindowShowPrintSetupBtn = True
  Cry.WindowState = crptMaximized
  iResult = Cry.PrintReport
  If iResult <> 0 Then
      MsgBox Cry.LastErrorNumber & " : " & Cry.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
Else
  MsgBox "No existen registros para imprimir", vbInformation + vbOKOnly, "ERROR de impresión"
End If
End Sub

'Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  CmdImprimir.Picture = LoadPicture("d:\Pragma\iconos\print2.ico") 'LoadPicture("C:\Mis documentos\star.ani")
''  CmdImprimir.BackColor = &HC0FFFF
'
'End Sub

Private Sub CmdModificar_Click()
    LblAccion = "Modificando registro..."
    v_añadir = 2
    FraOpciones.Visible = False
    FraOpciones2.Visible = True
    FraIngresosNav.Enabled = False
    FraIngresosDat.Enabled = True
    DtCFte_codigo.Enabled = False
    DtCOrg_codigo.Enabled = False
    swmodificar = 1
    Call pfil_cta_Fte(DtCFte_codigo.Text, 1)
    If AdoIngresos.Recordset("Codigo_tipo") = "REC" Or AdoIngresos.Recordset("Codigo_tipo") = "DYR" Or AdoIngresos.Recordset("Codigo_tipo") = "ANI" Then
      DtCCta_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("Cta_Codigo")) = True, "", AdoIngresos.Recordset("Cta_Codigo"))
      DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
    End If
    
    If swcopiar = 1 Then
'      If marca1 = 0 Then
         marca1 = AdoIngresos.Recordset.BookMark
      Else
        marca1 = AdoIngresos.Recordset.BookMark
'        Set AdoIngresos.Recordset = rstIngresos
'        AdoIngresos.Refresh
'        AdoIngresos.Recordset.Move marca1 - 1
      End If
'    Else
'      marca1 = AdoIngresos.Recordset.Bookmark
'    End If
    
'    If V_accion = "COPIA" Then
'      If marca1 = 0 Then
'         marca1 = AdoIngresos.Recordset.Bookmark
'      Else
'        Set AdoIngresos.Recordset = rstIngresos
'        AdoIngresos.Refresh
'        AdoIngresos.Recordset.Move marca1 - 1
'      End If
'    Else
'      marca1 = AdoIngresos.Recordset.Bookmark
'    End If
    
    correlativo_ingreso1 = AdoIngresos.Recordset("correlativo_ingreso")
    ges_gestion1 = AdoIngresos.Recordset("ges_gestion")
End Sub

Private Sub cmdSalir_Click()
   sino = MsgBox("¿Está seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
   If sino = vbYes Then
     Call cerrar
  
  If rstFte_financia.State = 1 Then rstFte_financia.Close
  If AdoFte_financia.Recordset.State = 1 Then AdoFte_financia.Recordset.Close
  If rstIngresos.RecordCount > 0 Then
    rstIngresos.Update
  End If
  If rstIngresos.State = 1 Then rstIngresos.Close
     Unload Me
   End If
End Sub

Private Sub CommandButton1_Click()

End Sub

'Private Sub Cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Cmdsalir.Picture = LoadPicture("d:\Pragma\iconos\salir1.ico")
'End Sub

Private Sub DtCcodigo_beneficiario_Click(Area As Integer)
   DtCdenominacion_beneficiario.Text = DtCcodigo_beneficiario.BoundText
End Sub

Private Sub DtCCodigo_documento_Click(Area As Integer)
    DtCDenominacion_documento.Text = DtCCodigo_documento.BoundText
End Sub

Private Sub DtCCta_codigo_Click(Area As Integer)
   DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
End Sub

Private Sub DtCCta_descripcion_larga_Click(Area As Integer)
   DtCCta_codigo.Text = DtCCta_descripcion_larga.BoundText
End Sub

Private Sub DtCdenominacion_beneficiario_Click(Area As Integer)
  DtCcodigo_beneficiario.Text = DtCdenominacion_beneficiario.BoundText
End Sub

Private Sub DtCDenominacion_documento_Click(Area As Integer)
    DtCCodigo_documento = DtCDenominacion_documento.BoundText
End Sub

Private Sub DtCDenominacion_tipo_Click(Area As Integer)

  If DtCDenominacion_tipo = "DEVENGADO" Then
    DtCCta_codigo.Visible = False
    DtCCta_descripcion_larga.Visible = False
    lblcuenta.Visible = False
    DtCcodigo_beneficiario.Visible = True
    DtCdenominacion_beneficiario.Visible = True
    lblBeneficiario.Visible = True
  End If

  If (DtCDenominacion_tipo = "DEVENGADO Y RECAUDADO") Or (DtCDenominacion_tipo = "RECAUDADO") Or (DtCDenominacion_tipo = "ANULADO") Then
    DtCCta_codigo.Visible = True
    DtCCta_descripcion_larga.Visible = True
    lblcuenta.Visible = True
    DtCcodigo_beneficiario.Visible = False
    DtCdenominacion_beneficiario.Visible = False
    lblBeneficiario.Visible = False
  End If

'3 codigo_beneficiario varchar 15  0 0 0   0     0
'0 denominacion_beneficiario varchar 60  0 0 1   0     0
'0 tipo_beneficiario varchar 1 0 0 1   0     0

End Sub

Private Sub DtCDenominacion_tipo_solicitud_KeyPress(KeyAscii As Integer)
  ' aqui cambiar de lugar
  If KeyAscii = 13 Then
    
  End If
End Sub

Private Sub DtCOrg_codigo_Click(Area As Integer)
   DtCOrg_descripcion.Text = DtCOrg_codigo.BoundText
    Call pfil_cta_Fte(Me.DtCOrg_codigo, 2)
End Sub

Private Sub DtCOrg_descripcion_Click(Area As Integer)
   DtCOrg_codigo.Text = DtCOrg_descripcion.BoundText
    Call pfil_cta_Fte(Me.DtCOrg_codigo, 2)
End Sub

Private Sub DtCrbr_codigo_Click(Area As Integer)
   DtCrbr_descripcion.Text = DtCrbr_codigo.BoundText
End Sub

Private Sub DtCrbr_descripcion_Click(Area As Integer)
    DtCrbr_codigo.Text = DtCrbr_descripcion.BoundText
End Sub

Private Sub DtCFte_codigo_Click(Area As Integer)
    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
    DtCOrg_codigo.Enabled = True
    Call pfil_Org_Fte(DtCFte_codigo.Text)
    Call pfil_cta_Fte(Me.DtCFte_codigo, 1)
End Sub

Private Sub DtCFte_descripcion_larga_Click(Area As Integer)
  DtCFte_codigo.Text = DtCFte_descripcion_larga.BoundText
  Call pfil_Org_Fte(DtCFte_descripcion_larga.BoundText)
  Call pfil_cta_Fte(DtCFte_descripcion_larga.BoundText, 1)
End Sub

Private Sub Form_Load()
  '===== Ini cargado de tablas de consulta y de datos de despliegue
  LblUsuario.Caption = LblUsuario.Caption + GlUsuario
  swgraba = 0
  marca1 = 0
  swcopiar = 0
  V_accion = "COPIA"
  
  Set rstfc_beneficiario = New ADODB.Recordset
  If rstfc_beneficiario.State = 1 Then rstfc_beneficiario.Close
  rstfc_beneficiario.Open "SELECT * from Fc_beneficiario order by codigo_beneficiario", db, adOpenStatic, adLockReadOnly
  Set Adofc_beneficiario.Recordset = rstfc_beneficiario
  Adofc_beneficiario.Refresh
  
  Set rstFc_Rubro_ingresos = New ADODB.Recordset
  If rstFc_Rubro_ingresos.State = 1 Then rstFc_Rubro_ingresos.Close
  rstFc_Rubro_ingresos.Open "select * from Fc_Rubro_ingresos order by rbr_codigo", db, adOpenKeyset, adLockReadOnly
  Set AdoFc_Rubro_ingresos.Recordset = rstFc_Rubro_ingresos
  AdoFc_Rubro_ingresos.Refresh
  If Not AdoFc_Rubro_ingresos.Recordset.BOF Then AdoFc_Rubro_ingresos.Recordset.MoveFirst
  
  Set rstTipo_moneda = New ADODB.Recordset
  If rstTipo_moneda.State = 1 Then rstTipo_moneda.Close
  rstTipo_moneda.Open "select * from Tipo_moneda order by denominacion_moneda", db, adOpenKeyset, adLockReadOnly
  Set Adotipo_moneda.Recordset = rstTipo_moneda
  Adotipo_moneda.Refresh
  If Not Adotipo_moneda.Recordset.BOF Then Adotipo_moneda.Recordset.MoveFirst
  
  Set rstTipo_comprobante = New ADODB.Recordset
  If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
  rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
  Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
  AdoTipo_comprobante.Refresh
  If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
  
  Set rstTipo_solicitud = New ADODB.Recordset
  If rstTipo_solicitud.State = 1 Then rstTipo_solicitud.Close
  rstTipo_solicitud.Open "select * from Tipo_solicitud order by Denominacion_tipo_solicitud", db, adOpenKeyset, adLockReadOnly
  Set AdoTipo_solicitud.Recordset = rstTipo_solicitud
  AdoTipo_solicitud.Refresh
  If Not AdoTipo_solicitud.Recordset.BOF Then AdoTipo_solicitud.Recordset.MoveFirst
  
  
  If rstFte_financia.State = 1 Then rstFte_financia.Close
  rstFte_financia.Open "Select * from Fc_fuente_financiamiento", db, adOpenDynamic, adLockReadOnly
  Set AdoFte_financia.Recordset = rstFte_financia
  AdoFte_financia.Refresh
  If Not AdoFte_financia.Recordset.BOF Then AdoFte_financia.Recordset.MoveFirst
  
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento", db, adOpenDynamic, adLockReadOnly
  Set AdoOrganismo_finan.Recordset = rstOrganismo_finan
  AdoOrganismo_finan.Refresh
  If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
  
  If rstFc_cuenta_bancaria.State = 1 Then rstFc_cuenta_bancaria.Close
  rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria", db, adOpenDynamic, adLockReadOnly
  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
  AdoFc_cuenta_bancaria.Refresh
  If Not AdoFc_cuenta_bancaria.Recordset.BOF Then AdoFc_cuenta_bancaria.Recordset.MoveFirst
  
  If rstac_documento_respaldo.State = 1 Then rstac_documento_respaldo.Close
  Set rstac_documento_respaldo = New ADODB.Recordset
  rstac_documento_respaldo.Open "select * from ac_documento_respaldo", db, adOpenDynamic, adLockReadOnly
  Set Adoac_documento_respaldo.Recordset = rstac_documento_respaldo
  Adoac_documento_respaldo.Refresh
  If Not Adoac_documento_respaldo.Recordset.BOF Then Adoac_documento_respaldo.Recordset.MoveFirst
  
  Set rstIngresos = New ADODB.Recordset
  ' pa busqueda QueryInicial = "select * from fo_ingresos where estado_aprobacion <> 'S'" 'ORDER BY correlativo_ingreso , org_codigo
  queryinicial = "select * from fo_ingresos where estado_aprobacion <> 'S' and estado_aprobacion <> 'E'" ' ORDER BY correlativo_ingreso , org_codigo"
  If rstIngresos.State = 1 Then rstIngresos.Close
'pa busqueda  rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
  rstIngresos.Open queryinicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic
'pa busqueda  rstIngresos.Sort = "correlativo_ingreso"
  Set AdoIngresos.Recordset = rstIngresos
  
  If (Not AdoIngresos.Recordset.BOF) And (Not AdoIngresos.Recordset.EOF) Then
    AdoIngresos.Recordset.MoveFirst
    DtCFte_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("fte_codigo")) = True, " ", AdoIngresos.Recordset("fte_codigo"))
    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
    DtCOrg_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("org_codigo")) = True, " ", AdoIngresos.Recordset("org_codigo"))
    DtCOrg_descripcion.Text = DtCOrg_codigo.BoundText
    DtCCta_codigo.Text = IIf(IsNull(AdoIngresos.Recordset("Cta_Codigo")) = True, " ", AdoIngresos.Recordset("Cta_Codigo"))
    DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
  End If
  '===== fin cargado de tablas de consulta y de datos de despliegue
  txtTipo_Cambio = GlTipoCambioOficial
End Sub

'Private Sub FraOpciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  CmdImprimir.Picture = LoadPicture("d:\Pragma\iconos\print3.ico") 'LoadPicture("C:\Mis documentos\star.ani")
'  CmdBuscar.Picture = LoadPicture("d:\Pragma\iconos\busca3.ico") 'LoadPicture("C:\Mis documentos\star.ani")
'  CmdSalir.Picture = LoadPicture("d:\Pragma\iconos\salir3.ico")
''  CmdBuscar.BackColor = &H8000000F
''  CmdImprimir.BackColor = &H8000000F
''  CmdAñadir.BackColor = &H8000000F
''  Image1.Visible = False
'
'End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros aprobados)
  queryinicial = "select * from fo_ingresos where estado_aprobacion <> 'S' and estado_aprobacion <> 'E'"
  If rstIngresos.State = 1 Then rstIngresos.CancelUpdate
  If rstIngresos.State = 1 Then rstIngresos.Close
  
'pa busqueda  rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
  rstIngresos.Open queryinicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic
'pa busqueda  rstIngresos.Sort = "correlativo_ingreso"
  
'  rstIngresos.Open QueryInicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic ' ORDER BY correlativo_ingreso , org_codigo "
  rstIngresos.Requery
'dul
Set AdoIngresos.Recordset = rstIngresos
'rstIngresos.Requery
'dul  Set DtGIngresos.DataSource = AdoIngresos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
  If rstIngresos.RecordCount > 0 Then rstIngresos.CancelUpdate
  If rstIngresos.State = 1 Then rstIngresos.Close
  queryinicial = "select * from fo_ingresos "
  
'pa busqueda  rstIngresos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
  rstIngresos.Open queryinicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic
'pa busqueda  rstIngresos.Sort = "correlativo_ingreso"
  
'  rstIngresos.Open QueryInicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic 'ORDER BY correlativo_ingreso , org_codigo
  rstIngresos.Requery
  Set AdoIngresos.Recordset = rstIngresos

'dul  Set AdoIngresos.Recordset = rstIngresos
  'rstIngresos.Requery
'dul  Set DtGIngresos.DataSource = AdoIngresos.Recordset
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub abrir()
  If rstIngresos.State = 1 Then rstIngresos.Close
  Set rstIngresos = New ADODB.Recordset
  rstIngresos.Open "select * from fo_ingresos order by correlativo_ingreso, org_codigo ", db, adOpenDynamic, adLockOptimistic
  Set AdoIngresos.Recordset = rstIngresos
  If AdoIngresos.Recordset.State = 1 Then AdoIngresos.Recordset.Close
  AdoIngresos.Refresh
  DtGIngresos.Refresh
  If Not rstIngresos.BOF Then rstIngresos.MoveFirst
  
  If rstFte_financia.State = 1 Then rstFte_financia.Close
  rstFte_financia.Open "Select * from Fc_fuente_financiamiento", db, adOpenDynamic, adLockReadOnly
  Set AdoFte_financia.Recordset = rstFte_financia
  AdoFte_financia.Refresh
  If Not rstFte_financia.BOF Then rstFte_financia.MoveFirst
  
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento", db, adOpenDynamic, adLockReadOnly
  Set AdoOrganismo_finan.Recordset = rstOrganismo_finan
  AdoOrganismo_finan.Refresh
  If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
  
  If rstFc_cuenta_bancaria.State = 1 Then rstFc_cuenta_bancaria.Close
  rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria", db, adOpenDynamic, adLockReadOnly
  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
  AdoFc_cuenta_bancaria.Refresh
  If Not rstFc_cuenta_bancaria.BOF Then rstFc_cuenta_bancaria.MoveFirst

  If (Not AdoIngresos.Recordset.BOF) And (Not AdoIngresos.Recordset.EOF) Then
  
  End If
End Sub

Sub cerrar()
'  If rstFte_financia.State = 1 Then rstFte_financia.Close
'  If AdoFte_financia.Recordset.State = 1 Then AdoFte_financia.Recordset.Close
'  If AdoIngresos.Recordset.State = 1 Then AdoIngresos.Recordset.Close
'  If rstIngresos.State = 1 Then rstIngresos.Close
End Sub

Private Sub Txtduracion_estimada_KeyPress(KeyAscii As Integer)
  If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub elimina()
'===== proceso para eliminar registros
  Dim rstelimina As New ADODB.Recordset
  If rstelimina.State = 1 Then rstelimina.Close
  Set rstelimina = New ADODB.Recordset
  If rstelimina.State = 1 Then rstelimina.Close
  marca1 = AdoIngresos.Recordset.BookMark
  rstelimina.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("Correlativo_ingreso") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
  If (Not rstelimina.BOF) Then rstelimina.MoveFirst
  '  rstIngresos.Find "Correlativo_ingreso= '" & AdoIngresos.Recordset("Correlativo_ingreso") & "'", , adSearchForward
  '  If Not rstIngresos.BOF Then
    If rstelimina!estado_devengado = "S" Then rstelimina!estado_devengado = "E"
    If rstelimina!estado_recaudado = "S" Then rstelimina!estado_recaudado = "E"
    If rstelimina!estado_desafectado = "S" Then rstelimina!estado_desafectado = "E"
    If rstelimina!estado_anulado = "S" Then rstelimina!estado_anulado = "E"
    rstelimina!estado_aprobacion = "E"
    rstelimina.Update
  '  End If
  If rstelimina.State = 1 Then rstelimina.Close
  rstIngresos.Update
  rstIngresos.Requery
  Set AdoIngresos.Recordset = rstIngresos
  AdoIngresos.Refresh
  Set AdoIngresos.Recordset = rstIngresos
  If rstIngresos.RecordCount > 0 Then
    AdoIngresos.Recordset.Move marca1 - 1
  End If
End Sub

Private Sub errado()
'===== proceso para eliminar registros
  Dim rsterrado As New ADODB.Recordset
  If rsterrado.State = 1 Then rsterrado.Close
  Set rsterrado = New ADODB.Recordset
  rsterrado.Open "select * from fo_ingresos where correlativo_ingreso = " & AdoIngresos.Recordset("Correlativo_ingreso") & " and org_codigo = '" & AdoIngresos.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
  If (Not rsterrado.BOF) Then rsterrado.MoveFirst
'  rsterrado.Find "Correlativo_ingreso= '" & AdoIngresos.Recordset("Correlativo_ingreso") & "'", , adSearchForward
'  If Not rsterrado.BOF Then
    If rsterrado("estado_devengado") = "N" Then
      rsterrado("estado_devengado") = "E"
      rsterrado("estado_aprobacion") = "E"
    End If
    If rsterrado("estado_recaudado") = "N" Then
      rsterrado("estado_recaudado") = "E"
      rsterrado("estado_aprobacion") = "E"
    End If
    If rsterrado("estado_desafectado") = "N" Then
      rsterrado("estado_desafectado") = "E"
      rsterrado("estado_aprobacion") = "E"
    End If
    
    rsterrado.Update
'  End If
  If rsterrado.State = 1 Then rsterrado.Close
  rstIngresos.Update
  rstIngresos.Requery
  Set AdoIngresos.Recordset = rstIngresos
  AdoIngresos.Refresh
End Sub

Private Sub valida()
'===== Validación para grabar datos
  swgraba = 1
  If Len(Trim(txtCodigo_solicitud)) < 1 Then swgraba = 0
  If IsNull(DTPFecha_Ingreso) Then swgraba = 0
  If txtTipo_Cambio = 0 Then swgraba = 0
  If Len(Trim(txtConcepto)) < 1 Then swgraba = 0
  If Len(Trim(Txtmonto_dolares)) < 1 Then swgraba = 0
  If Len(Trim(txtMonto_bolivianos.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCrbr_codigo.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCdenominacion_moneda.Text)) < 1 Then swgraba = 0
  If Len(Trim(Txtuni_codigo.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCDenominacion_tipo_solicitud.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCCodigo_documento.Text)) < 1 Then swgraba = 0
  If Len(Trim(txtConcepto.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCFte_codigo.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCOrg_codigo.Text)) < 1 Then swgraba = 0
  If DtCDenominacion_tipo.BoundText = "DEV" Then
    If (Len(Trim(DtCcodigo_beneficiario.Text)) < 1) Then swgraba = 0
  End If
  If (DtCDenominacion_tipo.BoundText = "REC") Or (DtCDenominacion_tipo.BoundText = "DYR") Or (DtCDenominacion_tipo.BoundText = "ANI") Then
    If (Len(Trim(DtCCta_codigo.Text)) < 1) Then swgraba = 0
  End If
  If Len(Trim(txtNumero_documento.Text)) < 1 Then swgraba = 0

End Sub

Private Sub TxtCodigo_beneficiario_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TDBNumber1_Click()

End Sub

Private Sub TxtCodigo_solicitud_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtConcepto_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
  If Len(txtTipo_Cambio.Text) > 0 Then
    If (Len(Trim(txtMonto_bolivianos.Text)) > 0) Then
       Txtmonto_dolares.Text = IIf(txtMonto_bolivianos.Text > 0, Round(txtMonto_bolivianos.Text / txtTipo_Cambio, 2), 0)
    Else
       Txtmonto_dolares.Text = 0
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
  If Len(txtTipo_Cambio.Text) > 0 Then
    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
      txtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Round(Txtmonto_dolares * txtTipo_Cambio, 2), 0)
    Else
      txtMonto_bolivianos.Text = 0
    End If
  End If
End Sub

Private Sub TxtNumero_documento_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtTipo_moneda_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtTipo_solicitud_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub pfil_Org_Fte(Codfte As String)
'===== Proceso para filtrar los Organismos en base a la Fuente de financiamiento
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento where fte_codigo = '" & Codfte & "'", db, adOpenDynamic, adLockReadOnly
  'If rstOrganismo_finan.RecordCount < 1 Then
    DtCOrg_codigo.Text = ""
    DtCOrg_descripcion.Text = " "
  'End If
  Set AdoOrganismo_finan.Recordset = rstOrganismo_finan
  AdoOrganismo_finan.Refresh
  If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
End Sub

Private Sub pfil_cta_Fte(cod, que)
'aqui ver las cuentas
  If rstFc_cuenta_bancaria.State = 1 Then rstFc_cuenta_bancaria.Close
  Select Case que
    Case 1
      rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria where fte_codigo = '" & cod & "' or cta_codigo = '01' ", db, adOpenDynamic, adLockReadOnly
    Case 2
      rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria where org_codigo = '" & cod & "' or cta_codigo = '01' ", db, adOpenDynamic, adLockReadOnly
  End Select
  Me.DtCCta_codigo.Text = ""
  Me.DtCCta_descripcion_larga.Text = ""
  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
  AdoFc_cuenta_bancaria.Refresh
  If Not AdoFc_cuenta_bancaria.Recordset.BOF Then AdoFc_cuenta_bancaria.Recordset.MoveFirst
End Sub

'Private Sub Cmdbuspri_Click()
''===== Proceso para buscar el primer registro en base al criterio seleccionado
'  Call parametros
'  If buscasi = 1 Then
'    If (Not AdoIngresos.Recordset.BOF) Then AdoIngresos.Recordset.MoveFirst
'    If operadorbus = "=" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchForward
'      If AdoIngresos.Recordset.EOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'    If operadorbus = "like" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchForward
'      If AdoIngresos.Recordset.EOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'  End If
'  buscasi = 0
'End Sub

'Private Sub Cmdbussig_Click()
''===== Proceso para buscar el siguiente registro en base al criterio seleccionado
'  Call parametros
'  If buscasi = 1 Then
'    If (Not AdoIngresos.Recordset.EOF) Then AdoIngresos.Recordset.MoveNext
'    If operadorbus = "=" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchForward
'      If AdoIngresos.Recordset.EOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'    If operadorbus = "like" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchForward
'      If AdoIngresos.Recordset.EOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'  End If
'  buscasi = 0
'End Sub
'
'Private Sub Cmdbusult_Click()
''===== Proceso para buscar el último registro en base al criterio seleccionado
'  Call parametros
'  If buscasi = 1 Then
'    If (Not AdoIngresos.Recordset.EOF) Then AdoIngresos.Recordset.MoveLast
'    If operadorbus = "=" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchBackward
'      If AdoIngresos.Recordset.BOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'    If operadorbus = "like" Then
'      AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchBackward
'      If AdoIngresos.Recordset.EOF Then
'        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'        AdoIngresos.Recordset.MoveFirst
'      End If
'    End If
'  End If
'  buscasi = 0
'End Sub
'
'Private Sub CmdBusAnt_Click()
''===== Proceso para buscar el anterior registro en base al criterio seleccionado
'  Call parametros
'  If buscasi = 1 Then
'    If (Not AdoIngresos.Recordset.BOF) Then AdoIngresos.Recordset.MovePrevious
'    If (Not AdoIngresos.Recordset.BOF) Then
'      If operadorbus = "=" Then
'        AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchBackward
'        If AdoIngresos.Recordset.BOF Then
'          MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'          AdoIngresos.Recordset.MoveFirst
'        End If
'      End If
'      If operadorbus = "like" Then
'        AdoIngresos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchBackward
'        If AdoIngresos.Recordset.BOF Then
'          MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
'          AdoIngresos.Recordset.MoveFirst
'        End If
'      End If
'    Else
'      MsgBox "Este es el primer registro", vbCritical + vbOKOnly, "Inicio de Registros"
'      AdoIngresos.Recordset.MoveFirst
'    End If
'  End If
'  buscasi = 0
'End Sub

'Private Sub parametros()
''===== Proceso para definir los criterios de búsqueda
'  buscasi = 1
'  If Len(Trim(Cmbcampobus.Text)) < 1 Then buscasi = 0
'  If Len(Trim(CmbOperador.Text)) < 1 Then buscasi = 0
'  If Len(Trim(Txtvarbus.Text)) < 1 Then buscasi = 0
'  If buscasi = 1 Then
'    Select Case Trim(Cmbcampobus.Text)
'      Case "Comprobante"
'        campobus = " correlativo_ingreso "
'      Case "Organismo Finan."
'        campobus = " org_codigo "
'      Case "Cuenta"
'        campobus = " cta_codigo "
'      Case "Fecha Ingreso"
'        campobus = " fecha_ingreso "
'        CmbOperador.Text = "="
'      Case "No.Solicitud Desembolso"
'        campobus = " codigo_solicitud "
'      Case Else
'    End Select
'
'    Select Case Trim(CmbOperador.Text)
'      Case "="
'        operadorbus = "="
'      Case "PARTE ="
'        operadorbus = "like"
'      Case Else
'    End Select
'  Else
'    MsgBox "Para poder realizar la búsqueda, por favor debe ingresar todos los parámetros ", vbCritical + vbOKOnly, "ERROR en búsqueda"
'  End If
'End Sub

'Private Sub DtGIngresos_Click()
'    TIPOFORMULARIO = DtcTipoDes.Text
'End Sub

Private Sub DtGIngresos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = vbRightButton Then Me.PopupMenu mnuAcciones
End Sub

Private Sub mnuAccion_Click(Index As Integer)
  correlativo_ingreso1 = AdoIngresos.Recordset("correlativo_ingreso")
  Org_Codigo1 = AdoIngresos.Recordset("org_codigo")
  Select Case Index
    Case 0 ' RECAUDADO      ' Devengado
      'if AdoIngresos.Recordset("estado_reversion_total")="S" then
      MsgBox "Realizando el RECAUDADO", vbInformation + vbOKOnly, "Atención"
      V_accion = "REC"
      CmdCopiar_Click
    Case 1 'DESAFECTADO     ' Reversión
      MsgBox "Realizando la Desafección", vbInformation + vbOKOnly, "Atención"
      V_accion = "DES"
      CmdCopiar_Click
    Case 2 'ANULAR RECAUDADO       ' Devolución
      MsgBox "Realizando la Anulación de lo Recaudado", vbInformation + vbOKOnly, "Atención"
      V_accion = "ANI"
      CmdCopiar_Click
  End Select
End Sub

Private Sub CmdCopiar_Click()
  v_añadir = 1
  swcopiar = 1
  CmdAñadir_Click
'  CmdGrabar_Click
'  CmdModificar_Click

'last  V_accion = "COPIA"
  swcopiar = 0
End Sub

Private Sub desactivar_Obj()
'  TxtTipo_cambio.Enabled = False
  txtConcepto.Enabled = False
  Txtmonto_dolares.Enabled = False
  txtMonto_bolivianos.Enabled = False
  DtCrbr_codigo.Enabled = False
  DtCrbr_descripcion.Enabled = False
  DtCdenominacion_moneda.Enabled = False
  DtCDenominacion_tipo_solicitud.Enabled = False
  DtCCodigo_documento.Enabled = False
  DtCDenominacion_documento.Enabled = False
  DtCFte_codigo.Enabled = False
  DtCFte_descripcion_larga.Enabled = False
  DtCOrg_codigo.Enabled = False
  DtCOrg_descripcion.Enabled = False
  DtCCta_codigo.Enabled = False
  DtCCta_descripcion_larga.Enabled = False
  DtCDenominacion_tipo.Enabled = False
  
  txtNumero_documento.Enabled = False
  txtCodigo_solicitud.Enabled = False
  
  Select Case AdoIngresos.Recordset("codigo_tipo")
    Case "REC"
      DtCCta_codigo.Enabled = True
      DtCCta_descripcion_larga.Enabled = True
      
'      TxtTipo_cambio.Enabled = True
      
      DtCdenominacion_moneda.Enabled = True
      Txtmonto_dolares.Enabled = True
      txtMonto_bolivianos.Enabled = True
  End Select
  
  cmdCopiar.Enabled = False
  
End Sub

Private Sub activar_Obj()
  DtCDenominacion_tipo.Enabled = True
  cmdCopiar.Enabled = True

  txtCodigo_solicitud.Enabled = True
'  DTPFecha_Ingreso.Enabled = True
'  TxtTipo_cambio.Enabled = True
  txtConcepto.Enabled = True
  Txtmonto_dolares.Enabled = True
  txtMonto_bolivianos.Enabled = True
  DtCrbr_codigo.Enabled = True
  DtCrbr_descripcion.Enabled = True
  DtCdenominacion_moneda.Enabled = True
  DtCDenominacion_tipo_solicitud.Enabled = True
  DtCCodigo_documento.Enabled = True
  DtCDenominacion_documento.Enabled = True
  DtCFte_codigo.Enabled = True
  DtCFte_descripcion_larga.Enabled = True
  DtCOrg_codigo.Enabled = True
  DtCOrg_descripcion.Enabled = True
  DtCCta_codigo.Enabled = True
  DtCCta_descripcion_larga.Enabled = True
  DtCdenominacion_moneda.Enabled = True
  txtNumero_documento.Enabled = True
  txtCodigo_solicitud.Enabled = True
  If swcopiar = 1 Then
    DtCFte_codigo.Enabled = False
    DtCOrg_codigo.Enabled = False
  Else
    DtCFte_codigo.Enabled = True
    DtCOrg_codigo.Enabled = True
  End If
  If swmodificar = 1 Then
    DtCFte_codigo.Enabled = False
    DtCOrg_codigo.Enabled = False
  End If
End Sub

'
Private Sub add_correl()
  Dim rstcorrel_ing As New ADODB.Recordset
  Set rstcorrel_ing = New ADODB.Recordset
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
  rstcorrel_ing.Open "select * from fc_correl_ingresos where org_codigo = '" & Trim(DtCOrg_codigo.Text) & "' and ges_gestion = '" & Trim(lblges_gestion.Caption) & "'", db, adOpenDynamic, adLockOptimistic
  If Not (rstcorrel_ing.BOF) Then rstcorrel_ing.MoveFirst
  rstcorrel_ing.Find "org_codigo = '" & (DtCOrg_codigo.Text) & "' ", , adSearchForward
  If rstcorrel_ing.EOF Then
     rstcorrel_ing.AddNew
     rstcorrel_ing("org_codigo") = Trim(DtCOrg_codigo.Text)
     rstcorrel_ing("ges_gestion") = Trim(lblges_gestion.Caption)
     rstcorrel_ing("correlativo") = 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo")
     FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  Else
     rstcorrel_ing("correlativo") = rstcorrel_ing("correlativo") + 1
     rstcorrel_ing.Update
     correlativo1 = rstcorrel_ing("correlativo")
     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
  End If
  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close

End Sub


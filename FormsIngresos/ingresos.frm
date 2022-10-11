VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmIngresosabm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     Registro de Ingresos..."
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   12015
   ClipControls    =   0   'False
   Icon            =   "ingresos.frx":0000
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ingresos.frx":0A02
   ScaleHeight     =   12540
   ScaleWidth      =   17190
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   0
      Picture         =   "ingresos.frx":3C4C
      ScaleHeight     =   960
      ScaleWidth      =   14880
      TabIndex        =   47
      Top             =   0
      Width           =   14940
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "ingresos.frx":6FC7E
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "ingresos.frx":6FE88
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "ingresos.frx":702CA
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "ingresos.frx":704D4
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "ingresos.frx":70A8C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "ingresos.frx":71049
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "ingresos.frx":71253
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "ingresos.frx":71F1D
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "ingresos.frx":724FD
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   360
         Left            =   9915
         TabIndex        =   46
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.Frame FraIngresosDat 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8130
      Left            =   5205
      TabIndex        =   19
      Top             =   1080
      Width           =   9690
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "ingresos.frx":72B21
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   83
         Top             =   7680
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtCorg_codigo 
         Bindings        =   "ingresos.frx":72B3B
         DataField       =   "org_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         ToolTipText     =   "Elije el Código de Organismo de Financiamiento"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Org_codigo"
         BoundColumn     =   "org_codigo"
         Text            =   ""
      End
      Begin VB.TextBox TxtMonto_bolivianos 
         Alignment       =   2  'Center
         DataField       =   "monto_bolivianos"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7800
         TabIndex        =   16
         ToolTipText     =   "Formato con Punto Decimal"
         Top             =   4200
         Width           =   1515
      End
      Begin MSDataListLib.DataCombo DtCrbr_codigo 
         Bindings        =   "ingresos.frx":72B6C
         DataField       =   "rubro_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5445
         TabIndex        =   11
         ToolTipText     =   "Elije el Código del Rubro"
         Top             =   3075
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rubro_codigo"
         BoundColumn     =   "rubro_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCrbr_descripcion 
         Bindings        =   "ingresos.frx":72B8F
         DataField       =   "rubro_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         ToolTipText     =   "Elije la Descripción del Rubro"
         Top             =   3195
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "rubro_descripcion"
         BoundColumn     =   "rubro_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_moneda 
         Bindings        =   "ingresos.frx":72BB2
         DataField       =   "tipo_moneda"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipo_moneda_descripcion"
         BoundColumn     =   "Tipo_moneda"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_tipo_solicitud 
         Bindings        =   "ingresos.frx":72BCF
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Elije el Tipo de Solicitud"
         Top             =   1640
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "solicitud_tipo_descripcion"
         BoundColumn     =   "solicitud_tipo"
         Text            =   ""
      End
      Begin VB.TextBox Txtmonto_dolares 
         Alignment       =   2  'Center
         DataField       =   "monto_dolares"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7800
         TabIndex        =   18
         ToolTipText     =   "Formato con Punto Decimal"
         Top             =   5040
         Width           =   1515
      End
      Begin MSDataListLib.DataCombo DtCFte_codigo 
         Bindings        =   "ingresos.frx":72BEF
         DataField       =   "fte_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         ToolTipText     =   "Elije el Código de Fuente de Financiamiento"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "fte_codigo"
         BoundColumn     =   "fte_codigo"
         Text            =   ""
      End
      Begin VB.TextBox TxtConcepto 
         DataField       =   "ingreso_concepto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   500
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "Acepta hasta 100 caracteres"
         Top             =   4960
         Width           =   7050
      End
      Begin VB.TextBox TxtTipo_cambio 
         Alignment       =   2  'Center
         DataField       =   "tipo_cambio"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         Top             =   3360
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker DTPFecha_Ingreso 
         DataField       =   "fecha_ingreso"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7740
         TabIndex        =   2
         Top             =   1650
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   84017153
         CurrentDate     =   36541
      End
      Begin MSDataListLib.DataCombo DtCCta_descripcion_larga 
         Bindings        =   "ingresos.frx":72C0D
         DataField       =   "cta_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Elije el Nombre de la Cuenta Bancaria"
         Top             =   4215
         Visible         =   0   'False
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "cta_descripcion"
         BoundColumn     =   "Cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCCta_codigo 
         Bindings        =   "ingresos.frx":72C31
         DataField       =   "cta_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Elije el Código de la Cuenta Bancaria"
         Top             =   4560
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Cta_codigo"
         BoundColumn     =   "cta_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCOrg_descripcion 
         Bindings        =   "ingresos.frx":72C55
         DataField       =   "org_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         ToolTipText     =   "Elije el Nombre del Organismo de Financiamiento"
         Top             =   2655
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "Org_descripcion"
         BoundColumn     =   "Org_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCFte_descripcion_larga 
         Bindings        =   "ingresos.frx":72C77
         DataField       =   "fte_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "Elije el Nombre de la Fuente de Financiamiento"
         Top             =   2145
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "fte_descripcion"
         BoundColumn     =   "fte_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCDenominacion_tipo 
         Bindings        =   "ingresos.frx":72C96
         DataField       =   "codigo_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4480
         TabIndex        =   4
         ToolTipText     =   "Elije el Tipo de Comprobante"
         Top             =   1640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Denominacion_tipo"
         BoundColumn     =   "codigo_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "ingresos.frx":72CB8
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1785
         TabIndex        =   56
         Top             =   720
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "ingresos.frx":72CD1
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   57
         Top             =   600
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCtipo_solicitud 
         Bindings        =   "ingresos.frx":72CEB
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   61
         ToolTipText     =   "Elije el Tipo de Solicitud"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "solicitud_tipo"
         BoundColumn     =   "solicitud_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DtCtipo_Comp 
         Bindings        =   "ingresos.frx":72D0B
         DataField       =   "codigo_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6120
         TabIndex        =   62
         ToolTipText     =   "Elije el Tipo de Comprobante"
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo_tipo"
         BoundColumn     =   "codigo_tipo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "ingresos.frx":72D2D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   63
         Top             =   3720
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "ingresos.frx":72D46
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   64
         Top             =   3480
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DtCmoneda 
         Bindings        =   "ingresos.frx":72D5F
         DataField       =   "tipo_moneda"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7920
         TabIndex        =   65
         Top             =   1920
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "tipo_moneda"
         BoundColumn     =   "Tipo_moneda"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "ingresos.frx":72D7C
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   69
         Top             =   5760
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "proceso_descripcion"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "ingresos.frx":72D95
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   70
         Top             =   6360
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "ingresos.frx":72DAE
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   71
         Top             =   6960
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "etapa_descripcion"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "ingresos.frx":72DC7
         DataField       =   "proceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4440
         TabIndex        =   72
         Top             =   5640
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "proceso_codigo"
         BoundColumn     =   "proceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "ingresos.frx":72DE0
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   73
         Top             =   6120
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "ingresos.frx":72DF9
         DataField       =   "etapa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3840
         TabIndex        =   74
         Top             =   6720
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "etapa_codigo"
         BoundColumn     =   "etapa_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "ingresos.frx":72E12
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   77
         Top             =   6000
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "clasif_descripcion"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "ingresos.frx":72E2B
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   78
         Top             =   6720
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "doc_codigo"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "ingresos.frx":72E44
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   82
         Top             =   7680
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "ingresos.frx":72E5E
         DataField       =   "doc_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   84
         Top             =   7200
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ListField       =   "doc_descripcion"
         BoundColumn     =   "doc_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "ingresos.frx":72E77
         DataField       =   "clasif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8400
         TabIndex        =   85
         Top             =   5640
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "clasif_codigo"
         BoundColumn     =   "clasif_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad del POA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   81
         Top             =   7680
         Width           =   1635
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "doc_numero"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7920
         TabIndex        =   80
         Top             =   6720
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   13
         Left            =   8040
         TabIndex        =   79
         Top             =   6460
         Width           =   1455
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Clasificación de Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5520
         TabIndex        =   76
         Top             =   5740
         Width           =   2610
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Código de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5520
         TabIndex        =   75
         Top             =   6460
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   68
         Top             =   5760
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sub Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   67
         Top             =   6375
         Width           =   1170
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Etapa Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   5
         Left            =   180
         TabIndex        =   66
         Top             =   7005
         Width           =   1350
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ingreso_codigo_anterior"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6360
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Anterior del Ingreso (Origen):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   3220
         TabIndex        =   59
         Top             =   240
         Width           =   3150
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFF80&
         X1              =   7440
         X2              =   7440
         Y1              =   1200
         Y2              =   5640
      End
      Begin VB.Label TxtCodigo_solicitud 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "solicitud_codigo"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8115
         TabIndex        =   58
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblmontoRecaudado 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Recaudado Bs: "
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
         Left            =   7680
         TabIndex        =   42
         Top             =   5280
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rubro del Ingreso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   37
         Top             =   3210
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF80&
         X1              =   0
         X2              =   9660
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF80&
         X1              =   15
         X2              =   9720
         Y1              =   5625
         Y2              =   5625
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Registro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4490
         TabIndex        =   36
         Top             =   1360
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   180
         TabIndex        =   35
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label LblGes_Gestion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8565
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblCorrelativo_ingreso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ingreso_codigo"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Bs."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7875
         TabIndex        =   32
         Top             =   3915
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Dólares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7800
         TabIndex        =   31
         Top             =   4755
         Width           =   1500
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fte Financiamiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   29
         Top             =   4680
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7800
         TabIndex        =   28
         Top             =   3080
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7800
         TabIndex        =   27
         Top             =   2240
         Width           =   1380
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7800
         TabIndex        =   26
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Solicitud:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   1365
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Solicitud/Negociación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   5700
         TabIndex        =   24
         Top             =   735
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   7800
         TabIndex        =   23
         Top             =   255
         Width           =   735
      End
      Begin VB.Label Lblcuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   4260
         Width           =   1500
      End
      Begin VB.Label LblCod_Poa 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Financiador:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   21
         Top             =   2685
         Width           =   1065
      End
      Begin VB.Label LblCod_Sol 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Ingreso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   180
         TabIndex        =   20
         Top             =   255
         Width           =   1350
      End
      Begin VB.Label lblBeneficiario 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente/Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   43
         Top             =   3780
         Width           =   1575
      End
   End
   Begin VB.Frame FraIngresosNav 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   8130
      Left            =   60
      TabIndex        =   17
      Top             =   1080
      Width           =   5160
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3000
         TabIndex        =   1
         Top             =   6240
         Width           =   795
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   0
         Top             =   6240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   60
         Top             =   6180
         Width           =   4965
         _ExtentX        =   8758
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
         Caption         =   ""
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "ingresos.frx":72E90
         Height          =   5955
         Left            =   60
         TabIndex        =   44
         Top             =   120
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   10504
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   31
         BeginProperty Column00 
            DataField       =   "ingreso_codigo"
            Caption         =   "Cod.Ingreso"
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
            Caption         =   "Financiador"
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
            Caption         =   "Tipo.Reg."
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
            DataField       =   "ingreso_codigo_anterior"
            Caption         =   "Cod.Origen"
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
            DataField       =   "estado_codigo_dr"
            Caption         =   "Etapa"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "etapa_codigo"
            Caption         =   "Etapa_Proceso"
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.Respaldo"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Cod.Solicitud"
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
            DataField       =   "numero_documento"
            Caption         =   "Cite.Proyecto"
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
            DataField       =   "solicitud_tipo"
            Caption         =   "Tipo.Solicitud"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cdo.Beneficiario"
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
            Caption         =   "Fecha.Registro"
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
            Caption         =   "Tipo.Cambio"
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
            Caption         =   "Tipo.Moneda"
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
            DataField       =   "ingreso_concepto"
            Caption         =   "Concepto.del.Ingreso"
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
            Caption         =   "Tipo.Cmpbte"
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
         BeginProperty Column19 
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
         BeginProperty Column20 
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
         BeginProperty Column21 
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
         BeginProperty Column22 
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
         BeginProperty Column23 
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
         BeginProperty Column24 
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
         BeginProperty Column25 
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
         BeginProperty Column26 
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
         BeginProperty Column27 
            DataField       =   "monto_recaudado_bolivianos"
            Caption         =   "monto_recaudado_bolivianos"
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
            DataField       =   "usr_codigo"
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
         BeginProperty Column29 
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
         BeginProperty Column30 
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
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   195.024
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column21 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column22 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column23 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column24 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column25 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column26 
               Object.Visible         =   0   'False
               ColumnWidth     =   1964.976
            EndProperty
            BeginProperty Column27 
               Object.Visible         =   0   'False
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column28 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column29 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column30 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Registro   (Tipo.Reg.)-->"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   1275
         Left            =   120
         TabIndex        =   86
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "DEV = Devengado     REC = Recaudado     DYR = DEV y REC  ANL = Anulado       DES = Desafectado  DVI = ANL y DES"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   1515
         Left            =   2340
         TabIndex        =   38
         Top             =   6585
         Width           =   2295
      End
   End
   Begin VB.Frame Frmmensaje 
      Height          =   2475
      Left            =   4140
      TabIndex        =   39
      Top             =   6720
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   120
         Width           =   5015
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9240
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
      Caption         =   "Ado_datos1"
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
   Begin MSAdodcLib.Adodc AdoFte_financia 
      Height          =   330
      Left            =   2280
      Top             =   9240
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
      Caption         =   "AdoFte_financia"
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
      Left            =   4440
      Top             =   9240
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
      Caption         =   "AdoOrganismo_finan"
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
      Left            =   6480
      Top             =   9240
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "AdoFc_Rubro_ingresos"
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
      Left            =   8760
      Top             =   9240
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Caption         =   "AdoTipo_solicitud"
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
      Left            =   11040
      Top             =   9240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoTipo_comprobante"
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   120
      Top             =   9720
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
      Caption         =   "Ado_datos4"
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
      Left            =   2280
      Top             =   9720
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
      Caption         =   "AdoTipo_moneda"
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
   Begin MSAdodcLib.Adodc AdoFc_cuenta_bancaria 
      Height          =   330
      Left            =   4440
      Top             =   9720
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
      Caption         =   "AdoFc_cuenta_bancaria"
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   120
      Top             =   10080
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
      Caption         =   "Ado_datos5"
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   2280
      Top             =   10080
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
      Caption         =   "Ado_datos6"
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   4440
      Top             =   10080
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
      Caption         =   "Ado_datos7"
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
   Begin MSAdodcLib.Adodc Ado_datos8 
      Height          =   330
      Left            =   6600
      Top             =   10080
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
      Caption         =   "Ado_datos8"
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   8640
      Top             =   10080
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
      Caption         =   "Ado_datos9"
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
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   10680
      Top             =   10080
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
      Caption         =   "Ado_datos10"
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
      Begin VB.Menu mnuAccion 
         Caption         =   "Devolucion"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmIngresosabm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''========================================================================================
'' Sistema:                  SIG-2014
'' Módulo:                   Ejecución Presupuestaria de Ingresos
'' Base de Datos:            SQL SERVER 2008 R2 (español)
'' Formulario :              FrmIngresosabm.frm
'' Descipción :              Registro de Ingresos Presupuestarios
'' Formularios relacionados: MainMenu.frm (Padre)
''                           ComprobIngreso.rpt (Crystal Reports ver. 8.0)
'' Versión:                  2.0
'' cd now 20140209
''========================================================================================
'
'' ULTIMAS MOD g- 28/05/014
'
Option Explicit

Dim sino As String
Dim v_añadir As Integer

Dim v_añadirstat As Integer
Dim v_cod_solicitud As Integer
Dim rs_datos As New ADODB.Recordset         'Ingresos
Dim rs_datos1 As New ADODB.Recordset        'Unidad Ejecutora
Dim rs_datos4 As New ADODB.Recordset        'Beneficiario Cliente
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset


Dim rstOrganismo_finan As New ADODB.Recordset
Dim rstFte_financia As New ADODB.Recordset
Dim rstFc_convenios As New ADODB.Recordset
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
Dim rstdestino2 As New ADODB.Recordset
Dim buscasi As Integer
Dim operadorbus As String
Dim campobus As String
Dim V_accion As String
Dim fte_codigo1 As String
Dim swcopiar As Integer
Dim swmodificar As Integer

Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim EntrarAdo As Boolean 'Para que al aprobar no muestre uno por uno
Dim queryinicial As String
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir, var_devuelto As String
Dim yacontabilizo As Integer

Private Sub BtnSalir_Click()
'===== Salida del Módulo
  sino = MsgBox("¿Está seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
  If sino = vbYes Then
    Unload Me
  End If
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub DtCDenominacion_tipo_Click(Area As Integer)
    DtCtipo_Comp.BoundText = DtCDenominacion_tipo.BoundText
End Sub

Private Sub DtCDenominacion_tipo_solicitud_Click(Area As Integer)
    DtCtipo_solicitud.BoundText = DtCDenominacion_tipo_solicitud.BoundText
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  V_accion = "COPIA"
'===== Actualización de Despliegue de Datos
   If (Not Ado_datos.Recordset.EOF) And (Not Ado_datos.Recordset.BOF) Then
        If Not IsNull(Ado_datos.Recordset("Correlativo_ingreso")) Then
'            LblCorrelativo_ingreso = IIf(IsNull(Ado_datos.Recordset("Correlativo_ingreso")) = True, " ", Ado_datos.Recordset("Correlativo_ingreso"))
'            LblGes_Gestion = IIf(IsNull(Ado_datos.Recordset("Ges_Gestion")) = True, " ", Ado_datos.Recordset("Ges_Gestion"))
'            TxtCodigo_solicitud = IIf(IsNull(Ado_datos.Recordset("Codigo_solicitud")) = True, " ", Ado_datos.Recordset("Codigo_solicitud"))
'            DTPFecha_Ingreso = IIf(IsNull(Ado_datos.Recordset("Fecha_Ingreso")) = True, " ", Ado_datos.Recordset("Fecha_Ingreso"))
'            TxtTipo_cambio = IIf(IsNull(Ado_datos.Recordset("Tipo_cambio")) = True, 0, Ado_datos.Recordset("Tipo_cambio"))
'            TxtConcepto = IIf(IsNull(Ado_datos.Recordset("Concepto")) = True, " ", Ado_datos.Recordset("Concepto"))
'            Txtmonto_dolares = IIf(IsNull(Ado_datos.Recordset("monto_dolares")) = True, 0, Ado_datos.Recordset("monto_dolares"))
'            TxtMonto_bolivianos = IIf(IsNull(Ado_datos.Recordset("Monto_bolivianos")) = True, 0, Ado_datos.Recordset("Monto_bolivianos"))
'            TxtNumero_documento.Text = IIf(IsNull(Ado_datos.Recordset("numero_documento")) = True, 0, Ado_datos.Recordset("numero_documento"))

'            DtCrbr_codigo.Text = IIf(IsNull(Ado_datos.Recordset("rbr_codigo")) = True, " ", Ado_datos.Recordset("rbr_codigo"))
'            DtCrbr_descripcion.Text = DtCrbr_codigo.BoundText

'            DtCDenominacion_moneda.BoundText = IIf(IsNull(Ado_datos.Recordset("tipo_moneda")) = True, "", Ado_datos.Recordset("tipo_moneda"))

            Select Case Ado_datos.Recordset("Codigo_tipo")
              Case "DYR"
                lblBeneficiario.Visible = False
                Lblcuenta.Visible = True
                DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
                DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
                DtCCta_codigo.Visible = True
                DtCCta_descripcion_larga.Visible = True
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
                
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Visible = False
''                Call activar_Obj 29/06/01
              Case "REC"
                lblBeneficiario.Visible = False
                Lblcuenta.Visible = True
                DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
                DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
                DtCCta_codigo.Visible = True
                DtCCta_descripcion_larga.Visible = True
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'REC' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Visible = False
''                Call desactivar_Obj 29/06/01
''                DtCDenominacion_tipo.Enabled = False
''                CmdCopiar.Enabled = False
              Case "DEV"
                lblBeneficiario.Visible = True
                Lblcuenta.Visible = False
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
                DtCCta_codigo.Visible = False
                DtCCta_descripcion_larga.Visible = False
                
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Caption = " Monto Recaudado: " & CStr(Ado_datos.Recordset("monto_recaudado_dolares"))
'                LblmontoRecaudado.Visible = True
''                Call activar_Obj 29/06/01
''                DtCDenominacion_tipo.Enabled = True
''                CmdCopiar.Enabled = True
              Case "DES"
                lblBeneficiario.Visible = False
                Lblcuenta.Visible = True
                DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
                DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
                DtCCta_codigo.Visible = True
                DtCCta_descripcion_larga.Visible = True
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'DES' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Visible = False
''                Call desactivar_Obj 29/06/01
'                Me.CmdCopiar.Enabled = False
''                DtCDenominacion_tipo.Enabled = False
              Case "ANI"
                lblBeneficiario.Visible = True
                Lblcuenta.Visible = False
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
                DtCCta_codigo.Visible = False
                DtCCta_descripcion_larga.Visible = False
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'ANI' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Visible = False
''                Call desactivar_Obj 29/06/01
'                Me.CmdCopiar.Enabled = False
''                DtCDenominacion_tipo.Enabled = False
              Case "DVI"
                lblBeneficiario.Visible = False
                Lblcuenta.Visible = True
                DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
                DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
                DtCCta_codigo.Visible = True
                DtCCta_descripcion_larga.Visible = True
                dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
                dtc_desc4.Text = dtc_codigo4.BoundText
                dtc_codigo4.Visible = True
                dtc_desc4.Visible = True
'                Set rstTipo_comprobante = New ADODB.Recordset
'                If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'                rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' and codigo_tipo = 'DVI' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'                Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'                AdoTipo_comprobante.Refresh
'                If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
'                DtCDenominacion_tipo.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo")) = True, " ", Ado_datos.Recordset("Codigo_tipo"))
'                LblmontoRecaudado.Visible = False
''                Call desactivar_Obj 29/06/01
'                Me.CmdCopiar.Enabled = False
''                DtCDenominacion_tipo.Enabled = False

            End Select


            '0000
'            DtCDenominacion_tipo_solicitud.BoundText = IIf(IsNull(Ado_datos.Recordset("Codigo_tipo_solicitud")) = True, " ", Ado_datos.Recordset("Codigo_tipo_solicitud"))
'
'            DtCCodigo_documento.Text = IIf(IsNull(Ado_datos.Recordset("Codigo_documento")) = True, " ", Ado_datos.Recordset("Codigo_documento"))
'            DtCDenominacion_documento.Text = DtCCodigo_documento.BoundText
'
'            DtCFte_codigo.Text = IIf(IsNull(Ado_datos.Recordset("fte_codigo")) = True, " ", Ado_datos.Recordset("fte_codigo"))
'            DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
'
'            DtCorg_codigo.Text = IIf(IsNull(Ado_datos.Recordset("org_codigo")) = True, " ", Ado_datos.Recordset("org_codigo"))
'            DtCOrg_descripcion.Text = DtCorg_codigo.BoundText
'
'            DtCcodigo_convenio.Text = IIf(IsNull(Ado_datos.Recordset("codigo_convenio")) = True, "", Ado_datos.Recordset("codigo_convenio"))
'            DtCDenominacion_Convenio.Text = DtCcodigo_convenio.BoundText

'            If Ado_datos.Recordset("Codigo_tipo") = "DEV" Then
'              lblBeneficiario.Visible = True
'              Lblcuenta.Visible = False
'              dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
'              dtc_desc4.Text = dtc_codigo4.BoundText
'              dtc_codigo4.Visible = True
'              dtc_desc4.Visible = True
'              DtCCta_codigo.Visible = False
'              DtCCta_descripcion_larga.Visible = False
'            Else
'                'If Ado_datos.Recordset("Codigo_tipo") = "REC" Or Ado_datos.Recordset("Codigo_tipo") = "DYR" Or Ado_datos.Recordset("Codigo_tipo") = "ANI" Or Ado_datos.Recordset("Codigo_tipo") = "DVI" Then
'              lblBeneficiario.Visible = False
'              Lblcuenta.Visible = True
'              DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
'              DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
'              DtCCta_codigo.Visible = True
'              DtCCta_descripcion_larga.Visible = True
'              dtc_codigo4.Text = IIf(IsNull(Ado_datos.Recordset("beneficiario_Codigo")) = True, " ", Ado_datos.Recordset("beneficiario_Codigo"))
'              dtc_desc4.Text = dtc_codigo4.BoundText
'              dtc_codigo4.Visible = True
'              dtc_desc4.Visible = True
'            End If
            Select Case Ado_datos.Recordset("estado_codigo")
              Case "REG"
                BtnModificar.Enabled = True
                BtnEliminar.Enabled = True
                BtnModificar.Enabled = True
                BtnEliminar.Enabled = True
              Case "ERR"
              Case "APR"
            End Select
            
            'AQUI VERIFICAR QUIEN TIENE ACCESO A APROBAR
            If ((Ado_datos.Recordset("estado_codigo") = "REG") Or (IsNull(Ado_datos.Recordset("estado_codigo")))) Then
                CmdAprueba.Enabled = True
            Else
                CmdAprueba.Enabled = False
            End If
            If (Ado_datos.Recordset!estado_aprobacion = "E") Or (Ado_datos.Recordset!Codigo_tipo = "DES" Or Ado_datos.Recordset!Codigo_tipo = "ANI") Then
              CmdCopiar.Enabled = False
            Else
              CmdCopiar.Enabled = True
            End If
            If (Ado_datos.Recordset("estado_aprobacion") = "S") Or (Ado_datos.Recordset("estado_aprobacion") = "E") Then
              BtnEliminar.Enabled = False
              BtnModificar.Enabled = False
            Else
              BtnEliminar.Enabled = True
              BtnModificar.Enabled = True
            End If

          mnuAccion(0).Enabled = False
          mnuAccion(1).Enabled = False
          mnuAccion(2).Enabled = False
          mnuAccion(3).Enabled = False
          With Ado_datos
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
                'mnuAccion(0).Enabled = False 03/07/01
                'mnuAccion(1).Enabled = True 03/07/01
                mnuAccion(2).Enabled = True
'                sw = 1
            End If
''            If (.Recordset!estado_devengado = "S") And (Trim(.Recordset!estado_recaudado) = "" Or IsNull(.Recordset!estado_recaudado)) Then
''                mnuAccion(0).Enabled = False
''                mnuAccion(2).Enabled = True
''' que se hace cuando se anula un recaudado, se anulan toDOS LOS REGISTROS RECUADADOS?
'''                sw = 2
''            End If
            If (.Recordset!estado_devengado = "S") And (.Recordset!estado_recaudado = "S") Then
                'mnuAccion(0).Enabled = False
                mnuAccion(3).Enabled = True
'                sw = 2
            End If
          End With
' FIN AHORA ***************************

        Else
          mnuAccion(0).Enabled = False
          mnuAccion(1).Enabled = False
          mnuAccion(2).Enabled = False
          mnuAccion(3).Enabled = False

          LblCorrelativo_ingreso = ""
          LblGes_Gestion = ""
          TxtCodigo_solicitud = ""
          DTPFecha_Ingreso = Format(Date, "dd/mm/yyy")
          TxtTipo_cambio = 0
          TxtConcepto = ""
          Txtmonto_dolares = 0
          TxtMonto_bolivianos = 0
          DtCFte_codigo.Text = ""
          DtCFte_descripcion_larga.Text = ""
          DtCorg_codigo.Text = ""
          DtCOrg_descripcion.Text = ""
          DtCCta_codigo.Text = ""
          DtCCta_descripcion_larga.Text = ""
      End If
   End If
End Sub
'
'Private Sub CmdActualTeso_Click()
'  Dim rstacum As New ADODB.Recordset
'  Dim rstdestino As New ADODB.Recordset
'  Set rstacum = New ADODB.Recordset
'  Set rstdestino = New ADODB.Recordset
'  If rstdestino.State = 1 Then rstdestino.Close
'  rstdestino.Open "select * from fc_cuenta_bancaria", db, adOpenKeyset, adLockOptimistic
'  While Not rstdestino.EOF
'    If rstacum.State = 1 Then rstacum.Close
'    rstacum.Open "select sum (monto_bolivianos) as cta_acum from fo_ingresos_cabecera where cta_codigo = '" & rstdestino("cta_codigo") & "'and estado_aprobacion = 'S' and estado_recaudado = 'S' and( estado_anulado <> 'S' and estado_desafectado <> 'S')", db, adOpenStatic, adLockReadOnly
'    If IsNull(rstacum("cta_acum")) Then
'    Else
'      Print rstacum("cta_acum")
'      rstdestino("cta_ingresos") = rstacum("cta_acum")
'      rstdestino.Update
'    End If
'    If rstacum.State = 1 Then rstacum.Close
'    rstdestino.MoveNext
'  Wend
'End Sub
'
'Private Sub CmdAñadir_Click()
''===== Proceso para Añadir y/o modificar Datos
'    var_devuelto = "N"
'    v_añadir = 1
'    FraNavega.Enabled = False
'    fraDatos.Enabled = True
'    fraOpciones.Visible = False
'    FraOpciones2.Visible = True
'    DTPFecha_Ingreso.Value = Date
'    LblGes_Gestion = Year(DTPFecha_Ingreso.Value)
'    If swcopiar = 1 Then
'      LblAccion = "Copiando registros..."
'      DtCorg_codigo.Enabled = False
'      If V_accion = "ANI" Or V_accion = "DES" Or V_accion = "DVI" Then
'        'Me.txtTipo_Cambio = GlTipoCambioOficial g- 28/06/001
'        Print rs_datos!monto_dolares
'        Print rs_datos!tipo_cambio
'        'Txtmonto_dolares = Round(((TxtMonto_bolivianos / GlTipoCambioOficial) * -1), 2) g- 28/06/01
'        Txtmonto_dolares = Txtmonto_dolares * -1
'        'TxtMonto_bolivianos = Round((TxtMonto_bolivianos * -1), 2) g- 28/06/01
'        TxtMonto_bolivianos = TxtMonto_bolivianos * -1
''        DtCDenominacion_tipo.BoundText = "ANI"
'      End If
'    Else
'      LblAccion = "Añadiendo registros..."
'    End If
'    If v_añadir = 1 Then
'        If Not (Ado_datos.Recordset.BOF) Or Not (Ado_datos.Recordset.EOF) Then
'          If swcopiar = 0 Then 'ultimo
'            LblCorrelativo_ingreso = ""
'            LblGes_Gestion = ""
'            TxtCodigo_solicitud = ""
'            DTPFecha_Ingreso = Format(Date, "dd/mm/yyyy")
'            'TxtTipo_cambio = 0
'            TxtConcepto = ""
'            Txtmonto_dolares = 0
'            TxtMonto_bolivianos = 0
'            DtCDenominacion_tipo_solicitud = ""
'            TxtCodigo_solicitud.Text = ""
'            DtCDenominacion_tipo.Text = ""
'            DtCCodigo_documento = ""
'            DtCDenominacion_documento = ""
'            TxtNumero_documento = ""
'            DtCrbr_codigo = ""
'            DtCrbr_descripcion = ""
'            DtCFte_codigo.Text = ""
'            DtCFte_descripcion_larga.Text = ""
'            DtCorg_codigo.Text = ""
'            DtCOrg_descripcion.Text = ""
'            DtCcodigo_convenio.Text = ""
'            DtCDenominacion_Convenio.Text = ""
'            DtCCta_codigo.Text = ""
'            DtCCta_descripcion_larga.Text = ""
'            LblGes_Gestion.Caption = Year(Date)
'            DTPFecha_Ingreso.Value = Format(Date, "dd/mm/yyyy")
'            TxtTipo_cambio = GlTipoCambioOficial
'            Set rstTipo_comprobante = New ADODB.Recordset
'            If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'            rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'            Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'            AdoTipo_comprobante.Refresh
'
'            'Public GlTipoCambioMercado As Currency
'
'          End If 'ultimo
'            Call activar_Obj
'            DtCFte_codigo.Enabled = True
''            DtCOrg_codigo.Enabled = True
'          Select Case V_accion
'            Case "REC"
'              var_devuelto = "REC"
'              Call desactivar_Obj
'              DtCCta_codigo.Enabled = True
'              DtCCta_descripcion_larga.Enabled = True
'              DtCDenominacion_moneda.Enabled = True
'              Txtmonto_dolares.Enabled = True
'              TxtMonto_bolivianos.Enabled = True
'              Set rstTipo_comprobante = New ADODB.Recordset
'              If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'              rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'              Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'              AdoTipo_comprobante.Refresh
'              Me.DtCDenominacion_tipo.BoundText = "REC"
'
'              Me.dtc_codigo4.Visible = False
'              Me.dtc_desc4.Visible = False
'              Me.lblBeneficiario.Visible = False
'              Me.Lblcuenta.Visible = True
'              DtCCta_codigo.Visible = True
'              DtCCta_descripcion_larga.Visible = True
'
'              DtCCta_codigo.Enabled = True
'              DtCCta_descripcion_larga.Enabled = True
'              DtCDenominacion_moneda.Enabled = True
'              Txtmonto_dolares.Enabled = True
'              TxtMonto_bolivianos.Enabled = True
'              TxtConcepto.Enabled = True
'              DtCrbr_codigo.Enabled = True
'              DtCrbr_descripcion.Enabled = True
'              TxtTipo_cambio.Enabled = True
'
'              DtCDenominacion_tipo_solicitud.Enabled = True
'              DtCCodigo_documento.Enabled = True
'              DtCDenominacion_documento.Enabled = True
'              DtCCta_codigo.Enabled = True
'              DtCCta_descripcion_larga.Enabled = True
'
'              TxtNumero_documento.Enabled = True
'              TxtCodigo_solicitud.Enabled = True
'
'            Case "ANI"
'              var_devuelto = "ANI"
'              Call desactivar_Obj
''              DtCCta_codigo.Enabled = True
''              DtCCta_descripcion_larga.Enabled = True
''              DtCDenominacion_moneda.Enabled = True
''LAST
''              Txtmonto_dolares.Text = (Txtmonto_dolares.Text * -1)
''              TxtMonto_bolivianos.Enabled = (TxtMonto_bolivianos * -1)
'              Txtmonto_dolares.Enabled = False
'              TxtMonto_bolivianos.Enabled = False
'              Me.TxtTipo_cambio.Enabled = False
'              Me.DtCCta_codigo.Enabled = False
'              Me.DtCCta_descripcion_larga.Enabled = False
'              Me.DtCrbr_codigo.Enabled = False
'              Me.DtCrbr_descripcion.Enabled = False
'              Me.DtCDenominacion_moneda.Enabled = False
'              Set rstTipo_comprobante = New ADODB.Recordset
'              If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'              rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'              Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'              AdoTipo_comprobante.Refresh
'              Me.DtCDenominacion_tipo.BoundText = "ANI"
'
'            Case "DES"
'              var_devuelto = "DES"
'              Call desactivar_Obj
'              Txtmonto_dolares.Enabled = False
'              TxtMonto_bolivianos.Enabled = False
'              Me.TxtTipo_cambio.Enabled = False
'              Me.DtCCta_codigo.Enabled = False
'              Me.DtCCta_descripcion_larga.Enabled = False
'          Case "DVI"
'              var_devuelto = "DVI"
'              Call desactivar_Obj
'              Txtmonto_dolares.Enabled = True
'              TxtMonto_bolivianos.Enabled = True
'              Me.TxtTipo_cambio.Enabled = True
'              Me.DtCCta_codigo.Enabled = False
'              Me.DtCCta_descripcion_larga.Enabled = False
'              Set rstTipo_comprobante = New ADODB.Recordset
'              If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'              rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'R' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'              Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'              AdoTipo_comprobante.Refresh
'              Me.DtCDenominacion_tipo.BoundText = "DVI"
'          End Select
'
'      ' DtCDenominacion_tipo.BoundText
'
'        End If
'    End If
'End Sub
'
''Private Sub CmdAñadir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''  CmdAñadir.BackColor = &HC0FFFF
''End Sub
'
'Private Sub cmdAprueba_Click()
''===== Proceso para generar Asientos Contables Automáticos "CAD" y "CAR"
'
'  sino = MsgBox("¿Está seguro de aprobar el Registro?", vbYesNo + vbQuestion, "CONFIRMAR...")
'  If sino = vbYes Then
'    If Ado_datos.Recordset("codigo_tipo") = "REC" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        If rstdestino("monto_dolares") < rstdestino("monto_recaudado_dolares") + Ado_datos.Recordset("monto_dolares") Then
'          MsgBox "El monto que está intentando recaudar en dolares es mayor al DEVENGADO, por fsavor corrija Monto: Devengado: " & CStr(rstdestino("monto_dolares")) & " Solo puede recaudar :" & CStr(rstdestino("monto_dolares") - rstdestino("monto_recaudado_dolares")), vbOKOnly + vbCritical, "ERROR en el minto de Recaudo"
'          Exit Sub
'        End If
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'
''**** aqui consultar tia que hacer ***************
'    If Ado_datos.Recordset("codigo_tipo") = "DES" Then
'
'    End If
'
'    If Ado_datos.Recordset("codigo_tipo") = "ANI" Then
'
'    End If
''**** aqui consultar tia que hacer ***************
'
'
'    Dim cta_deb1 As String
'    Dim Subcta_deb11 As String
'    Dim Subcta_deb21 As String
'
'    Dim cta_credito1 As String
'    Dim Subcta_cred11 As String
'    Dim Subcta_cred21 As String
'
'    Dim cod_ant As Integer
'    Dim org_ant As String
'
'    If DtCCta_codigo.Text <> "01" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstFc_cuenta_bancaria.Find " cta_codigo = '" & DtCCta_codigo & "'", , adSearchForward, 1
'      If Not rstFc_cuenta_bancaria.EOF Then
'        fte_codigo1 = rstFc_cuenta_bancaria("fte_codigo")
'      Else
'      End If
'    Else
'        fte_codigo1 = Me.DtCFte_codigo.Text
'    End If
'
'    If Ado_datos.Recordset!Codigo_tipo = "DEV" Or Ado_datos.Recordset!Codigo_tipo = "DES" Then
'      fte_codigo1 = Me.DtCFte_codigo.Text
'    End If
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim v_Tipo_Comp(1, 2)
'    If Ado_datos.Recordset("codigo_tipo") = "DYR" Then
'      j = 2
'      v_Tipo_Comp(1, 1) = "CAD"
'      v_Tipo_Comp(1, 2) = "CAR"
'    Else
'      j = 1
'      v_Tipo_Comp(1, 1) = IIf(Ado_datos.Recordset("codigo_tipo") = "DEV", "CAD", IIf(Ado_datos.Recordset("codigo_tipo") = "REC", "CAR", IIf(Ado_datos.Recordset("codigo_tipo") = "DES", "DES", IIf(Ado_datos.Recordset("codigo_tipo") = "ANI", "ANI", ""))))
'    End If
'
'    If Ado_datos.Recordset("codigo_tipo") = "DVI" Then
'      j = 1
'      v_Tipo_Comp(1, 1) = "DVI"
'    End If
'
'    For i = 1 To j
'      If rstdestino.State = 1 Then rstdestino.Close
'      If v_Tipo_Comp(1, i) = "CAD" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "CAR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Or v_Tipo_Comp(1, i) = "REC" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'
'      If v_Tipo_Comp(1, i) = "DEV" Or v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'
'      If v_Tipo_Comp(1, i) = "" Then
'        MsgBox "Antes de aprobar defina que tipo " & vbCrLf & "de registro está procesando", vbOKOnly + vbCritical, "Error de aprobación... "
'        Exit Sub
'      End If
'
'      If v_Tipo_Comp(1, i) = "DVI" Then
'' g- 02/07/01 VERIFICAR SI SE ESTA ABRIENDO BIEN
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount < 1 Or rstdestino2.RecordCount < 1 Then
'          MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'          Exit Sub
'        End If
'
'      End If
'
'      If rstdestino.RecordCount < 1 Then
'        MsgBox "Este comprobante no puede ser aprobado, Porque el RUBRO no EXISTE en el RELACIONADOR, por favor contáctese con el administrador", vbOKOnly + vbCritical, "Error al aprobar..."
'        Exit Sub
'      End If
'
'    Next
'
'    If rstdestino.State = 1 Then rstdestino.Close
'    db.BeginTrans
'    Frmmensaje.Visible = True
'    LblMensaje.Caption = "Este proceso tomará solo unos segundos, gracias"
'    Dim d_cta_nombre_1 As String
'    Dim d_aux1_1 As String
'    Dim d_aux2_1 As String
'    Dim d_aux3_1 As String
'    Dim h_cta_nombre_1 As String
'    Dim h_aux1_1 As String
'    Dim h_aux2_1 As String
'    Dim h_aux3_1 As String
'    If rstdestino.State = 1 Then rstdestino.Close
'    '===== ini registro de co_comprobante_M =====
'
'    For i = 1 To j
'' nuevo ini
'      If v_Tipo_Comp(1, i) = "CAD" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "CAR" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "ANI" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'      End If
'      If v_Tipo_Comp(1, i) = "DES" Then
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & "", db, adOpenKeyset, adLockReadOnly
'      End If
'
'      If v_Tipo_Comp(1, i) = "DVI" Then
'' g- 02/07/01 VERIFICAR SI SE ESTA ABRIENDO BIEN
'        If rstdestino.State = 1 Then rstdestino.Close
'        rstdestino.Open "select * from fc_relacionador_ingresos where inst = 'DEV' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo), db, adOpenKeyset, adLockReadOnly
'        If rstdestino2.State = 1 Then rstdestino2.Close
'        rstdestino2.Open "select * from fc_relacionador_ingresos where inst = 'EFE' and rec_rub_i <= " & (DtCrbr_codigo) & " and rec_rub_f >= " & (DtCrbr_codigo) & " and subcta_deb2 = '" & IIf(fte_codigo1 = "10", "01", IIf(fte_codigo1 = "43", "02", IIf(fte_codigo1 = "80", "03", ""))) & "' ", db, adOpenKeyset, adLockReadOnly
'        If rstdestino.RecordCount > 0 And rstdestino2.RecordCount > 0 Then
'          cta_deb1 = rstdestino!cta_credito
'          Subcta_deb11 = rstdestino!Subcta_cred1
'          Subcta_deb21 = rstdestino!Subcta_cred2
'
'          cta_credito1 = rstdestino2!cta_deb
'          Subcta_cred11 = rstdestino2!Subcta_deb1
'          Subcta_cred21 = rstdestino2!Subcta_deb2
'        Else
'          MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''          Exit Sub
'        End If
'      End If
'
'      If rstdestino.RecordCount > 0 And v_Tipo_Comp(1, i) <> "DVI" Then
'        cta_deb1 = rstdestino("cta_deb")
'        Subcta_deb11 = rstdestino("Subcta_deb1")
'        Subcta_deb21 = rstdestino("Subcta_deb2")
'        cta_credito1 = rstdestino("cta_credito")
'        Subcta_cred11 = rstdestino("Subcta_cred1")
'        Subcta_cred21 = rstdestino("Subcta_cred2")
''      Else
''        MsgBox "Rubro no presupuestado", vbCritical + vbOKOnly, "ERROR... "
''        Exit Sub
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_deb1 & "' and SubCta1 = '" & Subcta_deb11 & "' and SubCta2 = '" & Subcta_deb21 & "' ", db, adOpenKeyset, adLockReadOnly
'      If rstdestino.RecordCount > 0 Then
'        d_cta_nombre_1 = rstdestino("NombreCta")
'        d_aux1_1 = rstdestino("aux1")
'        d_aux2_1 = rstdestino("aux2")
'        d_aux3_1 = rstdestino("aux3")
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from cc_Plan_cuentas where Cuenta = '" & cta_credito1 & "' and SubCta1 = '" & Subcta_cred11 & "' and SubCta2 = '" & Subcta_cred21 & "' ", db, adOpenKeyset, adLockReadOnly
'      If rstdestino.RecordCount > 0 Then
'        h_cta_nombre_1 = rstdestino("NombreCta")
'        h_aux1_1 = rstdestino("aux1")
'        h_aux2_1 = rstdestino("aux2")
'        h_aux3_1 = rstdestino("aux3")
'      End If
'' nuevo fin
'
''========================================
'    '==== verifica ya fue contabilizado
'      yacontabilizo = 0
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_comprobante_m where Cod_trans = '" & Ado_datos.Recordset!correlativo_anterior & "' and org_codigo = '" & Ado_datos.Recordset!org_codigo & "' and tipo_comp = '" & v_Tipo_Comp(1, i) & "' AND STATUS = 'S'", db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'        yacontabilizo = 1
'      Else
'        yacontabilizo = 0
'      End If
'      If yacontabilizo = 1 Then
'        'MsgBox "aqui recontabilizar" & rstdestino!Cod_trans & " -- " & rstdestino!org_codigo & " / " & rstdestino!Cod_Comp
'        Cont_Comp = rstdestino!Cod_Comp
'      Else
'        '===== ini GENERA EL CODIGO DE COMPROBANTE ====
'        Set rstCodComp = New ADODB.Recordset
'        rstCodComp.CursorLocation = adUseClient
'        If rstCodComp.State = 1 Then rstCodComp.Close
'        rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
'        If rstCodComp.RecordCount > 0 Then
'          Cont_Comp = CDbl(rstCodComp!numero_correlativo)
'          Cont_Comp = Cont_Comp + 1
'          rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
'          rstCodComp.Update
'        End If
'        If rstCodComp.State = 1 Then rstCodComp.Close
'        '===== fin TERMINA GENERACION DE COMPROBANTE =====
'
'        '==== ini registro co_comprobantre_m
'
'        rstdestino.AddNew
'        rstdestino("cod_comp") = Cont_Comp
'      End If
'
''========================================
'
''      '===== ini GENERA EL CODIGO DE COMPROBANTE ====
''      Set rstCodComp = New ADODB.Recordset
''      rstCodComp.CursorLocation = adUseClient
''      If rstCodComp.State = 1 Then rstCodComp.Close
''      rstCodComp.Open "select * from fc_Correl  where tipo_tramite = 'cmbte'", db, adOpenDynamic, adLockOptimistic
''      If rstCodComp.RecordCount > 0 Then
''        Cont_Comp = Val(rstCodComp!numero_correlativo)
''        Cont_Comp = Cont_Comp + 1
''        rstCodComp!numero_correlativo = Trim(Str(Cont_Comp))
''        rstCodComp.Update
''      End If
''      If rstCodComp.State = 1 Then rstCodComp.Close
''      '===== fin TERMINA GENERACION DE COMPROBANTE =====
'
'
'      '==== ini registro co_comprobantre_m
''anterior
''      If rstdestino.State = 1 Then rstdestino.Close
''      rstdestino.Open "select * from co_comprobante_m where Cod_Comp = 0", db, adOpenKeyset, adLockOptimistic
''      If rstdestino.RecordCount > 0 Then
''      End If
''      rstdestino.AddNew
'
''      rstdestino("cod_comp") = Cont_Comp
''anterior
'      rstdestino("cod_trans") = Ado_datos.Recordset!correlativo_anterior  'Ado_datos.Recordset("correlativo_ingreso")
'      rstdestino("org_codigo") = Ado_datos.Recordset("org_codigo")
'      rstdestino("cod_trans_detalle") = 1
'      rstdestino("Num_Respaldo") = Ado_datos.Recordset("numero_documento")
'      'rstdestino("Fecha_A") = Date
'      If yacontabilizo = 0 Then
'        rstdestino("Fecha_A") = Date
'      End If
'      rstdestino("codigo_beneficiario") = "-"
'      rstdestino("glosa") = Ado_datos.Recordset("Concepto")
'      rstdestino("status") = "S"
'      rstdestino("ges_gestion") = Ado_datos.Recordset("ges_gestion")
'      rstdestino!tipo_moneda = Ado_datos.Recordset!tipo_moneda
'      rstdestino("codigo_documento") = Ado_datos.Recordset("codigo_documento")
'      rstdestino("Tipo_Comp") = v_Tipo_Comp(1, i) 'IIf(Ado_datos.Recordset("codigo_tipo") = "DEV", "CAD", IIf(Ado_datos.Recordset("codigo_tipo") = "REC", "CAR", v_Tipo_Comp(i)))
''      rstdestino("Usr_Usuario") = GlUsuario
''      rstdestino("Fecha_registro") = Format(Date, "dd/mm/yyyy")
''      rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
'      If yacontabilizo = 0 Then
'        rstdestino("Usr_Usuario") = GlUsuario
'        rstdestino("Fecha_registro") = Format(Date, "dd/mm/yyyy")
'        rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
'      End If
'      rstdestino.Update
'      '==== fin registro co_comprobantre_m
'
'      '===== ini registra CO_diaRIO =========
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from co_diario where Cod_Comp = " & Cont_Comp, db, adOpenKeyset, adLockOptimistic
'      If rstdestino.RecordCount > 0 Then
'
'      Else
'        rstdestino.AddNew
'        rstdestino("Cod_Comp") = Cont_Comp
'      End If
'
'      rstdestino("Tipo_Comp") = v_Tipo_Comp(1, i)
'      rstdestino("Cod_Comp_C") = Cont_Comp
'      'If v_Tipo_Comp(1, i) = "DEV" Or v_Tipo_Comp(1, i) = "REC" Then
'      If (Ado_datos.Recordset("codigo_tipo") = "DEV") Or (Ado_datos.Recordset("codigo_tipo") = "REC") Or (Ado_datos.Recordset("codigo_tipo") = "DYR") Then
'        rstdestino("D_Cuenta") = cta_deb1
'        rstdestino("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("D_Subcta1") = Subcta_deb11
'        rstdestino("D_SubCta2") = Subcta_deb21
'        rstdestino("D_Aux1") = d_aux1_1
'        rstdestino("D_Aux2") = d_aux2_1
'        rstdestino("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino("D_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino("D_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
''        rstdestino("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino("D_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("D_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("D_Cambio") = Ado_datos.Recordset("tipo_cambio")
''AQUI MONEDA 02/07/01
'        rstdestino("D_Cambio") = Ado_datos.Recordset("tipo_cambio")
'
'        rstdestino("H_Cuenta") = cta_credito1
''        rstdestino("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("H_SubCta1") = Subcta_cred11
'        rstdestino("H_SubCta2") = Subcta_cred21
'        rstdestino("H_Aux1") = h_aux1_1
'        rstdestino("H_Aux2") = h_aux2_1
'        rstdestino("H_Aux3") = h_aux3_1
'        'rstdestino("H_Cta_Larga") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino("h_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino("h_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
''        rstdestino("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino("H_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("H_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("H_Cambio") = Ado_datos.Recordset("tipo_cambio")
'      End If
'
'      'If (v_Tipo_Comp(1, i) = "DES") Or (v_Tipo_Comp(1, i) = "ANI") Then
'      If (Ado_datos.Recordset("codigo_tipo") = "DES") Or (Ado_datos.Recordset("codigo_tipo") = "ANI") Then
'        'desafecta un devengado
'        rstdestino("D_Cuenta") = cta_credito1
'        rstdestino("D_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("D_Subcta1") = Subcta_cred11
'        rstdestino("D_SubCta2") = Subcta_cred21
'        rstdestino("D_Aux1") = h_aux1_1
'        rstdestino("D_Aux2") = h_aux2_1
'        rstdestino("D_Aux3") = h_aux3_1
''        rstdestino("D_Cta_Larga") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino("D_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino("D_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
''        rstdestino("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino("D_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("D_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("D_Cambio") = Ado_datos.Recordset("tipo_cambio")
'
'        rstdestino("H_Cuenta") = cta_deb1
''        rstdestino("H_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("H_SubCta1") = Subcta_deb11
'        rstdestino("H_SubCta2") = Subcta_deb21
'        rstdestino("H_Aux1") = d_aux1_1
'        rstdestino("H_Aux2") = d_aux2_1
'        rstdestino("H_Aux3") = d_aux3_1
''        rstdestino("H_Cta_Larga") = "VESCT"
'        If d_aux1_1 = "01" Then
'          rstdestino("h_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'          'DtCCta_descripcion_larga
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino("h_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
'
''        rstdestino("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino("H_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("H_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("H_Cambio") = Ado_datos.Recordset("tipo_cambio")
'      End If
'
'      '==== INI DVI ====
'      If (Ado_datos.Recordset!Codigo_tipo = "DVI") Then
'        rstdestino("D_Cuenta") = cta_deb1
''        rstdestino("D_Nombre") = d_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("D_Subcta1") = Subcta_deb11
'        rstdestino("D_SubCta2") = Subcta_deb21
'        rstdestino("D_Aux1") = d_aux1_1
'        rstdestino("D_Aux2") = d_aux2_1
'        rstdestino("D_Aux3") = d_aux3_1
'        If d_aux1_1 = "01" Then
'          rstdestino("D_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'        End If
'        If d_aux1_1 = "02" Then
'          rstdestino("D_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
''        rstdestino("D_Des_Larga") = "-" ' CAMPO PARA ELIMINAR
'        rstdestino("D_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("D_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("D_Cambio") = Ado_datos.Recordset("tipo_cambio")
'        rstdestino("H_Cuenta") = cta_credito1
''        rstdestino("H_Nombre") = h_cta_nombre_1 ' CAMPO PARA ELIMINAR
'        rstdestino("H_SubCta1") = Subcta_cred11
'        rstdestino("H_SubCta2") = Subcta_cred21
'        rstdestino("H_Aux1") = h_aux1_1
'        rstdestino("H_Aux2") = h_aux2_1
'        rstdestino("H_Aux3") = h_aux3_1
'        'rstdestino("H_Cta_Larga") = "VESCT"
'        If h_aux1_1 = "01" Then
'          rstdestino("h_Cta_Larga") = IIf(Len(Trim(Ado_datos.Recordset("codigo_beneficiario"))) > 0, Ado_datos.Recordset("codigo_beneficiario"), "-")
'          'DtCCta_descripcion_larga
'        End If
'        If h_aux1_1 = "02" Then
'          rstdestino("h_Cta_Larga") = Ado_datos.Recordset("cta_codigo")
'        End If
''        rstdestino("H_Des_Larga") = "-"   ' CAMPO PARA ELIMINAR
'        rstdestino("H_MontoBs") = IIf(Ado_datos.Recordset("monto_bolivianos") < 0, (Ado_datos.Recordset("monto_bolivianos") * -1), Ado_datos.Recordset("monto_bolivianos"))
'        rstdestino("H_MontoDl") = IIf(Ado_datos.Recordset("monto_dolares") < 0, (Ado_datos.Recordset("monto_dolares") * -1), Ado_datos.Recordset("monto_dolares"))
'        rstdestino("H_Cambio") = Ado_datos.Recordset("tipo_cambio")
'      End If
'      '==== FIN DVI ====
'
''      rstdestino("Usr_Usuario") = GlUsuario
''      rstdestino("Fecha_registro") = Date
''      rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
'      If yacontabilizo = 0 Then
'        rstdestino("Usr_Usuario") = GlUsuario
'        rstdestino("Fecha_registro") = Date
'        rstdestino("Hora_registro") = Format(Time, "hh:mm:ss")
'      End If
'      rstdestino.Update
'      If rstdestino.State = 1 Then rstdestino.Close
'      '======= fin registra co_diario ==========
'    Next i
'    '======= inI Actualiza campos de estatus de ingresos ==========
'    If rstdestino.State = 1 Then rstdestino.Close
'    rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = '" & Ado_datos.Recordset("correlativo_ingreso") & "' and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' ", db, adOpenDynamic, adLockOptimistic
'    rstdestino.MoveFirst
'    If Not (rstdestino.EOF) Then
'      rstdestino("estado_aprobacion") = "S"
'        If Ado_datos.Recordset("codigo_tipo") = "DEV" Then
'          rstdestino("estado_devengado") = "S"
'        End If
'        If Ado_datos.Recordset("codigo_tipo") = "REC" Then
'          rstdestino("estado_recaudado") = "S"
'        End If
'        If Ado_datos.Recordset("codigo_tipo") = "DYR" Then
'          rstdestino("estado_devengado") = "S"
'          rstdestino("estado_recaudado") = "S"
'        End If
'
'        If Ado_datos.Recordset("codigo_tipo") = "DES" Then
'          rstdestino("estado_desafectado") = "S"
'        End If
'        If Ado_datos.Recordset("codigo_tipo") = "ANI" Then
'          rstdestino("estado_anulado") = "S"
'        End If
'        If Ado_datos.Recordset("codigo_tipo") = "DVI" Then
'          rstdestino!estado_desafectado = "S"
'          rstdestino!estado_anulado = "S"
'        End If
'       rstdestino.Update
'       If rstdestino.State = 1 Then rstdestino.Close
'    End If
'    '======= fin Actualiza campos de estatus de ingresos ==========
'
'    cod_ant = 0
'    org_ant = ""
'    '======= ini Actualiza el monto recaudado  ==========
'    If (Ado_datos.Recordset("codigo_tipo") = "REC") Then
'      '      If rstdestino.State = 1 Then rstdestino.Close
'      '      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      '      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'      '        cod_ant = rstdestino("correlativo_anterior")
'      '        org_ant = rstdestino("org_codigo")
'      '      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") + Ado_datos.Recordset("monto_dolares")
'          rstdestino.Update
'      End If
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'
'    If (Ado_datos.Recordset("codigo_tipo") = "DES") Then
''      If rstdestino.State = 1 Then rstdestino.Close
''      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
''      Print Ado_datos.Recordset("correlativo_anterior")
''      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
''        cod_ant = IIf(IsNull(rstdestino("correlativo_anterior")), 0, rstdestino("correlativo_anterior"))
''        org_ant = rstdestino("org_codigo")
''      End If
'
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        If rstdestino("codigo_tipo") = "DEV" Then 'And Ado_datos.Recordset("codigo_tipo") = "DES"
''          rstdestino!estado_desafectado = "S" 02/07/01
'          rstdestino!estado_devengado = "L"
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'        Else
'          rstdestino("estado_desafectado") = "S"
''          rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - Ado_datos.Recordset("monto_dolares")
'          cod_ant = IIf(IsNull(rstdestino("correlativo_anterior")), 0, rstdestino("correlativo_anterior"))
'          org_ant = rstdestino("org_codigo")
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'          rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & cod_ant & " and org_codigo = '" & org_ant & "' ", db, adOpenKeyset, adLockOptimistic
'          If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'            rstdestino("monto_recaudado_dolares") = rstdestino("monto_recaudado_dolares") - Ado_datos.Recordset("monto_dolares")
'          End If
'          rstdestino.Update
'          If rstdestino.State = 1 Then rstdestino.Close
'        End If
'      End If
'    End If
'
'    If (Ado_datos.Recordset("codigo_tipo") = "ANI") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("correlativo_anterior") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        If rstdestino("codigo_tipo") = "REC" Then
''          rstdestino("estado_desafectado") = ""
'          rstdestino("estado_recaudado") = "L"
''          rstdestino("estado_devengado") = "S" 02/07/01
''          rstdestino("estado_anulado") = ""
''          rstdestino("codigo_tipo") = "DEV" 02/07/01
'          rstdestino("monto_recaudado_dolares") = 0
'        End If
'      End If
'      rstdestino.Update
''      Print rstdestino!correlativo_anterior
''      Print rstdestino!monto_recaudado
'      cod_ant = 0
'      org_ant = ""
'      Call f_actual_rec(rstdestino!org_codigo, rstdestino!correlativo_anterior)
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'    If (Ado_datos.Recordset!Codigo_tipo = "DVI") Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset!correlativo_anterior & " and org_codigo = '" & Ado_datos.Recordset!org_codigo & "' ", db, adOpenKeyset, adLockOptimistic
'      If (Not rstdestino.BOF) And (Not rstdestino.EOF) Then
'        rstdestino!estado_recaudado = "V"
'        rstdestino!estado_devengado = "V"
'      End If
'      rstdestino.Update
'      If rstdestino.State = 1 Then rstdestino.Close
'    End If
'    '======= fin Actualiza el monto recaudado  ==========
'
'    '======= ini Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
'    If Ado_datos.Recordset("codigo_tipo") = "REC" Or Ado_datos.Recordset("codigo_tipo") = "DYR" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & Ado_datos.Recordset("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'      If Not rstdestino.EOF Then
'        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + Ado_datos.Recordset("monto_bolivianos")
'        rstdestino.Update
'      End If
'    End If
'    If Ado_datos.Recordset("codigo_tipo") = "ANI" Then
'      If rstdestino.State = 1 Then rstdestino.Close
'      rstdestino.Open "select * from fc_cuenta_bancaria where cta_codigo = '" & Ado_datos.Recordset("cta_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'      If Not rstdestino.EOF Then
'        rstdestino("cta_ingresos") = rstdestino("cta_ingresos") + Ado_datos.Recordset("monto_bolivianos")
'        rstdestino.Update
'      End If
'    End If
'    '======= fin Actualiza el monto bolivianos de fc_cuenta_bancaria ==========
'    LblMensaje.Caption = "El proceso concluyó exitosamente, gracias"
'    Frmmensaje.Visible = False
'    db.CommitTrans
'  End If
'  marca1 = Ado_datos.Recordset.Bookmark
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'  db.Execute "EXEC ts_mf_ActualizaCtaBancaria"
'End Sub
'
'Private Sub BtnEliminar_Click()
'' ===== Proceso para confirmar el eliminado de registros
'  v_añadir = 3
'  sino = MsgBox("¿Está seguro de ANULAR este registro?", vbYesNo + vbQuestion, "Atención...")
'  If sino = vbYes Then
'    Call elimina
'    Call errado
'  End If
'End Sub
'
'Private Sub CmdBuscar_Click()
''JQA
''  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
''  Dim ClBuscaSec As ClBuscaSecuencialEnRS
'  PosibleApliqueFiltro = False
'  Dim rsNada As ADODB.Recordset
'  Dim GrSqlAux As String
'  Set ClBuscaGrid = New ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.EsTdbGrid = False
'  Set ClBuscaGrid.GridTrabajo = dg_datos
'  ClBuscaGrid.QueryUtilizado = queryinicial
'  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'  ClBuscaGrid.CamposVisibles = "110"
'  ClBuscaGrid.Ejecutar
'  PosibleApliqueFiltro = True
'End Sub
'
''Private Sub CmdBuscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'' CmdBuscar.Picture = LoadPicture("d:\Pragma\iconos\busca1.ico")
'' CmdBuscar.BackColor = &HC0FFFF
'' Image1.Visible = True
'
''End Sub
'
''Private Sub Cmdbusfin_Click()
''  FrmBuscar.Visible = False
''  FraOpciones.Enabled = True
''End Sub
'
'Private Sub CmdCancelar_Click()
''===== Ini cancela actualizaciones ==========
'   FraOpciones2.Visible = False
'   fraOpciones.Visible = True
'   FraNavega.Enabled = True
'   fraDatos.Enabled = False
''   Ado_datos.Refresh
''  Set Ado_datos.Recordset = rs_datos
'  rs_datos.Requery
''  Set dg_datos.DataSource = Ado_datos.Recordset
'  LblAccion = ""
'  Call activar_Obj
'End Sub
'
'Private Sub CmdGrabar_Click()
''======= Ini grabado de datos
'   swgraba = 0
'   Call valida
'
'   If swgraba = 1 Then
'      FraOpciones2.Visible = False
'      fraOpciones.Visible = True
'      FraNavega.Enabled = True
'      fraDatos.Enabled = False
'
'      If v_añadir = 1 Then
'        db.BeginTrans
'         Call add_correl
'         Set rstdestino = New ADODB.Recordset
'         rstdestino.Open "select * from fo_ingresos_cabecera order by correlativo_ingreso, org_codigo  ", db, adOpenDynamic, adLockOptimistic
'         rstdestino.AddNew
''         lblges_gestion.Caption = "2002"
'         rstdestino("Correlativo_ingreso") = correlativo1
'         rstdestino("Ges_Gestion") = Trim(LblGes_Gestion.Caption)
'         rstdestino("Codigo_solicitud") = TxtCodigo_solicitud.Text
'         rstdestino("rbr_codigo") = DtCrbr_codigo.Text
'         rstdestino("tipo_moneda") = DtCDenominacion_moneda.BoundText
'         rstdestino("UNI_CODIGO") = dtc_codigo1
'
'         Select Case V_accion
'            Case "REC"
'              rstdestino("Codigo_tipo") = "REC"
'              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
'              rstdestino("estado_recaudado") = "N"
'            Case "DES"
'              rstdestino("Codigo_tipo") = "DES"
'              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
'              rstdestino("estado_desafectado") = "N"
'            Case "ANI"
'              rstdestino("Codigo_tipo") = "ANI"
'              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
'              rstdestino("estado_anulado") = "N"
'            Case "DVI"
'              rstdestino("Codigo_tipo") = "DVI"
'              rstdestino("correlativo_anterior") = CDbl(LblCorrelativo_ingreso)
'              rstdestino("estado_desafectado") = "N"
'              rstdestino("estado_anulado") = "N"
'            Case "COPIA"
'              rstdestino("Codigo_tipo") = DtCDenominacion_tipo.BoundText
'              If DtCDenominacion_tipo.BoundText = "DEV" Then
'               rstdestino("estado_devengado") = "N"
'               rstdestino("correlativo_anterior") = correlativo1
'              End If
'              If DtCDenominacion_tipo.BoundText = "REC" Then
'               rstdestino("estado_recaudado") = "N"
'              End If
'              If DtCDenominacion_tipo.BoundText = "DYR" Then
'               rstdestino("correlativo_anterior") = correlativo1
'               rstdestino("estado_recaudado") = "N"
'               rstdestino("estado_devengado") = "N"
'              End If
'
'         End Select ' DtCDenominacion_tipo.BoundText
'
'         rstdestino("Codigo_tipo_solicitud") = IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
'         rstdestino("Codigo_documento") = DtCCodigo_documento.Text
''         DTPFecha_Ingreso.Value = Date
'         rstdestino("Fecha_Ingreso") = DTPFecha_Ingreso.Value
'         rstdestino("Tipo_Cambio") = TxtTipo_cambio.Text
'         rstdestino("Concepto") = (TxtConcepto.Text)
'         rstdestino("fte_codigo") = DtCFte_codigo.Text
'         rstdestino("org_codigo") = DtCorg_codigo.Text
'         rstdestino("codigo_convenio") = DtCcodigo_convenio
''         rstdestino("cta_codigo") = DtCCta_codigo.Text
'         If DtCDenominacion_tipo.BoundText = "DEV" Then
'           rstdestino("Codigo_beneficiario") = dtc_codigo4.Text
'           rstdestino("cta_codigo") = ""
'         End If
'
'         If DtCDenominacion_tipo.BoundText = "REC" Or DtCDenominacion_tipo.BoundText = "DYR" Or DtCDenominacion_tipo.BoundText = "ANI" Or DtCDenominacion_tipo.BoundText = "DVI" Then
'           rstdestino("cta_codigo") = DtCCta_codigo.Text
'           rstdestino("Codigo_beneficiario") = ""
'         End If
'
'         rstdestino("numero_documento") = TxtNumero_documento.Text
'         rstdestino("monto_dolares") = Round(Txtmonto_dolares.Text, 2)
'         rstdestino("monto_bolivianos") = Round(TxtMonto_bolivianos.Text, 2)
'         rstdestino("usr_usuario") = GlUsuario
'         rstdestino("fecha_registro") = Date
'         rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
'
'         rstdestino("estado_aprobacion") = "N"
'         rstdestino("monto_recaudado_dolares") = 0
'         If v_añadir = 1 Then
'            rstdestino("ultimo") = "S"
'         End If
'         rstdestino.Update
'         If rstdestino.State = 1 Then rstdestino.Close
'        db.CommitTrans
'
''          If rs_datos.State = 1 Then rs_datos.Close
''          rs_datos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
''          rs_datos.Sort = "correlativo_ingreso"
'          rs_datos.Requery
'
''          rs_datos.Requery
'          Set Ado_datos.Recordset = rs_datos
'          Ado_datos.Refresh
'          Ado_datos.Recordset.Find "ultimo = 'S'"
'          If Not (Ado_datos.Recordset.EOF) Then
'            marca1 = Ado_datos.Recordset.Bookmark
'            Ado_datos.Recordset("ultimo") = "N"
'            Ado_datos.Recordset.Update
'          End If
''          rs_datos.Find "ultimo = 'S'"
''          If Not (rs_datos.EOF) Then
''            rs_datos("ultimo") = "N"
''            rs_datos.Update
''          End If
'
''          Ado_datos.Recordset.Move marca1 - 1
'
''          marca1 = 0
'      End If
'
'      If v_añadir = 2 Then
'        '===== modifica un registro =====
'         Set rstdestino = New ADODB.Recordset
'         If rstdestino.State = 1 Then rstdestino.Close
'         rstdestino.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = '" & Ado_datos.Recordset("correlativo_ingreso") & "' and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "' and ges_gestion = '" & Ado_datos.Recordset("ges_gestion") & "' order by correlativo_ingreso, org_codigo ", db, adOpenDynamic, adLockOptimistic
'         rstdestino.MoveFirst
'         If Not (rstdestino.EOF) Then
''            If rstdestino("org_codigo") <> DtCOrg_codigo.Text Then
''              Call add_correl
''              rstdestino("Correlativo_ingreso") = correlativo1
''              rstdestino("correlativo_anterior") = correlativo1
''            End If
'            rstdestino("Codigo_solicitud") = TxtCodigo_solicitud.Text
'            rstdestino("rbr_codigo") = DtCrbr_codigo.Text
'            rstdestino("tipo_moneda") = DtCDenominacion_moneda.BoundText
'            rstdestino("Codigo_tipo_solicitud") = IIf(DtCDenominacion_tipo_solicitud.BoundText = "", 0, DtCDenominacion_tipo_solicitud.BoundText)
'            rstdestino("Codigo_documento") = DtCCodigo_documento.Text
'            rstdestino("Codigo_tipo") = IIf(DtCDenominacion_tipo.BoundText = "", "", DtCDenominacion_tipo.BoundText)
'            rstdestino("Fecha_Ingreso") = DTPFecha_Ingreso.Value
'            rstdestino("Tipo_Cambio") = TxtTipo_cambio.Text
'            rstdestino("Concepto") = TxtConcepto.Text
'            rstdestino("UNI_CODIGO") = dtc_codigo1
'            rstdestino("fte_codigo") = DtCFte_codigo.Text
'            rstdestino("org_codigo") = DtCorg_codigo.Text
'            rstdestino("codigo_convenio") = DtCcodigo_convenio
''            rstdestino("cta_codigo") = DtCCta_codigo.Text
'             If DtCDenominacion_tipo.BoundText = "DEV" Then
'               rstdestino("Codigo_beneficiario") = dtc_codigo4.Text
'               rstdestino("cta_codigo") = ""
'             End If
'
'             If DtCDenominacion_tipo.BoundText = "REC" Or DtCDenominacion_tipo.BoundText = "DYR" Or DtCDenominacion_tipo.BoundText = "ANI" Then
'               rstdestino("cta_codigo") = DtCCta_codigo.Text
'               rstdestino("Codigo_beneficiario") = ""
'             End If
'
'            rstdestino("numero_documento") = TxtNumero_documento.Text
'            rstdestino("monto_dolares") = Round(Txtmonto_dolares.Text, 2)
'            rstdestino("monto_bolivianos") = Round(TxtMonto_bolivianos.Text, 2)
'            If DtCDenominacion_tipo.BoundText = "DEV" Then
'             rstdestino("estado_devengado") = "N"
'             rstdestino("estado_recaudado") = ""
'             rstdestino("estado_desafectado") = ""
'            End If
'            If DtCDenominacion_tipo.BoundText = "REC" Then
'             rstdestino("estado_recaudado") = "N"
'             rstdestino("estado_devengado") = ""
'             rstdestino("estado_desafectado") = ""
'            End If
'            If DtCDenominacion_tipo.BoundText = "DYR" Then
'             rstdestino("estado_recaudado") = "N"
'             rstdestino("estado_devengado") = "N"
'             rstdestino("estado_desafectado") = ""
'            End If
'            rstdestino("estado_Aprobacion") = "N"
'            rstdestino("ultimo") = "N"
'            rstdestino("usr_usuario") = GlUsuario
'            rstdestino("fecha_registro") = Date
'            rstdestino("hora_registro") = Left(CStr(Time()), 8)
'            rstdestino.Update
'            If rstdestino.State = 1 Then rstdestino.Close
'
'            marca1 = Ado_datos.Recordset.Bookmark
'            rs_datos.CancelUpdate
'            rs_datos.Requery
''            rs_datos.Sort = "correlativo_ingreso"
'            Set Ado_datos.Recordset = rs_datos
''            Ado_datos.Refresh
'            Ado_datos.Recordset.Move marca1 - 1
'         End If
''         marca1 = 0
'      End If
'   Else
'      MsgBox "ERROR Los datos no están completos, no se realizará la grabación..."
''      FraOpciones2.Visible = False
''      FraOpciones.Visible = True
''      FraNavega.Enabled = True
''      fraDatos.Enabled = False
''      Ado_datos.Refresh
'   End If
'   LblAccion = ""
'End Sub
'
'Private Sub Cmdimprimir_Click()
'  If rs_datos.RecordCount > 0 Then
'    '===== Ini comando para iniciar impresión
'    Call prt_cmbteIng(Ado_datos.Recordset!ges_gestion, Ado_datos.Recordset!org_codigo, Ado_datos.Recordset!correlativo_ingreso)
'  Else
'    MsgBox "No existen registros para imprimir", vbInformation + vbOKOnly, "ERROR de impresión"
'  End If
'
'End Sub
'
''Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''  CmdImprimir.Picture = LoadPicture("d:\Pragma\iconos\print2.ico") 'LoadPicture("C:\Mis documentos\star.ani")
'''  CmdImprimir.BackColor = &HC0FFFF
''
''End Sub
'
'Private Sub BtnModificar_Click()
'    LblAccion = "Modificando registro..."
'    v_añadir = 2
'    fraOpciones.Visible = False
'    FraOpciones2.Visible = True
'    FraNavega.Enabled = False
'    fraDatos.Enabled = True
'    DtCFte_codigo.Enabled = False
'    DtCorg_codigo.Enabled = False
'    swmodificar = 1
'    Call pfil_cta_Fte(DtCFte_codigo.Text, 3) 'Call pfil_cta_Fte(DtCFte_codigo.Text, 1)
'    Call pfil_conv(DtCFte_codigo, DtCorg_codigo.Text)
'    If Ado_datos.Recordset("Codigo_tipo") = "REC" Or Ado_datos.Recordset("Codigo_tipo") = "DYR" Or Ado_datos.Recordset("Codigo_tipo") = "ANI" Or Ado_datos.Recordset("Codigo_tipo") = "DVI" Then
'      DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, "", Ado_datos.Recordset("Cta_Codigo"))
'      DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
'    End If
'
'    If swcopiar = 1 Then
''      If marca1 = 0 Then
'         marca1 = Ado_datos.Recordset.Bookmark
'      Else
'        marca1 = Ado_datos.Recordset.Bookmark
''        Set Ado_datos.Recordset = rs_datos
''        Ado_datos.Refresh
''        Ado_datos.Recordset.Move marca1 - 1
'      End If
''    Else
''      marca1 = Ado_datos.Recordset.Bookmark
''    End If
'
''    If V_accion = "COPIA" Then
''      If marca1 = 0 Then
''         marca1 = Ado_datos.Recordset.Bookmark
''      Else
''        Set Ado_datos.Recordset = rs_datos
''        Ado_datos.Refresh
''        Ado_datos.Recordset.Move marca1 - 1
''      End If
''    Else
''      marca1 = Ado_datos.Recordset.Bookmark
''    End If
'
'  Call desactivar_Obj
'    correlativo_ingreso1 = Ado_datos.Recordset("correlativo_ingreso")
'    ges_gestion1 = Ado_datos.Recordset("ges_gestion")
'End Sub
'
'Private Sub CmdSalir_Click()
'   sino = MsgBox("¿Está seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
'   If sino = vbYes Then
'     Call cerrar
'
'  If rstFte_financia.State = 1 Then rstFte_financia.Close
'  If AdoFte_financia.Recordset.State = 1 Then AdoFte_financia.Recordset.Close
'  If rs_datos.RecordCount > 0 Then
'    rs_datos.Update
'  End If
'  If rs_datos.State = 1 Then rs_datos.Close
'     Unload Me
'   End If
'End Sub
'
'Private Sub CommandButton1_Click()
'
'End Sub
'
''Private Sub Cmdsalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''  Cmdsalir.Picture = LoadPicture("d:\Pragma\iconos\salir1.ico")
''End Sub
'Private Sub DtCcodigo_convenio_Click(Area As Integer)
'  DtCDenominacion_Convenio.Text = DtCcodigo_convenio.BoundText
'End Sub
'
'Private Sub DtCCodigo_documento_Click(Area As Integer)
'    DtCDenominacion_documento.Text = DtCCodigo_documento.BoundText
'End Sub
'
'Private Sub DtCCta_codigo_Click(Area As Integer)
'   DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
'End Sub
'
'Private Sub DtCCta_descripcion_larga_Click(Area As Integer)
'   DtCCta_codigo.Text = DtCCta_descripcion_larga.BoundText
'End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
   dtc_desc4.Text = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
  dtc_codigo4.Text = dtc_desc4.BoundText
End Sub

'Private Sub DtCDenominacion_Convenio_Click(Area As Integer)
'  DtCcodigo_convenio.Text = DtCDenominacion_Convenio.BoundText
'End Sub
'
'Private Sub DtCDenominacion_documento_Click(Area As Integer)
'    DtCCodigo_documento = DtCDenominacion_documento.BoundText
'End Sub
'
'Private Sub DtCDenominacion_tipo_Click(Area As Integer)
'
'  If DtCDenominacion_tipo = "DEVENGADO" Then
'    DtCCta_codigo.Visible = False
'    DtCCta_descripcion_larga.Visible = False
'    Lblcuenta.Visible = False
'    dtc_codigo4.Visible = True
'    dtc_desc4.Visible = True
'    lblBeneficiario.Visible = True
'  End If
'
'  If (DtCDenominacion_tipo = "DEVENGADO Y RECAUDADO") Or (DtCDenominacion_tipo = "RECAUDADO") Or (DtCDenominacion_tipo = "ANULADO") Then
'    DtCCta_codigo.Visible = True
'    DtCCta_descripcion_larga.Visible = True
'    Lblcuenta.Visible = True
'    dtc_codigo4.Visible = False
'    dtc_desc4.Visible = False
'    lblBeneficiario.Visible = False
'  End If
'
''3 codigo_beneficiario varchar 15  0 0 0   0     0
''0 denominacion_beneficiario varchar 60  0 0 1   0     0
''0 tipo_beneficiario varchar 1 0 0 1   0     0
'
'End Sub
'
'Private Sub DtCDenominacion_tipo_solicitud_KeyPress(KeyAscii As Integer)
'  ' aqui cambiar de lugar
'  If KeyAscii = 13 Then
'
'  End If
'End Sub

Private Sub DtCOrg_codigo_Click(Area As Integer)
  DtCOrg_descripcion.BoundText = DtCorg_codigo.BoundText
  Call pfil_cta_Fte(Me.DtCorg_codigo, 3) 'Call pfil_cta_Fte(Me.DtCOrg_codigo, 2)
  'Call pfil_conv(DtCFte_codigo, DtCorg_codigo.Text)
End Sub

Private Sub DtCOrg_descripcion_Click(Area As Integer)
  DtCorg_codigo.BoundText = DtCOrg_descripcion.BoundText
  Call pfil_cta_Fte(Me.DtCorg_codigo, 3) 'Call pfil_cta_Fte(Me.DtCOrg_codigo, 2)
End Sub

Private Sub DtCrbr_codigo_Click(Area As Integer)
   DtCrbr_descripcion.BoundText = DtCrbr_codigo.BoundText
End Sub

Private Sub DtCrbr_descripcion_Click(Area As Integer)
    DtCrbr_codigo.BoundText = DtCrbr_descripcion.BoundText
End Sub

Private Sub DtCFte_codigo_Click(Area As Integer)
    DtCFte_descripcion_larga.BoundText = DtCFte_codigo.BoundText
    DtCorg_codigo.Enabled = True
    Call pfil_Org_Fte(DtCFte_codigo.Text)
    Call pfil_cta_Fte(Me.DtCorg_codigo, 3) 'Call pfil_cta_Fte(Me.DtCFte_codigo, 1)
    'Call pfil_conv(DtCFte_codigo, "")
End Sub

Private Sub DtCFte_descripcion_larga_Click(Area As Integer)
  DtCFte_codigo.BoundText = DtCFte_descripcion_larga.BoundText
  Call pfil_Org_Fte(DtCFte_descripcion_larga.BoundText)
  Call pfil_cta_Fte(Me.DtCorg_codigo, 3) 'Call pfil_cta_Fte(DtCFte_descripcion_larga.BoundText, 1)
End Sub

Private Sub DtCtipo_Comp_Click(Area As Integer)
    DtCDenominacion_tipo.BoundText = DtCtipo_Comp.BoundText
End Sub

Private Sub DtCtipo_solicitud_Click(Area As Integer)
    DtCDenominacion_tipo_solicitud.BoundText = DtCtipo_solicitud.BoundText
End Sub

'
Private Sub Form_Load()
  swgraba = 0
  marca1 = 0
  swcopiar = 0
  V_accion = "COPIA"
    
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
 If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
    Ado_datos.Recordset.MoveFirst
    DtCFte_codigo.Text = IIf(IsNull(Ado_datos.Recordset("fte_codigo")) = True, " ", Ado_datos.Recordset("fte_codigo"))
    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
    DtCorg_codigo.Text = IIf(IsNull(Ado_datos.Recordset("org_codigo")) = True, " ", Ado_datos.Recordset("org_codigo"))
    DtCOrg_descripcion.Text = DtCorg_codigo.BoundText
    DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, " ", Ado_datos.Recordset("Cta_Codigo"))
    DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
  End If
  '===== fin cargado de tablas de consulta y de datos de despliegue
  TxtTipo_cambio = GlTipoCambioOficial
End Sub

Private Sub ABRIR_TABLAS_AUX()
  '===== Ini cargado de tablas de consulta y de datos de despliegue
  'LblUsuario.Caption = LblUsuario.Caption + GlUsuario
    
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
  
  
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

  Set rstFc_Rubro_ingresos = New ADODB.Recordset
  If rstFc_Rubro_ingresos.State = 1 Then rstFc_Rubro_ingresos.Close
  rstFc_Rubro_ingresos.Open "select * from fc_ingresos_rubro order by rubro_codigo", db, adOpenKeyset, adLockReadOnly
  Set AdoFc_Rubro_ingresos.Recordset = rstFc_Rubro_ingresos
  AdoFc_Rubro_ingresos.Refresh
  If Not AdoFc_Rubro_ingresos.Recordset.BOF Then AdoFc_Rubro_ingresos.Recordset.MoveFirst

  Set rstTipo_solicitud = New ADODB.Recordset
  If rstTipo_solicitud.State = 1 Then rstTipo_solicitud.Close
  rstTipo_solicitud.Open "select * from gc_tipo_solicitud order by solicitud_tipo_descripcion", db, adOpenKeyset, adLockReadOnly
  Set AdoTipo_solicitud.Recordset = rstTipo_solicitud
  AdoTipo_solicitud.Refresh
  If Not AdoTipo_solicitud.Recordset.BOF Then AdoTipo_solicitud.Recordset.MoveFirst

  Set rstTipo_comprobante = New ADODB.Recordset
  If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
  rstTipo_comprobante.Open "select * from gc_tipo_comprobante where ingresos = 'I' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
  Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
  AdoTipo_comprobante.Refresh
  If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
    
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "Select * from gc_beneficiario where (tipoben_codigo < 20 and tipoben_codigo <> 1) order by beneficiario_denominacion", db, adOpenStatic
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText

'  Set rstfc_beneficiario = New ADODB.Recordset
'  If rstfc_beneficiario.State = 1 Then rstfc_beneficiario.Close
'  rstfc_beneficiario.Open "SELECT * from gc_beneficiario order by beneficiario_codigo", db, adOpenStatic, adLockReadOnly
'  Set AdoFc_beneficiario.Recordset = rstfc_beneficiario
'  AdoFc_beneficiario.Refresh
  
  Set rstTipo_moneda = New ADODB.Recordset
  If rstTipo_moneda.State = 1 Then rstTipo_moneda.Close
  rstTipo_moneda.Open "select * from gc_tipo_moneda order by tipo_moneda_descripcion", db, adOpenKeyset, adLockReadOnly
  Set AdoTipo_moneda.Recordset = rstTipo_moneda
  AdoTipo_moneda.Refresh
  If Not AdoTipo_moneda.Recordset.BOF Then AdoTipo_moneda.Recordset.MoveFirst
  
'  Set rstFc_convenios = New ADODB.Recordset
'  If rstFc_convenios.State = 1 Then rstFc_convenios.Close
'  rstFc_convenios.Open " select * from Fc_convenios order by codigo_convenio ", db, adOpenKeyset, adLockReadOnly
'  Set AdoFc_convenios.Recordset = rstFc_convenios
'  AdoFc_convenios.Refresh

  If rstFc_cuenta_bancaria.State = 1 Then rstFc_cuenta_bancaria.Close
  rstFc_cuenta_bancaria.Open "Select * from fc_cuenta_bancaria order by cta_codigo", db, adOpenDynamic, adLockReadOnly
  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
  AdoFc_cuenta_bancaria.Refresh
  If Not AdoFc_cuenta_bancaria.Recordset.BOF Then AdoFc_cuenta_bancaria.Recordset.MoveFirst

'  If rstac_documento_respaldo.State = 1 Then rstac_documento_respaldo.Close
'  Set rstac_documento_respaldo = New ADODB.Recordset
'  rstac_documento_respaldo.Open "select * from gc_documentos_respaldo", db, adOpenDynamic, adLockReadOnly
'  Set Adoac_documento_respaldo.Recordset = rstac_documento_respaldo
'  Adoac_documento_respaldo.Refresh
'  If Not Adoac_documento_respaldo.Recordset.BOF Then Adoac_documento_respaldo.Recordset.MoveFirst


    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    'rs_datos6.Open "Select * from gc_proceso_nivel2 order by subproceso_descripcion", db, adOpenStatic
    rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
          
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos9.Close
    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
'  Set rs_datos = New ADODB.Recordset
'  ' pa busqueda QueryInicial = "select * from fo_ingresos_cabecera where estado_aprobacion <> 'S'" 'ORDER BY correlativo_ingreso , org_codigo
'  queryinicial = "select * from fo_ingresos_cabecera where estado_aprobacion <> 'S' and estado_aprobacion <> 'E'" ' ORDER BY correlativo_ingreso , org_codigo"
'  If rs_datos.State = 1 Then rs_datos.Close
''pa busqueda  rs_datos.Open QueryInicial, db, adOpenKeyset, adLockOptimistic
'  rs_datos.Open queryinicial & " ORDER BY correlativo_ingreso , org_codigo ", db, adOpenDynamic, adLockOptimistic
''pa busqueda  rs_datos.Sort = "correlativo_ingreso"
'  Set Ado_datos.Recordset = rs_datos

'  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
'    Ado_datos.Recordset.MoveFirst
'    DtCFte_codigo.Text = IIf(IsNull(Ado_datos.Recordset("fte_codigo")) = True, " ", Ado_datos.Recordset("fte_codigo"))
'    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
'    DtCorg_codigo.Text = IIf(IsNull(Ado_datos.Recordset("org_codigo")) = True, " ", Ado_datos.Recordset("org_codigo"))
'    DtCOrg_descripcion.Text = DtCorg_codigo.BoundText
'    DtCCta_codigo.Text = IIf(IsNull(Ado_datos.Recordset("Cta_Codigo")) = True, " ", Ado_datos.Recordset("Cta_Codigo"))
'    DtCCta_descripcion_larga.Text = DtCCta_codigo.BoundText
'  End If
'  '===== fin cargado de tablas de consulta y de datos de despliegue
'  TxtTipo_cambio = GlTipoCambioOficial
End Sub

'
''Private Sub FraOpciones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''  CmdImprimir.Picture = LoadPicture("d:\Pragma\iconos\print3.ico") 'LoadPicture("C:\Mis documentos\star.ani")
''  CmdBuscar.Picture = LoadPicture("d:\Pragma\iconos\busca3.ico") 'LoadPicture("C:\Mis documentos\star.ani")
''  CmdSalir.Picture = LoadPicture("d:\Pragma\iconos\salir3.ico")
'''  CmdBuscar.BackColor = &H8000000F
'''  CmdImprimir.BackColor = &H8000000F
'''  CmdAñadir.BackColor = &H8000000F
'''  Image1.Visible = False
''
''End Sub
'
Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From fo_ingresos_cabecera WHERE estado_codigo = 'REG' "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub
'
Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
  
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From fo_ingresos_cabecera "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Sub abrir()
'  If rs_datos.State = 1 Then rs_datos.Close
'  Set rs_datos = New ADODB.Recordset
'  rs_datos.Open "select * from fo_ingresos_cabecera order by correlativo_ingreso, org_codigo ", db, adOpenDynamic, adLockOptimistic
'  Set Ado_datos.Recordset = rs_datos
'  If Ado_datos.Recordset.State = 1 Then Ado_datos.Recordset.Close
'  Ado_datos.Refresh
'  dg_datos.Refresh
'  If Not rs_datos.BOF Then rs_datos.MoveFirst
'
'  If rstFte_financia.State = 1 Then rstFte_financia.Close
'  rstFte_financia.Open "Select * from Fc_fuente_financiamiento", db, adOpenDynamic, adLockReadOnly
'  Set AdoFte_financia.Recordset = rstFte_financia
'  AdoFte_financia.Refresh
'  If Not rstFte_financia.BOF Then rstFte_financia.MoveFirst
'
'  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
'  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento", db, adOpenDynamic, adLockReadOnly
'  Set AdoOrganismo_finan.Recordset = rstOrganismo_finan
'  AdoOrganismo_finan.Refresh
'  If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
'
'  If rstFc_cuenta_bancaria.State = 1 Then rstFc_cuenta_bancaria.Close
'  rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria order by cta_codigo", db, adOpenDynamic, adLockReadOnly
'  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
'  AdoFc_cuenta_bancaria.Refresh
'  If Not rstFc_cuenta_bancaria.BOF Then rstFc_cuenta_bancaria.MoveFirst
'
'  If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
'
'  End If
'End Sub
'
'Sub cerrar()
''  If rstFte_financia.State = 1 Then rstFte_financia.Close
''  If AdoFte_financia.Recordset.State = 1 Then AdoFte_financia.Recordset.Close
''  If Ado_datos.Recordset.State = 1 Then Ado_datos.Recordset.Close
''  If rs_datos.State = 1 Then rs_datos.Close
'End Sub
'
'Private Sub Txtduracion_estimada_KeyPress(KeyAscii As Integer)
'  If KeyAscii < 58 And KeyAscii > 47 Or KeyAscii = 8 Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub elimina()
''===== proceso para eliminar registros
'  Dim rstelimina As New ADODB.Recordset
'  If rstelimina.State = 1 Then rstelimina.Close
'  Set rstelimina = New ADODB.Recordset
'  If rstelimina.State = 1 Then rstelimina.Close
'  marca1 = Ado_datos.Recordset.Bookmark
'  rstelimina.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("Correlativo_ingreso") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
'  If (Not rstelimina.BOF) Then rstelimina.MoveFirst
'  '  rs_datos.Find "Correlativo_ingreso= '" & Ado_datos.Recordset("Correlativo_ingreso") & "'", , adSearchForward
'  '  If Not rs_datos.BOF Then
'    If rstelimina!estado_devengado = "N" Then rstelimina!estado_devengado = "E"
'    If rstelimina!estado_recaudado = "N" Then rstelimina!estado_recaudado = "E"
'    If rstelimina!estado_desafectado = "N" Then rstelimina!estado_desafectado = "E"
'    If rstelimina!estado_anulado = "N" Then rstelimina!estado_anulado = "E"
'    rstelimina!estado_aprobacion = "E"
'    rstelimina.Update
'  '  End If
'  If rstelimina.State = 1 Then rstelimina.Close
'  rs_datos.Update
'  rs_datos.Requery
'  Set Ado_datos.Recordset = rs_datos
'  Ado_datos.Refresh
'  Set Ado_datos.Recordset = rs_datos
'  If rs_datos.RecordCount > 0 Then
'    Ado_datos.Recordset.Move marca1 - 1
'  End If
'End Sub
'
'Private Sub errado()
'''===== proceso para eliminar registros
''  Dim rsterrado As New ADODB.Recordset
''  If rsterrado.State = 1 Then rsterrado.Close
''  Set rsterrado = New ADODB.Recordset
''  rsterrado.Open "select * from fo_ingresos_cabecera where correlativo_ingreso = " & Ado_datos.Recordset("Correlativo_ingreso") & " and org_codigo = '" & Ado_datos.Recordset("org_codigo") & "'", db, adOpenKeyset, adLockOptimistic
''  If (Not rsterrado.BOF) Then rsterrado.MoveFirst
'''  rsterrado.Find "Correlativo_ingreso= '" & Ado_datos.Recordset("Correlativo_ingreso") & "'", , adSearchForward
'''  If Not rsterrado.BOF Then
''    If rsterrado("estado_devengado") = "N" Then
''      rsterrado("estado_devengado") = "E"
''      rsterrado("estado_aprobacion") = "E"
''    End If
''    If rsterrado("estado_recaudado") = "N" Then
''      rsterrado("estado_recaudado") = "E"
''      rsterrado("estado_aprobacion") = "E"
''    End If
''    If rsterrado("estado_desafectado") = "N" Then
''      rsterrado("estado_desafectado") = "E"
''      rsterrado("estado_aprobacion") = "E"
''    End If
''
''    rsterrado.Update
'''  End If
''  If rsterrado.State = 1 Then rsterrado.Close
''  rs_datos.Update
''  rs_datos.Requery
''  Set Ado_datos.Recordset = rs_datos
''  Ado_datos.Refresh
'End Sub
'
'Private Sub valida()
''===== Validación para grabar datos
'  swgraba = 1
'  If Len(Trim(TxtCodigo_solicitud)) < 1 Then swgraba = 0
'  If IsNull(DTPFecha_Ingreso) Then swgraba = 0
'  If TxtTipo_cambio = 0 Then swgraba = 0
'  If Len(Trim(TxtConcepto)) < 1 Then swgraba = 0
'  If Len(Trim(Txtmonto_dolares)) < 1 Then swgraba = 0
'  If Len(Trim(TxtMonto_bolivianos.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCrbr_codigo.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCDenominacion_moneda.Text)) < 1 Then swgraba = 0
'  If Len(Trim(dtc_codigo1.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCDenominacion_tipo_solicitud.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCCodigo_documento.Text)) < 1 Then swgraba = 0
'  If Len(Trim(TxtConcepto.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCFte_codigo.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCorg_codigo.Text)) < 1 Then swgraba = 0
'  If Len(Trim(DtCcodigo_convenio.Text)) < 1 Then swgraba = 0
'  If DtCDenominacion_tipo.BoundText = "DEV" Then
'    If (Len(Trim(dtc_codigo4.Text)) < 1) Then swgraba = 0
'  End If
'  If (DtCDenominacion_tipo.BoundText = "REC") Or (DtCDenominacion_tipo.BoundText = "DYR") Or (DtCDenominacion_tipo.BoundText = "ANI") Or (DtCDenominacion_tipo.BoundText = "DVI") Then
'    If (Len(Trim(DtCCta_codigo.Text)) < 1) Then swgraba = 0
'  End If
'  If Len(Trim(TxtNumero_documento.Text)) < 1 Then swgraba = 0
'
'End Sub
'
'Private Sub TxtCodigo_beneficiario_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'
'End Sub
'
'Private Sub TDBNumber1_Click()
'
'End Sub
'
'Private Sub TxtCodigo_solicitud_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub TxtMonto_bolivianos_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
'       Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, Round(TxtMonto_bolivianos.Text / TxtTipo_cambio, 2), 0)
'    Else
'       Txtmonto_dolares.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'End Sub
'
'Private Sub Txtmonto_dolares_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
'      TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Round(Txtmonto_dolares * TxtTipo_cambio, 2), 0)
'    Else
'      TxtMonto_bolivianos.Text = 0
'    End If
'  End If
'End Sub
'
'Private Sub TxtNumero_documento_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtTipo_moneda_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub TxtTipo_solicitud_KeyPress(KeyAscii As Integer)
'  KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub pfil_Org_Fte(Codfte As String)
'===== Proceso para filtrar los Organismos en base a la Fuente de financiamiento
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento where fte_codigo = '" & Codfte & "'", db, adOpenDynamic, adLockReadOnly
  'If rstOrganismo_finan.RecordCount < 1 Then
    DtCorg_codigo.Text = ""
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
      rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria where fte_codigo = '" & cod & "' or cta_codigo = '01' order by cta_codigo ", db, adOpenDynamic, adLockReadOnly
    Case 2
      rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria where org_codigo = '" & cod & "' or cta_codigo = '01' order by cta_codigo ", db, adOpenDynamic, adLockReadOnly
    Case 3
      rstFc_cuenta_bancaria.Open "Select * from Fc_cuenta_bancaria order by cta_codigo ", db, adOpenDynamic, adLockReadOnly
  End Select
  Me.DtCCta_codigo.Text = ""
  Me.DtCCta_descripcion_larga.Text = ""
  Set AdoFc_cuenta_bancaria.Recordset = rstFc_cuenta_bancaria
  AdoFc_cuenta_bancaria.Refresh
  If Not AdoFc_cuenta_bancaria.Recordset.BOF Then AdoFc_cuenta_bancaria.Recordset.MoveFirst
End Sub
'
'Private Sub pfil_conv(Codfte, codorg As String)
''===== Proceso para filtrar los Cojnvenios en base a la Fuente de financiamiento y el organismo
'  'If Len(Trim(Codfte)) > 0 And Len(Trim(codorg)) > 0 Then
'    If rstFc_convenios.State = 1 Then rstFc_convenios.Close
'    rstFc_convenios.Open "Select * from Fc_convenios where fte_codigo = '" & Codfte & "' and org_codigo = '" & codorg & "' ", db, adOpenDynamic, adLockReadOnly
'    'If rstOrganismo_finan.RecordCount < 1 Then
'      DtCcodigo_convenio.Text = ""
'      DtCDenominacion_Convenio.Text = ""
'    'End If
'    Set AdoFc_convenios.Recordset = rstFc_convenios
'    AdoFc_convenios.Refresh
'    'If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
'  'Else
'
'  'End If
'End Sub
'
''Private Sub Cmdbuspri_Click()
'''===== Proceso para buscar el primer registro en base al criterio seleccionado
''  Call parametros
''  If buscasi = 1 Then
''    If (Not Ado_datos.Recordset.BOF) Then Ado_datos.Recordset.MoveFirst
''    If operadorbus = "=" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchForward
''      If Ado_datos.Recordset.EOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''    If operadorbus = "like" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchForward
''      If Ado_datos.Recordset.EOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''  End If
''  buscasi = 0
''End Sub
'
''Private Sub Cmdbussig_Click()
'''===== Proceso para buscar el siguiente registro en base al criterio seleccionado
''  Call parametros
''  If buscasi = 1 Then
''    If (Not Ado_datos.Recordset.EOF) Then Ado_datos.Recordset.MoveNext
''    If operadorbus = "=" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchForward
''      If Ado_datos.Recordset.EOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''    If operadorbus = "like" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchForward
''      If Ado_datos.Recordset.EOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''  End If
''  buscasi = 0
''End Sub
''
''Private Sub Cmdbusult_Click()
'''===== Proceso para buscar el último registro en base al criterio seleccionado
''  Call parametros
''  If buscasi = 1 Then
''    If (Not Ado_datos.Recordset.EOF) Then Ado_datos.Recordset.MoveLast
''    If operadorbus = "=" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchBackward
''      If Ado_datos.Recordset.BOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''    If operadorbus = "like" Then
''      Ado_datos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchBackward
''      If Ado_datos.Recordset.EOF Then
''        MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''        Ado_datos.Recordset.MoveFirst
''      End If
''    End If
''  End If
''  buscasi = 0
''End Sub
''
''Private Sub CmdBusAnt_Click()
'''===== Proceso para buscar el anterior registro en base al criterio seleccionado
''  Call parametros
''  If buscasi = 1 Then
''    If (Not Ado_datos.Recordset.BOF) Then Ado_datos.Recordset.MovePrevious
''    If (Not Ado_datos.Recordset.BOF) Then
''      If operadorbus = "=" Then
''        Ado_datos.Recordset.Find campobus & " " & operadorbus & " '" & Trim(Txtvarbus) & "'", , adSearchBackward
''        If Ado_datos.Recordset.BOF Then
''          MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''          Ado_datos.Recordset.MoveFirst
''        End If
''      End If
''      If operadorbus = "like" Then
''        Ado_datos.Recordset.Find campobus & " " & operadorbus & " '*" & Trim(Txtvarbus) & "*'", , adSearchBackward
''        If Ado_datos.Recordset.BOF Then
''          MsgBox "Parámetro no encontrado", vbCritical + vbOKOnly, "Error de Búsqueda"
''          Ado_datos.Recordset.MoveFirst
''        End If
''      End If
''    Else
''      MsgBox "Este es el primer registro", vbCritical + vbOKOnly, "Inicio de Registros"
''      Ado_datos.Recordset.MoveFirst
''    End If
''  End If
''  buscasi = 0
''End Sub
'
''Private Sub parametros()
'''===== Proceso para definir los criterios de búsqueda
''  buscasi = 1
''  If Len(Trim(Cmbcampobus.Text)) < 1 Then buscasi = 0
''  If Len(Trim(CmbOperador.Text)) < 1 Then buscasi = 0
''  If Len(Trim(Txtvarbus.Text)) < 1 Then buscasi = 0
''  If buscasi = 1 Then
''    Select Case Trim(Cmbcampobus.Text)
''      Case "Comprobante"
''        campobus = " correlativo_ingreso "
''      Case "Organismo Finan."
''        campobus = " org_codigo "
''      Case "Cuenta"
''        campobus = " cta_codigo "
''      Case "Fecha Ingreso"
''        campobus = " fecha_ingreso "
''        CmbOperador.Text = "="
''      Case "No.Solicitud Desembolso"
''        campobus = " codigo_solicitud "
''      Case Else
''    End Select
''
''    Select Case Trim(CmbOperador.Text)
''      Case "="
''        operadorbus = "="
''      Case "PARTE ="
''        operadorbus = "like"
''      Case Else
''    End Select
''  Else
''    MsgBox "Para poder realizar la búsqueda, por favor debe ingresar todos los parámetros ", vbCritical + vbOKOnly, "ERROR en búsqueda"
''  End If
''End Sub
'
''Private Sub dg_datos_Click()
''    TIPOFORMULARIO = DtcTipoDes.Text
''End Sub
'
'Private Sub dg_datos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'  If Button = vbRightButton Then Me.PopupMenu mnuAcciones
'End Sub
'
'Private Sub mnuAccion_Click(Index As Integer)
'  correlativo_ingreso1 = Ado_datos.Recordset("correlativo_ingreso")
'  Org_Codigo1 = Ado_datos.Recordset("org_codigo")
'  Select Case Index
'    Case 0 ' RECAUDADO          ' Devengado
'      'if Ado_datos.Recordset("estado_reversion_total")="S" then
'      MsgBox "Realizando el RECAUDADO", vbInformation + vbOKOnly, "Atención"
'      V_accion = "REC"
'      CmdCopiar_Click
'    Case 1 'DESAFECTADO         ' Reversión
'      MsgBox "Realizando la Desafección", vbInformation + vbOKOnly, "Atención"
'      V_accion = "DES"
'      CmdCopiar_Click
'    Case 2 'ANULAR RECAUDADO    ' anulacion
'      MsgBox "Realizando la Anulación de lo Recaudado", vbInformation + vbOKOnly, "Atención"
'      V_accion = "ANI"
'      CmdCopiar_Click
'    Case 3 'DEVOLUCION          ' Devolución
'      MsgBox "Realizando la Devolución de Devengado y Recaudado", vbInformation + vbOKOnly, "Atención"
'      V_accion = "DVI"
'      CmdCopiar_Click
'  End Select
'End Sub
'
'Private Sub CmdCopiar_Click()
'  v_añadir = 1
'  swcopiar = 1
'  CmdAñadir_Click
''  CmdGrabar_Click
''  BtnModificar_Click
'
''last  V_accion = "COPIA"
'  swcopiar = 0
'End Sub
'
'Private Sub desactivar_Obj()
'  If (var_devuelto = "DVI") Then
'    TxtTipo_cambio.Enabled = True
'    TxtConcepto.Enabled = True
'    Txtmonto_dolares.Enabled = True
'    TxtMonto_bolivianos.Enabled = True
'    DtCrbr_codigo.Enabled = False
'    DtCrbr_descripcion.Enabled = False
'    DtCDenominacion_moneda.Enabled = False
'    DtCDenominacion_tipo_solicitud.Enabled = False
'    DtCCodigo_documento.Enabled = True
'    DtCDenominacion_documento.Enabled = True
'    DtCFte_codigo.Enabled = False
'    DtCFte_descripcion_larga.Enabled = False
'    DtCorg_codigo.Enabled = False
'    DtCOrg_descripcion.Enabled = False
'    DtCCta_codigo.Enabled = True
'    DtCCta_descripcion_larga.Enabled = True
'    DtCDenominacion_tipo.Enabled = False
'    TxtNumero_documento.Enabled = True
'    TxtCodigo_solicitud.Enabled = True
'  Else
'    TxtTipo_cambio.Enabled = False
'    TxtConcepto.Enabled = False
'    Txtmonto_dolares.Enabled = False
'    TxtMonto_bolivianos.Enabled = False
'    DtCrbr_codigo.Enabled = False
'    DtCrbr_descripcion.Enabled = False
'    DtCDenominacion_moneda.Enabled = False
'    DtCDenominacion_tipo_solicitud.Enabled = False
'    DtCCodigo_documento.Enabled = False
'    DtCDenominacion_documento.Enabled = False
'    DtCFte_codigo.Enabled = False
'    DtCFte_descripcion_larga.Enabled = False
'    DtCorg_codigo.Enabled = False
'    DtCOrg_descripcion.Enabled = False
'    DtCCta_codigo.Enabled = False
'    DtCCta_descripcion_larga.Enabled = False
'    DtCDenominacion_tipo.Enabled = False
'    TxtNumero_documento.Enabled = False
'    TxtCodigo_solicitud.Enabled = False
'  End If
'  Select Case Ado_datos.Recordset("codigo_tipo")
'    Case "DYR"
'      DtCCta_codigo.Enabled = True
'      DtCCta_descripcion_larga.Enabled = True
'      DtCDenominacion_moneda.Enabled = True
'      Txtmonto_dolares.Enabled = True
'      TxtMonto_bolivianos.Enabled = True
'      TxtConcepto.Enabled = True
'      DtCrbr_codigo.Enabled = True
'      DtCrbr_descripcion.Enabled = True
'      TxtTipo_cambio.Enabled = True
'
'      DtCDenominacion_tipo_solicitud.Enabled = True
'      DtCCodigo_documento.Enabled = True
'      DtCDenominacion_documento.Enabled = True
'      DtCCta_codigo.Enabled = True
'      DtCCta_descripcion_larga.Enabled = True
'
'      TxtNumero_documento.Enabled = True
'      TxtCodigo_solicitud.Enabled = True
'    Case "REC"
'      DtCCta_codigo.Enabled = True
'      DtCCta_descripcion_larga.Enabled = True
'      DtCDenominacion_moneda.Enabled = True
'      Txtmonto_dolares.Enabled = True
'      TxtMonto_bolivianos.Enabled = True
'      TxtConcepto.Enabled = True
'      DtCrbr_codigo.Enabled = True
'      DtCrbr_descripcion.Enabled = True
'      TxtTipo_cambio.Enabled = True
'
'      DtCDenominacion_tipo_solicitud.Enabled = True
'      DtCCodigo_documento.Enabled = True
'      DtCDenominacion_documento.Enabled = True
'      DtCCta_codigo.Enabled = True
'      DtCCta_descripcion_larga.Enabled = True
'
'      TxtNumero_documento.Enabled = True
'      TxtCodigo_solicitud.Enabled = True
'    Case "ANI"
'      TxtConcepto.Enabled = True
'      TxtTipo_cambio.Enabled = True
'      Txtmonto_dolares.Enabled = True
'      TxtMonto_bolivianos.Enabled = True
'    Case "DES"
'      TxtConcepto.Enabled = True
'      TxtTipo_cambio.Enabled = True
'      Txtmonto_dolares.Enabled = True
'      TxtMonto_bolivianos.Enabled = True
'    Case "DVI"
'      TxtConcepto.Enabled = True
'      TxtTipo_cambio.Enabled = True
'      Txtmonto_dolares.Enabled = True
'      TxtMonto_bolivianos.Enabled = True
'  End Select
'  CmdCopiar.Enabled = False
'End Sub
'
'Private Sub activar_Obj()
'  DtCDenominacion_tipo.Enabled = True
'  CmdCopiar.Enabled = True
'
'  TxtCodigo_solicitud.Enabled = True
''  DTPFecha_Ingreso.Enabled = True
''  TxtTipo_cambio.Enabled = True
'  TxtConcepto.Enabled = True
'  Txtmonto_dolares.Enabled = True
'  TxtMonto_bolivianos.Enabled = True
'  DtCrbr_codigo.Enabled = True
'  DtCrbr_descripcion.Enabled = True
'  DtCDenominacion_moneda.Enabled = True
'  DtCDenominacion_tipo_solicitud.Enabled = True
'  DtCCodigo_documento.Enabled = True
'  DtCDenominacion_documento.Enabled = True
'  DtCFte_codigo.Enabled = True
'  DtCFte_descripcion_larga.Enabled = True
'  DtCorg_codigo.Enabled = True
'  DtCOrg_descripcion.Enabled = True
'  DtCCta_codigo.Enabled = True
'  DtCCta_descripcion_larga.Enabled = True
'  DtCDenominacion_moneda.Enabled = True
'  TxtNumero_documento.Enabled = True
'  TxtCodigo_solicitud.Enabled = True
'  If swcopiar = 1 Then
'    DtCFte_codigo.Enabled = False
'    DtCorg_codigo.Enabled = False
'  Else
'    DtCFte_codigo.Enabled = True
'    DtCorg_codigo.Enabled = True
'  End If
'  If swmodificar = 1 Then
'    DtCFte_codigo.Enabled = False
'    DtCorg_codigo.Enabled = False
'  End If
'End Sub
'
''
'Private Sub add_correl()
'  Dim rstcorrel_ing As New ADODB.Recordset
'  Set rstcorrel_ing = New ADODB.Recordset
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'  rstcorrel_ing.Open "select * from fc_correl_ingresos where org_codigo = '" & Trim(DtCorg_codigo.Text) & "' and ges_gestion = '" & Trim(LblGes_Gestion.Caption) & "'", db, adOpenDynamic, adLockOptimistic
'  If Not (rstcorrel_ing.BOF) Then rstcorrel_ing.MoveFirst
'  rstcorrel_ing.Find "org_codigo = '" & (DtCorg_codigo.Text) & "' ", , adSearchForward
'  If rstcorrel_ing.EOF Then
'     rstcorrel_ing.AddNew
'     rstcorrel_ing("org_codigo") = Trim(DtCorg_codigo.Text)
'     rstcorrel_ing("ges_gestion") = Trim(LblGes_Gestion.Caption)
'     rstcorrel_ing("correlativo") = 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo")
'     FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
'  Else
'     rstcorrel_ing("correlativo") = rstcorrel_ing("correlativo") + 1
'     rstcorrel_ing.Update
'     correlativo1 = rstcorrel_ing("correlativo")
'     'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
'  End If
'  If rstcorrel_ing.State = 1 Then rstcorrel_ing.Close
'
'End Sub
'
'Private Sub f_actual_rec(org, codant)
'  Dim acumDl As Double
'  Dim rsrecalc As New ADODB.Recordset
'  Set rsrecalc = New ADODB.Recordset
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select sum(monto_dolares) as acumDl from fo_ingresos_cabecera where org_codigo = '" & org & "' and  correlativo_anterior = '" & codant & "' and codigo_tipo = 'REC' and estado_recaudado= 'S'", db, adOpenKeyset, adLockReadOnly
'  If rsrecalc.RecordCount > 0 Then
'    acumDl = IIf(IsNull(rsrecalc!acumDl), 0, rsrecalc!acumDl)
'  Else
'    acumDl = 0
'  End If
'  If rsrecalc.State = 1 Then rsrecalc.Close
'  rsrecalc.Open "select * from fo_ingresos_cabecera where org_codigo = '" & org & "' and correlativo_ingreso = '" & codant & "' ", db, adOpenKeyset, adLockOptimistic
'  If rsrecalc.RecordCount > 0 Then
'    rsrecalc!monto_recaudado_dolares = acumDl
'  End If
'  rsrecalc.Update
'  If rsrecalc.State = 1 Then rsrecalc.Close
'
'End Sub
'
'Private Sub TxtTipo_cambio_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If
'
'End Sub
'
'Private Sub TxtTipo_cambio_KeyUp(KeyCode As Integer, Shift As Integer)
'  If Len(TxtTipo_cambio.Text) > 0 Then
'    If DtCDenominacion_moneda.BoundText = "Bs" Then
'      If (Len(Trim(TxtMonto_bolivianos.Text)) > 0) Then
'        Txtmonto_dolares.Text = IIf(TxtMonto_bolivianos.Text > 0, Round(TxtMonto_bolivianos.Text / TxtTipo_cambio, 2), 0)
'      Else
'        Txtmonto_dolares.Text = 0
'      End If
'    Else
'      If Len(Trim(Txtmonto_dolares.Text)) > 0 Then
'        TxtMonto_bolivianos.Text = IIf(Txtmonto_dolares.Text > 0, Round(Txtmonto_dolares * TxtTipo_cambio, 2), 0)
'      Else
'        TxtMonto_bolivianos.Text = 0
'      End If
'    End If
'  End If
'End Sub


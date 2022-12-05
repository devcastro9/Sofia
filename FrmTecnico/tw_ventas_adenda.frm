VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_ventas_adenda 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ventas - Adendas de Contratos"
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   16605
   Icon            =   "tw_ventas_adenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   16605
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      Height          =   1020
      Left            =   120
      ScaleHeight     =   960
      ScaleWidth      =   16320
      TabIndex        =   24
      Top             =   120
      Width           =   16380
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "tw_ventas_adenda.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   36
         ToolTipText     =   "Anula Todo el Tramite"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   14040
         Picture         =   "tw_ventas_adenda.frx":114E
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   41
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   120
         Width           =   1245
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "tw_ventas_adenda.frx":1910
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   37
         ToolTipText     =   "Aprueba el Registro Elegido"
         Top             =   120
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   12000
         Picture         =   "tw_ventas_adenda.frx":2146
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   40
         ToolTipText     =   "Busca Registros "
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnDesAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   3840
         Picture         =   "tw_ventas_adenda.frx":28FB
         ScaleHeight     =   735
         ScaleWidth      =   1395
         TabIndex        =   39
         ToolTipText     =   "Cambiar Contrato a Provisional o Viceversa"
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5280
         Picture         =   "tw_ventas_adenda.frx":32F2
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   38
         ToolTipText     =   "Imprimir el Listado de los Registros"
         Top             =   120
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "tw_ventas_adenda.frx":3BBF
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   35
         ToolTipText     =   "Modifica datos del Contrato elegido"
         Top             =   120
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "tw_ventas_adenda.frx":44D4
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   34
         ToolTipText     =   "Nueva Adenda"
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADENDA DEL CONTRATO"
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
         Left            =   8130
         TabIndex        =   25
         Top             =   300
         Width           =   3900
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   120
      ScaleHeight     =   960
      ScaleWidth      =   16320
      TabIndex        =   20
      Top             =   120
      Width           =   16380
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   1560
         Picture         =   "tw_ventas_adenda.frx":4C93
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H80000015&
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "tw_ventas_adenda.frx":5469
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADENDA DEL CONTRATO"
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
         Left            =   8130
         TabIndex        =   23
         Top             =   300
         Width           =   3900
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00C00000&
      Height          =   5415
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   9015
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "venta_codigo_adenda"
            Caption         =   "IdAdenda"
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
            DataField       =   "motivo_codigo"
            Caption         =   "Cód.Motivo"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "monto_total_bs"
            Caption         =   "Total.Bs."
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
            DataField       =   "monto_total_dol"
            Caption         =   "Total.Dol."
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
            DataField       =   "fecha_inicio"
            Caption         =   "Fecha Inicio"
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
            DataField       =   "fecha_fin"
            Caption         =   "Fecha Finalización"
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
            DataField       =   "venta_codigo"
            Caption         =   "Nro. Venta"
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
            DataField       =   "cantidad_total"
            Caption         =   "Cantidad.Meses"
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
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3929.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   4920
         Width           =   8745
         _ExtentX        =   15425
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
         Caption         =   " <-- Inicio                                                                             Fin -->"
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
   Begin VB.Frame Fra_ABM 
      BackColor       =   &H00E0E0E0&
      Height          =   5415
      Left            =   9165
      TabIndex        =   11
      Top             =   1200
      Width           =   7335
      Begin VB.ComboBox TxtMoneda 
         DataField       =   "tipo_moneda"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "tw_ventas_adenda.frx":5D55
         Left            =   3120
         List            =   "tw_ventas_adenda.frx":5D5F
         TabIndex        =   46
         Text            =   "BOB"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txt_TipoCambio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "tipo_cambio"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   44
         Text            =   "0"
         Top             =   3435
         Width           =   1455
      End
      Begin VB.TextBox Txt_montoDol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_total_dol"
         DataSource      =   "Ado_datos"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   42
         Text            =   "0"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox Txt_tiempo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad_total"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0"
         Top             =   4155
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "descripcion"
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
         Height          =   885
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2040
         Width           =   6645
      End
      Begin VB.TextBox Txt_MontoBs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "monto_total_bs"
         DataSource      =   "Ado_datos"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   0
         Text            =   "0"
         Top             =   4440
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_ventas_adenda.frx":5D6D
         DataField       =   "motivo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "motivo_codigo"
         BoundColumn     =   "motivo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_ventas_adenda.frx":5D86
         DataField       =   "motivo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   1305
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "motivo_denominacion"
         BoundColumn     =   "motivo_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   1680
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5640
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfechaIni 
         DataField       =   "fecha_inicio"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   29
         Top             =   3435
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   119865345
         CurrentDate     =   44197
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker DTPfechaFin 
         DataField       =   "fecha_fin"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Top             =   3435
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   119865345
         CurrentDate     =   44197
         MinDate         =   36526
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Elija el Tipo de Moneda :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   45
         Top             =   3960
         Width           =   2235
      End
      Begin VB.Label lblTipoCambio 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5520
         TabIndex        =   43
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblMontoDol 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe (Incremento/Reducción) de Adenda en Dolares."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   33
         Top             =   4920
         Width           =   4995
      End
      Begin VB.Label lblMonto 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe (Incremento/Reducción) de la Adenda en Bs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   28
         Top             =   4440
         Width           =   4725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin Adenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3120
         TabIndex        =   27
         Top             =   3120
         Width           =   1650
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lbl_enlace2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio Adenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   3120
         Width           =   1845
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "venta_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label lbl_enlace1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones sobre el Motivo del Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   1725
         Width           =   3885
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo del Cambio de Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1005
         Width           =   2760
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REG"
         DataField       =   "estado_codigo"
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   5580
         TabIndex        =   12
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   16605
      TabIndex        =   5
      Top             =   6795
      Width           =   16605
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   10
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   2400
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   6720
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2880
      Top             =   6720
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
      Caption         =   "Ado_datos2"
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
Attribute VB_Name = "tw_ventas_adenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
         'db.Execute "update ao_ventas_cabecera set venta_monto_origen_bs = venta_monto_total_bs Where venta_codigo = " & NumComp & "  "
         'db.Execute "update ao_ventas_cabecera set venta_monto_adenda_bs = " & CDbl(Txt_MontoBs.Text) & " Where venta_codigo = " & NumComp & "   "
         db.Execute "update ao_ventas_cabecera set venta_monto_origen_bs = venta_monto_total_bs Where venta_codigo = " & NumComp & "  AND venta_monto_origen_bs = '0' "
         db.Execute "update ao_ventas_cabecera set venta_monto_adenda_bs = venta_monto_adenda_bs + " & CDbl(Txt_MontoBs.Text) & " Where venta_codigo = " & NumComp & "   "
         'db.Execute "update ao_ventas_cabecera set venta_monto_adenda_bs = venta_monto_adenda_bs + " & CDbl(Txt_MontoBs.Text) & " Where venta_codigo = " & NumComp & "   "
         db.Execute "update ao_ventas_cabecera set venta_monto_total_bs= venta_monto_origen_bs + venta_monto_adenda_bs Where venta_codigo = " & NumComp & "   "
         
         'db.Execute "update ao_ventas_cabecera set venta_monto_origen_dol = venta_monto_total_dol Where venta_codigo = " & NumComp & "   "
         'db.Execute "update ao_ventas_cabecera set venta_monto_adenda_dol = " & CDbl(Txt_montoDol.Text) & " Where venta_codigo = " & NumComp & "   "
         db.Execute "update ao_ventas_cabecera set venta_monto_origen_dol = venta_monto_total_dol Where venta_codigo = " & NumComp & " AND venta_monto_origen_dol = '0'  "
         db.Execute "update ao_ventas_cabecera set venta_monto_adenda_dol = venta_monto_adenda_dol + " & CDbl(Txt_montoDol.Text) & " Where venta_codigo = " & NumComp & "   "
         db.Execute "update ao_ventas_cabecera set venta_monto_total_dol= venta_monto_origen_dol + venta_monto_adenda_dol Where venta_codigo = " & NumComp & "   "
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ERR) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_ABM.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        txt_codigo.Enabled = True
        dtc_desc1.Enabled = True
    End If
End Sub

Private Sub BtnEliminar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If ExisteReg2(Ado_datos.Recordset!calle_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ERR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Está Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    txt_codigo.Caption = NumComp
    If VAR_SW = "ADD" Then
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "Select * from ao_ventas_adenda where venta_codigo = " & Val(txt_codigo.Caption) & "  and motivo_codigo = " & Val(dtc_codigo1.Text) & "   ", db, adOpenStatic
'        If rs_aux2.RecordCount > 0 Then
'            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'            MsgBox "El registro ya Existe, Vuelva a intentar ...", vbExclamation, "Validación de Registro"
'            Exit Sub
'        End If
        'rs_datos!ges_gestion = glGestion
        rs_datos!estado_codigo = "REG"  ' no cambia
        rs_datos!venta_codigo = txt_codigo.Caption    'Codigo del padre
        'rs_datos!motivo_codigo = dtc_codigo1.Text   'Codigo del padre 1
        'Guarda en el Padre, en el campo ctrl de correlativos para codigos que se generan
        'db.Execute "Update gc_zonas Set correl = " & var_cod & " Where zona_codigo= '" & dtc_codigo1.Text & "' "
     End If
     'venta_codigo_adenda, venta_codigo, venta_tipo, monto_total_bs, monto_total_dol, fecha_inicio, fecha_fin, cantidad_total, motivo_codigo, descripcion, estado_codigo, usr_codigo, fecha_registro, hora_registro
     rs_datos!motivo_codigo = dtc_codigo1.Text   'Codigo del padre 1
     rs_datos!descripcion = txt_obs.Text
     
     If Txt_tiempo.Text = "" Then
        rs_datos!cantidad_total = "1"
     'Else
     '   rs_datos!cantidad_total = CDbl(Txt_tiempo.Text) / 30
     End If
     If TxtMoneda.Text = "" Then
        rs_datos!tipo_moneda = "BOB"
     Else
        rs_datos!tipo_moneda = RTrim(TxtMoneda.Text)
     End If
     Select Case dtc_codigo1.Text
        Case 19
            rs_datos!monto_total_bs = IIf(Txt_MontoBs.Text = "", "0", CDbl(Txt_MontoBs.Text))
            rs_datos!monto_total_dol = IIf(Txt_montoDol.Text = "", "0", CDbl(Txt_montoDol.Text))
            'If rs_datos!monto_total_bs = 0 Then
            '   rs_datos!monto_total_dol = 0
            'Else
            '   rs_datos!monto_total_dol = rs_datos!monto_total_bs / GlTipoCambioOficial
            'End If
        Case 22
            rs_datos!monto_total_bs = IIf(Txt_MontoBs.Text = "", "0", CDbl(Txt_MontoBs.Text) * (-1))
            rs_datos!monto_total_dol = IIf(Txt_montoDol.Text = "", "0", CDbl(Txt_montoDol.Text) * (-1))
            
'            If rs_datos!monto_total_bs = 0 Then
'               rs_datos!monto_total_dol = 0
'            Else
'               rs_datos!monto_total_dol = rs_datos!monto_total_bs / GlTipoCambioOficial
'            End If
        Case 18, 21
            rs_datos!fecha_inicio = IIf(IsNull(DTPfechaIni.Value), "01/01/1900", DTPfechaIni.Value)
            rs_datos!fecha_fin = IIf(IsNull(DTPfechaFin.Value), "01/01/1900", DTPfechaFin.Value)
            rs_datos!monto_total_bs = 0
            rs_datos!monto_total_dol = 0
        Case 20
            rs_datos!monto_total_bs = IIf(Txt_MontoBs.Text = "", "0", CDbl(Txt_MontoBs.Text))
            rs_datos!monto_total_dol = IIf(Txt_montoDol.Text = "", "0", CDbl(Txt_montoDol.Text))
            
'            rs_datos!monto_total_bs = IIf(Txt_MontoBs.Text = "", "0", CDbl(Txt_MontoBs.Text))
'            If rs_datos!monto_total_bs = 0 Then
'               rs_datos!monto_total_dol = 0
'            Else
'               rs_datos!monto_total_dol = rs_datos!monto_total_bs / GlTipoCambioOficial
'            End If
            rs_datos!fecha_inicio = IIf(IsNull(DTPfechaIni.Value), "01/01/1900", DTPfechaIni.Value)
            rs_datos!fecha_fin = IIf(IsNull(DTPfechaFin.Value), "01/01/1900", DTPfechaFin.Value)
        Case 23
            rs_datos!monto_total_bs = IIf(Txt_MontoBs.Text = "", "0", CDbl(Txt_MontoBs.Text) * (-1))
            rs_datos!monto_total_dol = IIf(Txt_montoDol.Text = "", "0", CDbl(Txt_montoDol.Text) * (-1))

'            If rs_datos!monto_total_bs = 0 Then
'               rs_datos!monto_total_dol = 0
'            Else
'               rs_datos!monto_total_dol = rs_datos!monto_total_bs / GlTipoCambioOficial
'            End If
            rs_datos!fecha_inicio = IIf(IsNull(DTPfechaIni.Value), "01/01/1900", DTPfechaIni.Value)
            rs_datos!fecha_fin = IIf(IsNull(DTPfechaFin.Value), "01/01/1900", DTPfechaFin.Value)
        Case Else
            rs_datos!monto_total_bs = 0
            rs_datos!monto_total_dol = 0
            rs_datos!fecha_inicio = IIf(IsNull(DTPfechaIni.Value), "01/01/1900", DTPfechaIni.Value)
            rs_datos!fecha_fin = IIf(IsNull(DTPfechaFin.Value), "01/01/1900", DTPfechaFin.Value)
     End Select

     'venta_codigo_adenda, venta_codigo,monto_total_bs, monto_total_dol,fecha_inicio, fecha_fin, cantidad_total,motivo_codigo,
     ' venta_tipo, estado_codigo, usr_codigo, fecha_registro, hora_registro
     'If dtc_codigo1.Text = "6" TheN
     '   var_cod = Round((Val(Txt_MontoBs.Text) / 30), 0)
     '   db.Execute "update ao_ventas_detalle set bien_cantidad_por_empaque = " & var_cod & " Where venta_codigo = " & txt_codigo.Caption & "   "
     'End If
     
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
     marca1 = Ado_datos.Recordset.Bookmark
     Call ABRIR_TABLA
     'rs_datos.MoveLast
     mbDataChanged = False
     Ado_datos.Recordset.Move marca1 - 1
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      txt_codigo.Enabled = True
      dtc_desc1.Enabled = True
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
    'habilitar codigo cuando se transcribe
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
    
    Select Case dtc_codigo1.Text
        Case 19, 22
            If Txt_MontoBs.Text = "" Then
              MsgBox "Debe registrar el " + lblMonto.Caption, vbCritical + vbExclamation, "Validación de datos"
              VAR_VAL = "ERR"
              Exit Sub
            End If
        Case 18, 21
            If Val(Txt_tiempo) < 1 Then
              MsgBox "La Fecha de Inicio NO puede ser MAYOR ni IGUAL a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
              DTPfechaFin.SetFocus
              VAR_VAL = "ERR"
              Exit Sub
            End If
        Case 20, 23
            If Txt_MontoBs.Text = "" Then
              MsgBox "Debe registrar el " + lblMonto.Caption, vbCritical + vbExclamation, "Validación de datos"
              VAR_VAL = "ERR"
              Exit Sub
            End If
            If Val(Txt_tiempo) < 1 Then
              MsgBox "La Fecha de Inicio NO puede ser MAYOR ni IGUAL a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
              DTPfechaFin.SetFocus
              VAR_VAL = "ERR"
              Exit Sub
            End If
        Case Else
            If Txt_MontoBs.Text = "" Then
              MsgBox "Debe registrar el " + lblMonto.Caption, vbCritical + vbExclamation, "Validación de datos"
              VAR_VAL = "ERR"
              Exit Sub
            End If
            If Val(Txt_tiempo) < 1 Then
              MsgBox "La Fecha de Inicio NO puede ser MAYOR ni IGUAL a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
              DTPfechaFin.SetFocus
              VAR_VAL = "ERR"
              Exit Sub
            End If
    End Select
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_orden_adenda.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
        '  cr01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
        '  cr01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "

        CR01.StoredProcParam(0) = Ado_datos.Recordset!venta_codigo_adenda      'ges_gestion
        'cr01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        'cr01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
        CR01.WindowState = crptMaximized
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

'  Dim iResult As Integer
'  CR01.WindowShowPrintSetupBtn = True
'  CR01.WindowShowRefreshBtn = True
'  CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\gr_direccion_general.rpt"
'  iResult = CR01.PrintReport
'  If iResult <> 0 Then
'      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
'  End If
'  CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo EditErr
  If rs_datos!estado_codigo = "REG" Then
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
    'txt_codigo.Enabled = False
    'dtc_desc1.Enabled = False
  Else
      MsgBox "No se puede MODIFICAR un registro Aprobado(APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
  End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
'  If glPersOtro = "O" Then
'    frmmo_pacientes.Dtc_ocupac = rs_datos!ocup_codigo
'    frmmo_pacientes.Dtc_OcupacDes = rs_datos!ocup_descripcion
'  End If
'  glPersOtro = "N"
  Unload Me
End Sub

Private Sub DtcUE_Click(Area As Integer)
    DtcUE_Des.BoundText = DtcUE.BoundText
End Sub

Private Sub DtcUE_Des_Click(Area As Integer)
    DtcUE.BoundText = DtcUE_Des.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub dtc_desc1_LostFocus()
  Select Case dtc_codigo1.Text
    Case 17
        DTPfechaIni.Visible = False
        DTPfechaFin.Visible = False
        Txt_MontoBs.Visible = False
        
        lblTipoCambio.Visible = False
        txt_TipoCambio.Visible = False
        lblMontoDol.Visible = False
        Txt_montoDol.Visible = False
    Case 18
        DTPfechaIni.Visible = True
        DTPfechaFin.Visible = True
        Txt_MontoBs.Visible = False
        
        lblTipoCambio.Visible = False
        txt_TipoCambio.Visible = False
        lblMontoDol.Visible = False
        Txt_montoDol.Visible = False
    Case 19
        DTPfechaIni.Visible = False
        DTPfechaFin.Visible = False
        Txt_MontoBs.Visible = True
        
        lblTipoCambio.Visible = True
        txt_TipoCambio.Visible = True
        lblMontoDol.Visible = True
        Txt_montoDol.Visible = True
    Case 20
        DTPfechaIni.Visible = True
        DTPfechaFin.Visible = True
        Txt_MontoBs.Visible = True
        
        lblTipoCambio.Visible = True
        txt_TipoCambio.Visible = True
        lblMontoDol.Visible = True
        Txt_montoDol.Visible = True
    Case 21
        DTPfechaIni.Visible = True
        DTPfechaFin.Visible = True
        Txt_MontoBs.Visible = False
    
        lblTipoCambio.Visible = False
        txt_TipoCambio.Visible = False
        lblMontoDol.Visible = False
        Txt_montoDol.Visible = False
    
    Case 22
        DTPfechaIni.Visible = False
        DTPfechaFin.Visible = False
        Txt_MontoBs.Visible = True
        
        lblTipoCambio.Visible = True
        txt_TipoCambio.Visible = True
        lblMontoDol.Visible = True
        Txt_montoDol.Visible = True
    Case 23
        DTPfechaIni.Visible = True
        DTPfechaFin.Visible = True
        Txt_MontoBs.Visible = True
        
        lblTipoCambio.Visible = True
        txt_TipoCambio.Visible = True
        lblMontoDol.Visible = True
        Txt_montoDol.Visible = True
    Case Else
        MsgBox "El Motivo NO corresponde para este Proceso, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        dtc_desc1.SetFocus
  End Select
  dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub DTPFechaFin_LostFocus()
    'Me.Print Format(DateDiff("y", Fecha_Inicial, Fecha_Final), Formato) & " dias"
    'Txt_MontoBs = Format(DateDiff("y", DTPfechaIni, DTPfechaFin), Formato)
    Txt_tiempo = DateDiff("y", DTPfechaIni, DTPfechaFin)
    If Val(Txt_tiempo) < 0 Then
        MsgBox "La Fecha de Inicio NO puede ser MAYOR a la Fecha de Finalización, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        DTPfechaFin.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
    txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
    
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from ao_ventas_adenda WHERE venta_codigo = " & NumComp & "  "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
  If Ado_datos.Recordset.RecordCount = 0 Then
        Txt_MontoBs.Text = "0"
  End If
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from rc_motivo_proceso where motivo_tipo = '0' order by motivo_denominacion ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

'Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
      Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
'      If Label2.Caption = "" Then
'        Label2.Caption = "Tiempo Estimado en Días Calendario"
'      Else
'        Label2.Caption = "Tiempo Estimado en Meses Calendario"
'      End If
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aquí se coloca el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub BtnAñadir_Click()
    If glusuario = "CCRUZ" Or glusuario = "LNAVA" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo AddErr
    Call ABRIR_TABLA
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
    Select Case Cod_Comp
        Case 3, 4, 6, 9
            lblTipoCambio.Visible = True
            txt_TipoCambio.Visible = True
            lblMontoDol.Visible = True
            Txt_montoDol.Visible = True
        Case 7, 8, 10
            lblTipoCambio.Visible = False
            txt_TipoCambio.Visible = False
            lblMontoDol.Visible = False
            Txt_montoDol.Visible = False
        Case Else
            lblTipoCambio.Visible = False
            txt_TipoCambio.Visible = False
            lblMontoDol.Visible = False
            Txt_montoDol.Visible = False
    End Select
'    txt_codigo.Enabled = False
    'Txt_MontoBs.SetFocus
    Txt_MontoBs.Text = "0"
    dtc_desc1.SetFocus
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto sólo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM gc_beneficiario WHERE estado_codigo = 'APR' and calle_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Function ExisteReg2(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM gc_edificaciones WHERE estado_codigo = 'APR' and calle_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg2 = rs!Cuantos > 0
End Function

Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub Txt_MontoBs_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub Txt_MontoBs_LostFocus()
    If txt_TipoCambio.Text = "" Or txt_TipoCambio.Text = "0" Then
        txt_TipoCambio.Text = GlTipoCambioOficial
    End If
    If Txt_MontoBs.Text = "" Or Txt_MontoBs.Text = "0" Then
       Txt_montoDol.Text = 0
    Else
       Txt_montoDol.Text = Round(CDbl(Txt_MontoBs.Text) / CDbl(txt_TipoCambio.Text), 2)
    End If
End Sub

Private Sub Txt_montoDol_LostFocus()
    If txt_TipoCambio.Text = "" Or txt_TipoCambio.Text = "0" Then
        txt_TipoCambio.Text = GlTipoCambioOficial
    End If
    If Txt_montoDol.Text = "" Or Txt_montoDol.Text = "0" Then
       Txt_MontoBs.Text = 0
    Else
       Txt_MontoBs.Text = CDbl(Txt_montoDol.Text) * CDbl(txt_TipoCambio.Text)
    End If
End Sub

Private Sub TxtMoneda_LostFocus()
    If TxtMoneda.Text = "USD" Then
        Txt_montoDol.Enabled = True
        Txt_MontoBs.Enabled = False
    Else
        Txt_montoDol.Enabled = False
        Txt_MontoBs.Enabled = True
    End If
End Sub

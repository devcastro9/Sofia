VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form mw_solicitud_cotiza_venta 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Módulo Comercial - Cotización de Equipos"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "mw_solicitud_cotiza_venta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Fra_datos2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000040&
      Height          =   2580
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   15465
      Begin VB.PictureBox BtnSalir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7320
         Picture         =   "mw_solicitud_cotiza_venta.frx":0A02
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   77
         ToolTipText     =   "Buscar Registros"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7260
         TabIndex        =   42
         Top             =   735
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   14520
         TabIndex        =   41
         Top             =   435
         Width           =   270
      End
      Begin MSDataListLib.DataCombo txt_aux3 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":11C4
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   13680
         TabIndex        =   40
         Top             =   420
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_tipo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo txt_codigo3 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":11DE
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   12360
         TabIndex        =   39
         Top             =   420
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   5880
         TabIndex        =   37
         Top             =   435
         Width           =   270
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "estado_codigo"
         DataSource      =   "Ado_datos0"
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Txt_campo11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "unidad_codigo_ant"
         DataSource      =   "Ado_datos0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "cotiza_fecha"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   13140
         TabIndex        =   13
         Top             =   1080
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12632256
         CalendarTitleBackColor=   -2147483638
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   117506051
         CurrentDate     =   44235
         MaxDate         =   55153
         MinDate         =   32874
      End
      Begin MSDataListLib.DataCombo Txt_campo12 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":11F8
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   1560
         TabIndex        =   35
         Top             =   420
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo Txt_campo1 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":1212
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   5040
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo txt_desc3 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":122D
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos0"
         Height          =   315
         Left            =   6480
         TabIndex        =   38
         Top             =   420
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label dtc_codigo11 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "h_nro_total_equipos"
         DataSource      =   "Ado_datos0"
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
         Left            =   1200
         TabIndex        =   66
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label dtc_desc10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   4800
         TabIndex        =   65
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label dtc_desc11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "h_nro_total_equipos"
         DataSource      =   "Ado_datos0"
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
         Height          =   300
         Left            =   6120
         TabIndex        =   64
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label dtc_desc14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "h_capacidad_trafico"
         DataSource      =   "Ado_datos0"
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
         Height          =   300
         Left            =   3360
         TabIndex        =   63
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label dtc_desc13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "h_intervalo_trafico"
         DataSource      =   "Ado_datos0"
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
         Height          =   300
         Left            =   1800
         TabIndex        =   62
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label dtc_desc12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "h_partidas_por_hora"
         DataSource      =   "Ado_datos0"
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
         Height          =   300
         Left            =   240
         TabIndex        =   61
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "De                      ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7230
         TabIndex        =   22
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label dtc_desc15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   7560
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label dtc_desc16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   9135
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label dtc_desc17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   14280
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   $"mw_solicitud_cotiza_venta.frx":1247
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
         Left            =   240
         TabIndex        =   16
         Top             =   825
         Width           =   14430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No.Trámite         Unidad Ejecutora"
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
         Left            =   240
         TabIndex        =   15
         Top             =   180
         Width           =   2970
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"mw_solicitud_cotiza_venta.frx":1303
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
         Index           =   0
         Left            =   6480
         TabIndex        =   12
         Top             =   180
         Width           =   8295
      End
      Begin VB.Label txt_codigo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Height          =   300
         Left            =   9840
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
         DataSource      =   "Ado_datos0"
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
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE COSTOS POR EQUIPO"
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   2040
      TabIndex        =   5
      Top             =   6495
      Width           =   13785
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":13A4
         Height          =   2895
         Left            =   195
         TabIndex        =   6
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "unidad_codigo"
            Caption         =   "Codigo Unidad"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "No. Negocia"
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
            DataField       =   "edif_codigo"
            Caption         =   "Codigo Edificio"
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
            DataField       =   "cotiza_codigo"
            Caption         =   "Nro. Cotización"
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
            DataField       =   "codigo_costo"
            Caption         =   "Codigo.Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "costo_porcentaje"
            Caption         =   "% Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "costo_monto"
            Caption         =   "Costo Unitario Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "costo_monto_usd"
            Caption         =   "Costo Unitario ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "costo_monto2"
            Caption         =   "Costo Grupo Bs"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "costo_monto_usd2"
            Caption         =   "Costo Grupo ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "costo_monto3"
            Caption         =   "Costo Grupo Bs"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "costo_monto_usd3"
            Caption         =   "Costo Grupo ME"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "costo_observaciones"
            Caption         =   "Detalle"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               DividerStyle    =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               DividerStyle    =   1
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               ColumnWidth     =   6480
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega0 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REGISTRO "
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Width           =   15690
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H80000014&
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   5280
         TabIndex        =   33
         Top             =   2355
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H80000014&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   9000
         TabIndex        =   32
         Top             =   2355
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos0 
         Height          =   330
         Left            =   120
         Top             =   2280
         Width           =   15405
         _ExtentX        =   27173
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
         BackColor       =   -2147483628
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
         Caption         =   " "
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
      Begin MSDataGridLib.DataGrid dg_datos0 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":13BF
         Height          =   2070
         Left            =   135
         TabIndex        =   34
         Top             =   195
         Width           =   15405
         _ExtentX        =   27173
         _ExtentY        =   3651
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483628
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "No.Tramite"
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
            DataField       =   "unidad_codigo"
            Caption         =   "U.Ejecutora"
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
            DataField       =   "cotiza_codigo"
            Caption         =   "Cotizacion"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Negociación"
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
            DataField       =   "cotiza_fecha"
            Caption         =   "Fecha.Cotiza"
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
            DataField       =   "usr_codigo"
            Caption         =   "Usuario"
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
            DataField       =   "pais_codigo"
            Caption         =   "Pais"
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
            DataField       =   "tipo_eqp"
            Caption         =   "Tipo.Equipo"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo"
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
            DataField       =   "cotiza_nro_montador"
            Caption         =   "#Montadores"
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
            DataField       =   "dimension_fosa_fondo"
            Caption         =   "Fosa.Fondo"
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
            DataField       =   "dimension_fosa_frente"
            Caption         =   "Fosa.Frente"
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
            DataField       =   "dimension_fosa_m"
            Caption         =   "Espacio.Dintel"
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
            DataField       =   "cotiza_energia"
            Caption         =   "Energía"
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
            DataField       =   "cotiza_luz"
            Caption         =   "Luz"
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
            DataField       =   "estado_codigo_cot"
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
         BeginProperty Column17 
            DataField       =   "bien_cotiza_num_accesos"
            Caption         =   "#Accesos"
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
         BeginProperty Column18 
            DataField       =   "cotiza_fecha"
            Caption         =   "Fecha.Cotización"
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column16 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1365.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      Height          =   740
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   15645
      TabIndex        =   0
      Top             =   45
      Width           =   15705
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3600
         Picture         =   "mw_solicitud_cotiza_venta.frx":13D8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   71
         ToolTipText     =   "Buscar Registros"
         Top             =   40
         Width           =   1215
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   14040
         Picture         =   "mw_solicitud_cotiza_venta.frx":1F96
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   70
         ToolTipText     =   "Buscar Registros"
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":2758
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   69
         ToolTipText     =   "Buscar Registros"
         Top             =   40
         Width           =   1455
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1800
         Picture         =   "mw_solicitud_cotiza_venta.frx":3321
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   68
         ToolTipText     =   "Buscar Registros"
         Top             =   40
         Width           =   1215
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   600
         Left            =   5400
         Picture         =   "mw_solicitud_cotiza_venta.frx":3AD6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COTIZA"
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
         Left            =   8595
         TabIndex        =   2
         Top             =   180
         Width           =   1155
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2805
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   4948
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Proveedor de AMERICA"
      TabPicture(0)   =   "mw_solicitud_cotiza_venta.frx":3CE0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraNavega"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Proveedor de ASIA"
      TabPicture(1)   =   "mw_solicitud_cotiza_venta.frx":3CFC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraNavegaA"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Proveedor de EUROPA"
      TabPicture(2)   =   "mw_solicitud_cotiza_venta.frx":3D18
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FraNavegaE"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame FraNavegaE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "REGISTRO DE DATOS PARA LA COTIZACION"
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   -74940
         TabIndex        =   26
         Top             =   345
         Width           =   15615
         Begin MSDataGridLib.DataGrid dg_datosE 
            Bindings        =   "mw_solicitud_cotiza_venta.frx":3D34
            Height          =   1320
            Left            =   120
            TabIndex        =   57
            Top             =   1035
            Width           =   15360
            _ExtentX        =   27093
            _ExtentY        =   2328
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
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
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "pais_codigo"
               Caption         =   "País"
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
               DataField       =   "cotiza_codigo"
               Caption         =   "Cotización"
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
               DataField       =   "cotiza_precio_fob_dol"
               Caption         =   "Precio.FOB_Usd"
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
               DataField       =   "cotiza_precio_seg_dol"
               Caption         =   "Seguro.Transp.Usd"
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
               DataField       =   "cotiza_fob_seg_dol"
               Caption         =   "FOB+Seguro.Usd"
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
               DataField       =   "cotiza_precio_flete_dol"
               Caption         =   "Flete.Front.Usd"
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
               DataField       =   "cotiza_precio_cif_dol"
               Caption         =   "Precio.CIF.Usd"
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
               DataField       =   "cotiza_precio_total_dol"
               Caption         =   "Sub.Total.Usd"
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
               DataField       =   "cotiza_precio_total_dol_cli"
               Caption         =   "Importacion.Directa.Usd"
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
               DataField       =   "cotiza_precio_total_dol_cge"
               Caption         =   "Facturacion.Local.Usd"
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
               DataField       =   "cotiza_gasto_local_dol"
               Caption         =   "Tot.Gasto.Local.Usd"
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
            BeginProperty Column12 
               DataField       =   "cotiza_precio_dcto_dol"
               Caption         =   "Descuento.Usd"
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
               DataField       =   "cotiza_saldo_tac_billing_dol"
               Caption         =   "Saldo.Tac.Billing.Usd"
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
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   884.976
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1500.095
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1395.213
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1230.236
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   629.858
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   840.189
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox fraOpciones1E 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   750
            Left            =   120
            ScaleHeight     =   690
            ScaleWidth      =   15375
            TabIndex        =   27
            Top             =   240
            Width           =   15435
            Begin VB.CommandButton BtnImprimir2E 
               BackColor       =   &H00FFFFFF&
               Caption         =   "R-224"
               Height          =   650
               Left            =   3850
               Picture         =   "mw_solicitud_cotiza_venta.frx":3D4D
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnImprimirE 
               BackColor       =   &H00FFFFFF&
               Caption         =   "R-222"
               Height          =   650
               Left            =   2610
               Picture         =   "mw_solicitud_cotiza_venta.frx":430A
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnAprobarE 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Verificar"
               Height          =   650
               Left            =   1365
               Picture         =   "mw_solicitud_cotiza_venta.frx":48C7
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Ok, envía datos para Contrato de Venta"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnModificar1E 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Hoja.de.Costos"
               Height          =   650
               Left            =   120
               Picture         =   "mw_solicitud_cotiza_venta.frx":4AD1
               Style           =   1  'Graphical
               TabIndex        =   46
               ToolTipText     =   "Registra Hoja de Costos"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnModificarE 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Datos.Iniciales"
               Height          =   680
               Left            =   4560
               Picture         =   "mw_solicitud_cotiza_venta.frx":50B1
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Registra Datos Iniciales para Cotización"
               Top             =   30
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HOJA DE COSTOS - EUROPA"
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
               Left            =   8100
               TabIndex        =   60
               Top             =   120
               Width           =   4425
            End
         End
         Begin MSAdodcLib.Adodc Ado_datosE 
            Height          =   330
            Left            =   1800
            Top             =   1920
            Visible         =   0   'False
            Width           =   5265
            _ExtentX        =   9287
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
      End
      Begin VB.Frame FraNavegaA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "REGISTRO DE DATOS PARA LA COTIZACION"
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   -74940
         TabIndex        =   23
         Top             =   345
         Width           =   15615
         Begin MSDataGridLib.DataGrid dg_datosA 
            Bindings        =   "mw_solicitud_cotiza_venta.frx":5691
            Height          =   1320
            Left            =   120
            TabIndex        =   54
            Top             =   1035
            Width           =   15360
            _ExtentX        =   27093
            _ExtentY        =   2328
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "pais_codigo"
               Caption         =   "País"
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
               DataField       =   "cotiza_codigo"
               Caption         =   "Cotización"
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
               DataField       =   "cotiza_precio_fob_dol"
               Caption         =   "Precio.FOB_Usd"
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
               DataField       =   "cotiza_precio_seg_dol"
               Caption         =   "Seguro.Transp.Usd"
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
               DataField       =   "cotiza_fob_seg_dol"
               Caption         =   "FOB+Seguro.Usd"
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
               DataField       =   "cotiza_precio_flete_dol"
               Caption         =   "Flete.Front.Usd"
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
               DataField       =   "cotiza_precio_cif_dol"
               Caption         =   "Precio.CIF.Usd"
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
               DataField       =   "cotiza_precio_total_dol"
               Caption         =   "Sub.Total.Usd"
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
               DataField       =   "cotiza_precio_total_dol_cli"
               Caption         =   "Importacion.Directa.Usd"
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
               DataField       =   "cotiza_precio_total_dol_cge"
               Caption         =   "Facturacion.Local.Usd"
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
               DataField       =   "cotiza_gasto_local_dol"
               Caption         =   "Tot.Gasto.Local.Usd"
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
            BeginProperty Column12 
               DataField       =   "cotiza_precio_dcto_dol"
               Caption         =   "Descuento.Usd"
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
               DataField       =   "cotiza_saldo_tac_billing_dol"
               Caption         =   "Saldo.Tac.Billing.Usd"
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
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   884.976
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1319.811
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1409.953
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1289.764
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1574.929
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   629.858
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   1035.213
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox fraOpciones1A 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   750
            Left            =   120
            ScaleHeight     =   690
            ScaleWidth      =   15375
            TabIndex        =   24
            Top             =   240
            Width           =   15435
            Begin VB.CommandButton BtnImprimir2A 
               BackColor       =   &H00C0FFFF&
               Caption         =   "R-224"
               Height          =   650
               Left            =   3850
               Picture         =   "mw_solicitud_cotiza_venta.frx":56AB
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnImprimirA 
               BackColor       =   &H00C0FFFF&
               Caption         =   "R-222"
               Height          =   650
               Left            =   2610
               Picture         =   "mw_solicitud_cotiza_venta.frx":5C68
               Style           =   1  'Graphical
               TabIndex        =   55
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnAprobarA 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Verificar"
               Height          =   650
               Left            =   1365
               Picture         =   "mw_solicitud_cotiza_venta.frx":6225
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Ok, envía datos para Contrato de Venta"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnModificar1A 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Hoja.de.Costos"
               Height          =   650
               Left            =   120
               Picture         =   "mw_solicitud_cotiza_venta.frx":642F
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Registra Hoja de Costos"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnModificarA 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Datos.Iniciales"
               Height          =   680
               Left            =   5160
               Picture         =   "mw_solicitud_cotiza_venta.frx":6871
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Registra Datos Iniciales para Cotización"
               Top             =   30
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HOJA DE COSTOS - ASIA"
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
               Left            =   8520
               TabIndex        =   53
               Top             =   120
               Width           =   3825
            End
         End
         Begin MSAdodcLib.Adodc Ado_datosA 
            Height          =   330
            Left            =   120
            Top             =   1320
            Visible         =   0   'False
            Width           =   5265
            _ExtentX        =   9287
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
      End
      Begin VB.Frame FraNavega 
         BackColor       =   &H00C0C0C0&
         Caption         =   "HOJA DE COSTOS"
         ForeColor       =   &H00C00000&
         Height          =   2415
         Left            =   60
         TabIndex        =   11
         Top             =   345
         Width           =   15615
         Begin MSDataGridLib.DataGrid dg_datos 
            Bindings        =   "mw_solicitud_cotiza_venta.frx":6E51
            Height          =   1320
            Left            =   120
            TabIndex        =   49
            Top             =   1035
            Width           =   15360
            _ExtentX        =   27093
            _ExtentY        =   2328
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12640511
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
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "pais_codigo"
               Caption         =   "País"
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
               DataField       =   "cotiza_codigo"
               Caption         =   "Cotización"
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
               DataField       =   "cotiza_precio_fob_dol"
               Caption         =   "Precio.FOB_Usd"
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
               DataField       =   "cotiza_precio_seg_dol"
               Caption         =   "Seguro.Transport."
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
               DataField       =   "cotiza_precio_flete_dol"
               Caption         =   "Flete.Frontera"
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
               DataField       =   "cotiza_fob_seg_dol"
               Caption         =   "FOB+Seguro.Usd"
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
               DataField       =   "cotiza_precio_cif_dol"
               Caption         =   "Precio.CIF.Usd"
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
               DataField       =   "cotiza_gasto_local_dol"
               Caption         =   "Gasto.Local.Usd"
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
               DataField       =   "cotiza_precio_total_dol"
               Caption         =   "Sub.Total.Usd"
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
               DataField       =   "cotiza_precio_total_dol_cli"
               Caption         =   "Importacion.Directa"
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
               DataField       =   "cotiza_precio_total_dol_cge"
               Caption         =   "Facturacion.Local"
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
            BeginProperty Column12 
               DataField       =   "cotiza_precio_dcto_dol"
               Caption         =   "Descuento.Usd"
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
               DataField       =   "cotiza_saldo_tac_billing_dol"
               Caption         =   "Saldo.Tac.Billing.Usd"
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
                  ColumnWidth     =   569.764
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   1470.047
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   629.858
               EndProperty
               BeginProperty Column12 
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   840.189
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox fraOpciones1 
            BackColor       =   &H80000015&
            FillColor       =   &H00FFFFFF&
            Height          =   750
            Left            =   120
            ScaleHeight     =   690
            ScaleWidth      =   15300
            TabIndex        =   17
            Top             =   240
            Width           =   15360
            Begin VB.CommandButton BtnAprobar 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Verificar"
               Height          =   650
               Left            =   1365
               Picture         =   "mw_solicitud_cotiza_venta.frx":6E69
               Style           =   1  'Graphical
               TabIndex        =   67
               ToolTipText     =   "Ok, envía datos para Contrato de Venta"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnImprimir2 
               BackColor       =   &H00C0E0FF&
               Caption         =   "R-224"
               Height          =   650
               Left            =   3850
               Picture         =   "mw_solicitud_cotiza_venta.frx":7073
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnImprimir 
               BackColor       =   &H00C0E0FF&
               Caption         =   "R-222"
               Height          =   650
               Left            =   2610
               Picture         =   "mw_solicitud_cotiza_venta.frx":7630
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Imprime Formulario"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnModificar 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Datos Iniciales"
               Height          =   680
               Left            =   5160
               Picture         =   "mw_solicitud_cotiza_venta.frx":7BED
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Registra Datos Iniciales para Cotización"
               Top             =   30
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.CommandButton BtnModificar1 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Hoja de Costos"
               Height          =   650
               Left            =   120
               Picture         =   "mw_solicitud_cotiza_venta.frx":81CD
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Registra Hoja de Costos"
               Top             =   30
               Width           =   1245
            End
            Begin VB.CommandButton BtnAñadir 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Copia"
               Height          =   720
               Left            =   360
               Picture         =   "mw_solicitud_cotiza_venta.frx":860F
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Nuevo Registro"
               Top             =   60
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HOJA DE COSTOS - AMERICA"
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
               Left            =   7920
               TabIndex        =   52
               Top             =   120
               Width           =   4545
            End
         End
         Begin MSAdodcLib.Adodc Ado_datos 
            Height          =   330
            Left            =   120
            Top             =   1920
            Visible         =   0   'False
            Width           =   5265
            _ExtentX        =   9287
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
            BackColor       =   12640511
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
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      Height          =   3165
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   1755
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":8C33
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   76
         ToolTipText     =   "Graba los Cambios del Detalle de Costos"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox BtnAddDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":9421
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   75
         ToolTipText     =   "Crea un nuevo Item (Costo)..."
         Top             =   2160
         Width           =   1335
      End
      Begin VB.PictureBox BtnAnlDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":A10B
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   74
         ToolTipText     =   "Anula el Item elegido..."
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox BtnModDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":A857
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   73
         ToolTipText     =   "Modifica los Costos de los Items"
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   590
         Left            =   240
         Picture         =   "mw_solicitud_cotiza_venta.frx":B16C
         ScaleHeight     =   585
         ScaleWidth      =   1335
         TabIndex        =   72
         ToolTipText     =   "Borra los Items y los vuelve a cargar Todos..."
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   6480
      Top             =   9600
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
      Caption         =   "Ado_datos21"
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   2160
      Top             =   10320
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
      Caption         =   "Ado_datos3"
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
   Begin MSAdodcLib.Adodc Ado_datos41 
      Height          =   330
      Left            =   10800
      Top             =   9600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Ado_datos41"
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
   Begin MSAdodcLib.Adodc Ado_datos51 
      Height          =   330
      Left            =   12960
      Top             =   9600
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
      Caption         =   "Ado_datos51"
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
   Begin MSAdodcLib.Adodc Ado_datos61 
      Height          =   330
      Left            =   0
      Top             =   9960
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos61"
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
   Begin MSAdodcLib.Adodc Ado_datos31 
      Height          =   330
      Left            =   8640
      Top             =   9600
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
      Caption         =   "Ado_datos31"
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
   Begin Crystal.CrystalReport CR01 
      Left            =   4440
      Top             =   10320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   4320
      Top             =   9600
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos11"
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
   Begin MSAdodcLib.Adodc Ado_datos03 
      Height          =   330
      Left            =   0
      Top             =   10350
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
      Caption         =   "Ado_datos03"
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   2160
      Top             =   9960
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
      ConnectStringType=   3
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
      Caption         =   "Ado_detalle1"
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
      Left            =   4320
      Top             =   9960
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
      ConnectStringType=   3
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
   Begin MSAdodcLib.Adodc Ado_datos7 
      Height          =   330
      Left            =   6480
      Top             =   9960
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
      ConnectStringType=   3
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
      Left            =   8640
      Top             =   9960
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
      ConnectStringType=   3
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
      Left            =   10800
      Top             =   9960
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
      ConnectStringType=   3
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
   Begin VB.Frame FraDet1E 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE COSTOS POR EQUIPO"
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   2040
      TabIndex        =   29
      Top             =   6480
      Width           =   13785
      Begin MSDataGridLib.DataGrid dg_det1E 
         Bindings        =   "mw_solicitud_cotiza_venta.frx":B92B
         Height          =   2895
         Left            =   195
         TabIndex        =   30
         Top             =   240
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "unidad_codigo"
            Caption         =   "Codigo Unidad"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "No. Negocia"
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
            DataField       =   "edif_codigo"
            Caption         =   "Codigo Edificio"
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
            DataField       =   "cotiza_codigo"
            Caption         =   "Nro. Cotización"
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
            DataField       =   "codigo_costo"
            Caption         =   "Codigo.Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "costo_porcentaje"
            Caption         =   "% Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,##0.0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "costo_monto2"
            Caption         =   "Costo.Unitario.Eur"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "costo_monto"
            Caption         =   "Costo.Unitario.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "costo_monto_usd"
            Caption         =   "Costo.Unitario.Usd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "costo_monto_usd2"
            Caption         =   "Costo.Grupo.Eur"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "costo_monto3"
            Caption         =   "Costo.Grupo.Bs."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "costo_monto_usd3"
            Caption         =   "Costo.Grupo.Usd"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "costo_observaciones"
            Caption         =   "Detalle"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   6254.929
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos01 
      Height          =   330
      Left            =   0
      Top             =   9600
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
      ConnectStringType=   3
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
      Caption         =   "Ado_datos01"
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
Attribute VB_Name = "mw_solicitud_cotiza_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos0 As New ADODB.Recordset
Dim rs_datos01 As New ADODB.Recordset
Dim rs_datos As New ADODB.Recordset
Dim rs_datosA As New ADODB.Recordset
Dim rs_datosE As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos1A As New ADODB.Recordset
Dim rs_datos1E As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
'Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos31 As New ADODB.Recordset
Dim rs_datos41 As New ADODB.Recordset
Dim rs_datos51 As New ADODB.Recordset
Dim rs_datos61 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset
Dim RsTmp As New ADODB.Recordset
Dim rs_det1 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim VAR_FRA As String
Dim var_cod As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_SW2 As String
Public VAR_CONTI As String
Dim VAR_DA, VAR_UORIGEN As String
Dim VAR_DPTO As String

Dim VAR_COD2, VAR_PRDA As Integer
Dim VARCTRL, VAR_AME, VAR_ASI, VAR_EUR As Integer
Dim VAR_EQP, VAR_TOTEQP As Integer

Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_DOLCLI, VAR_DOLCLI2, VAR_BSCLI As Double
Dim VAR_DOLTOT, VAR_BSTOT As Double
Dim VAR_LOCAL, VAR_DOLCGE As Double
Dim VAR_SUBD, VAR_SUBB, SUBTOTD As Double
Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
     'If VAR_SW <> "MOD" Then
     If VAR_SW <> "MOD" And VAR_PAISC <> "NN" Then
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
'                Call ABRIR_TABLA_DET3
'            Case "4"
'
'        End Select
'        Call ABRIR_TABLA_AUX2
        If Ado_datos.Recordset.RecordCount > 0 Then
            
            Call ABRIR_TABLA_DET
            If Ado_datos.Recordset!estado_codigo = "REG" Then
                BtnAprobar.Visible = True
                BtnModificar1.Visible = True
                FrmABMDet.Visible = True
            Else
                BtnAprobar.Visible = False
                BtnModificar1.Visible = False
                FrmABMDet.Visible = False
            End If
            GlCotiza = Ado_datos0.Recordset!cotiza_codigo       'txt_codigo1.Caption
            GlUnidad = Ado_datos0.Recordset!unidad_codigo
            FraDet1.Visible = True
            FraDet1E.Visible = False
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub Ado_datosA_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'<-- Inicio                Identificación del Cliente                Fin -->
     If VAR_SW <> "MOD" Then
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
'                Call ABRIR_TABLA_DET3
'            Case "4"
'
'        End Select
'        Call ABRIR_TABLA_AUX2
        If Ado_datosA.Recordset.RecordCount > 0 Then
            'Call ABRIR_TABLA_DET
            If Ado_datosA.Recordset!estado_codigo = "REG" Then
                BtnAprobarA.Visible = True
                BtnModificar1A.Visible = True
            Else
                BtnAprobarA.Visible = False
                BtnModificar1A.Visible = False
            End If
            txt_codigo1.Caption = Ado_datosA.Recordset!cotiza_codigo
            Call ABRIR_TABLA_DET
            FraDet1.Visible = True
            FraDet1E.Visible = False
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        'Call ABRIR_TABLA_DET
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub Ado_datosE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'<-- Inicio                Identificación del Cliente                Fin -->
     If VAR_SW <> "MOD" Then
'        Call ABRIR_TABLA_AUX2
        If Ado_datosE.Recordset.RecordCount > 0 Then
            
            If Ado_datosE.Recordset!estado_codigo = "REG" Then
                BtnAprobarE.Visible = True
                BtnModificar1E.Visible = True
            Else
                BtnAprobarE.Visible = False
                BtnModificar1E.Visible = False
            End If
            txt_codigo1.Caption = Ado_datosE.Recordset!cotiza_codigo
            Call ABRIR_TABLA_DET
            FraDet1.Visible = False
            FraDet1E.Visible = True
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1E.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    VARCTRL = 0
    Select Case SSTab1.Tab
        Case 0
            marca1 = Ado_datos.Recordset.Bookmark
            If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
                FraNavega.Enabled = False
'                FraNavega1.Enabled = False
'                FraModeloCosto.Enabled = False
                VARCTRL = 1
                VAR_CONTI = "AMERICA"
            End If
        Case 1
            marca1 = Ado_datosA.Recordset.Bookmark
            If rs_datosA.RecordCount > 0 And rs_datosA!estado_codigo = "REG" Then
                FraNavegaA.Enabled = False
'                FraNavega1A.Enabled = False
'                FraModeloCostoA.Enabled = False
                VARCTRL = 3
                VAR_CONTI = "ASIA"
            End If
        Case 2
            marca1 = Ado_datosE.Recordset.Bookmark
            If rs_datosE.RecordCount > 0 And rs_datosE!estado_codigo = "REG" Then
                FraNavegaE.Enabled = False
    '            FraModeloCostoE.Enabled = False
    '            FraNavega1E.Enabled = False
                VARCTRL = 2
                VAR_CONTI = "EUROPA"
            End If
    End Select
    swnuevo = 1
    fraOpciones.Enabled = False
    Fra_datos2.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
  If VARCTRL = 1 Then
    aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.txt_codigo.Caption     ' Nro. Negociacion (Cod.solicitud)
    aw_p_ao_solicitud_cotiza_detalle.txt_campo1.Caption = txt_campo1.Text   ' Me.dtc_codigo1.Text       ' Codigo Unidad
    aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.txt_campo12    ' Descripcion Unidad
    aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.txt_codigo1.Caption        ' Nro. Cotización
    aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = GlEdificio    'Me.dtc_codigo3.Text       ' Codigo Edificio
    aw_p_ao_solicitud_cotiza_detalle.txt_campo5.Caption = VAR_CONTI     'Continente
    If Ado_datos.Recordset!cotiza_precio_fob_dol = "0" Or IsNull(Ado_datos.Recordset!cotiza_precio_fob_dol) Then
    'If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
    Else
        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Ado_datos.Recordset!cotiza_precio_fob_dol   ' Monto Modelo1(ME)
    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = Me.txt_dcto_bs1.Text   ' Monto Modelo2(ME)
'    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = Me.txt_seguro_bs1.Text   ' Monto Modelo3(ME)
'    End If
    Ado_detalle1.Recordset.AddNew
    aw_p_ao_solicitud_cotiza_detalle.Show vbModal
    Select Case SSTab1.Tab
        Case 0
            FraNavega.Enabled = True
'            FraNavega1.Enabled = True
'            FraModeloCosto.Enabled = True
            Call ABRIR_TABLA
            Ado_datos.Recordset.Move marca1 - 1
        Case 1
            FraNavegaA.Enabled = True
'            FraNavega1A.Enabled = True
'            FraModeloCostoA.Enabled = True
            Call ABRIR_TABLA
            Ado_datosA.Recordset.Move marca1 - 1
        Case 2
            FraNavegaE.Enabled = False
'            FraModeloCostoE.Enabled = False
'            FraNavega1E.Enabled = False
            Call ABRIR_TABLA
            Ado_datosE.Recordset.Move marca1 - 1
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
  'Else
  '  MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
'wwwwwwwwwwwwwwwwwwwwwwww EUROPA
  If VARCTRL = 2 Then
    aw_p_ao_solicitud_cotiza_det_eur.txt_codigo.Caption = Me.txt_codigo.Caption     ' Nro. Negociacion (Cod.solicitud)
    aw_p_ao_solicitud_cotiza_det_eur.txt_campo1.Caption = txt_campo1.Text   ' Me.dtc_codigo1.Text       ' Codigo Unidad
    aw_p_ao_solicitud_cotiza_det_eur.Txt_descripcion.Caption = Me.txt_campo12    ' Descripcion Unidad
    aw_p_ao_solicitud_cotiza_det_eur.Txt_Correl.Caption = Me.txt_codigo1.Caption        ' Nro. Cotización
    aw_p_ao_solicitud_cotiza_det_eur.Txt_campo2.Caption = GlEdificio    'Me.dtc_codigo3.Text       ' Codigo Edificio
    aw_p_ao_solicitud_cotiza_det_eur.txt_campo5.Caption = VAR_CONTI     'Continente
    If Ado_datosE.Recordset!cotiza_precio_fob_dol = "0" Or IsNull(Ado_datosE.Recordset!cotiza_precio_fob_dol) Then
    'If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
        aw_p_ao_solicitud_cotiza_det_eur.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
        aw_p_ao_solicitud_cotiza_det_eur.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
        aw_p_ao_solicitud_cotiza_det_eur.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
    Else
        aw_p_ao_solicitud_cotiza_det_eur.txt_monto01.Caption = Ado_datosE.Recordset!cotiza_precio_fob_dol   ' Monto Modelo1(ME)
    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = Me.txt_dcto_bs1.Text   ' Monto Modelo2(ME)
'    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = Me.txt_seguro_bs1.Text   ' Monto Modelo3(ME)
'    End If
    Ado_detalle1.Recordset.AddNew
    aw_p_ao_solicitud_cotiza_det_eur.Show vbModal
    Select Case SSTab1.Tab
        Case 0
            FraNavega.Enabled = True
'            FraNavega1.Enabled = True
'            FraModeloCosto.Enabled = True
            Call ABRIR_TABLA
            Ado_datos.Recordset.Move marca1 - 1
        Case 1
            FraNavegaA.Enabled = True
'            FraNavega1A.Enabled = True
'            FraModeloCostoA.Enabled = True
            Call ABRIR_TABLA
            Ado_datosA.Recordset.Move marca1 - 1
        Case 2
            FraNavegaE.Enabled = False
'            FraModeloCostoE.Enabled = False
'            FraNavega1E.Enabled = False
            Call ABRIR_TABLA
            Ado_datosE.Recordset.Move marca1 - 1
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If

  'wwwwwwwwwwwwwwwwwwwwww ASIA
  If VARCTRL = 3 Then
    aw_p_ao_solicitud_cotiza_det_asia.txt_codigo.Caption = Me.txt_codigo.Caption     ' Nro. Negociacion (Cod.solicitud)
    aw_p_ao_solicitud_cotiza_det_asia.txt_campo1.Caption = txt_campo1.Text   ' Me.dtc_codigo1.Text       ' Codigo Unidad
    aw_p_ao_solicitud_cotiza_det_asia.Txt_descripcion.Caption = Me.txt_campo12    ' Descripcion Unidad
    aw_p_ao_solicitud_cotiza_det_asia.Txt_Correl.Caption = Me.txt_codigo1.Caption        ' Nro. Cotización
    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo2.Caption = GlEdificio    'Me.dtc_codigo3.Text       ' Codigo Edificio
    aw_p_ao_solicitud_cotiza_det_asia.txt_campo5.Caption = VAR_CONTI     'Continente
    aw_p_ao_solicitud_cotiza_det_asia.lbl_decA.Caption = Ado_datos.Recordset!cotiza_dec       'cmd_decA.Text      ' # Decimales
    If Ado_datos.Recordset!cotiza_precio_fob_dol = "0" Or IsNull(Ado_datos.Recordset!cotiza_precio_fob_dol) Then
    'If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
        aw_p_ao_solicitud_cotiza_det_asia.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
    Else
        aw_p_ao_solicitud_cotiza_det_asia.txt_monto01.Caption = Ado_datos.Recordset!cotiza_precio_fob_dol   ' Monto Modelo1(ME)
    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_det_asia.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_det_asia.txt_monto02.Caption = Me.txt_dcto_bs1.Text   ' Monto Modelo2(ME)
'    End If
'    If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        aw_p_ao_solicitud_cotiza_det_asia.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
'    Else
'        aw_p_ao_solicitud_cotiza_det_asia.txt_monto03.Caption = Me.txt_seguro_bs1.Text   ' Monto Modelo3(ME)
'    End If
    Ado_detalle1.Recordset.AddNew
    aw_p_ao_solicitud_cotiza_det_asia.Show vbModal
    Select Case SSTab1.Tab
        Case 0
            FraNavega.Enabled = True
'            FraNavega1.Enabled = True
'            FraModeloCosto.Enabled = True
            Call ABRIR_TABLA
            Ado_datos.Recordset.Move marca1 - 1
        Case 1
            FraNavegaA.Enabled = True
'            FraNavega1A.Enabled = True
'            FraModeloCostoA.Enabled = True
            Call ABRIR_TABLA
            Ado_datosA.Recordset.Move marca1 - 1
        Case 2
'            FraNavegaE.Enabled = False
'            FraModeloCostoE.Enabled = False
'            FraNavega1E.Enabled = False
'            Call ABRIR_TABLA
'            Ado_datosE.Recordset.Move marca1 - 1
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
  'wwwwwwwwwwwwwwwwwwwwww
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'     Call ABRIR_TABLA
''    Call OptFilGral1_Click
'  Else
'     Call ABRIR_TABLA
''    Call OptFilGral2_Click
'  End If
'  'Call ABRIR_TABLA_DET
  End Sub

Private Sub BtnAddDetalle2_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    Select Case SSTab1.Tab
        Case 0
            VAR_CONTI = "AMERICA"
        Case 1
            VAR_CONTI = "ASIA"
        Case 2
            VAR_CONTI = "EUROPA"
    End Select
    aw_p_ao_solicitud_item_costos.Show vbModal
    
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    If VAR_CONTI = "AMERICA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipo= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "ASIA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoA= 'B' ", db, adOpenStatic
    End If
    If VAR_CONTI = "EUROPA" Then
        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoE= 'B' ", db, adOpenStatic
    End If
    Set Ado_datos3.Recordset = rs_datos6
End Sub

Private Sub BtnAnlDetalle_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAprobar_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  
  On Error GoTo UpdateErr
   Set rs_aux2 = New ADODB.Recordset
   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux2.RecordCount > 0 Then
        VAR_CONT2 = rs_aux2.RecordCount
   Else
        MsgBox "No se puede APROBAR debe registrar el Detalle de Costos ...", vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   VAR_SW = "MOD"
   If Ado_datos.Recordset!estado_codigo = "REG" Then       'And Ado_datos.Recordset!correl_edificacion > 0
   'If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      GlSolicitud = Ado_datos.Recordset!solicitud_codigo
      sino = MsgBox("Está Seguro de VERIFICAR y enviar datos para el Registro del Contrato ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datos.Recordset.MoveFirst
         While Not Ado_datos.Recordset.EOF
            'WWWWWWWWWWWWWWWWWWWWW Acumula Precio Equipos
            Set RsTmp = New ADODB.Recordset
            RsTmp.Open "Select sum(cotiza_fob_seg_dol) as totdol, sum(cotiza_fob_seg_bs) as totbs, sum(cotiza_cantidad) as toteqp from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "'  and solicitud_codigo = " & GlSolicitud & " and estado_codigo_verif = 'APR'  ", db, adOpenStatic
            If RsTmp.RecordCount = 0 Or IsNull(RsTmp!totdol) Then
                VAR_DOLCLI = 0
                VAR_BSCLI = 0
                VAR_TOTEQP = 0
            Else
                VAR_DOLCLI = RsTmp!totdol
                VAR_BSCLI = RsTmp!totbs
                VAR_TOTEQP = RsTmp!toteqp
            End If
            'GRABA ao_ventas_cabecera
            Set rs_aux1 = New ADODB.Recordset
            'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "    "
            SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "    "
            rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If rs_aux1.RecordCount > 0 Then
                MsgBox "Una Cotización anterior ya fue procesada, los datos de este Registro actualizarán al que fue registrado anteriormente ..."
                'rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datos.Recordset!cotiza_fob_seg_bs        'cotiza_precio_fob_bs
                'rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datos.Recordset!cotiza_fob_seg_dol      'cotiza_precio_fob_dol
                rs_aux1!venta_monto_total_bs = VAR_BSCLI * VAR_TOTEQP
                rs_aux1!venta_monto_total_dol = VAR_DOLCLI * VAR_TOTEQP
                rs_aux1!venta_cantidad_total = VAR_TOTEQP
                rs_aux1!venta_monto_cobrado_bs = 0
                rs_aux1!venta_monto_cobrado_dol = 0
                rs_aux1!venta_saldo_p_cobrar_bs = VAR_BSCLI * VAR_TOTEQP
                rs_aux1!venta_saldo_p_cobrar_dol = VAR_DOLCLI * VAR_TOTEQP
                
                var_cod = rs_aux1!venta_codigo
                VAR_SW2 = 1
            Else
                'CREA VENTA CABECERA
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                If Not rs_aux2.EOF Then
                    var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                End If
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "   ", db, adOpenStatic
                If Not rs_aux2.EOF Then
                    VAR_AUX = rs_aux2!Codigo
                End If
                rs_aux1.AddNew
                'var_cod = rs_aux1.RecordCount + 1
                rs_aux1!ges_gestion = Year(Date)
                rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                rs_aux1!EDIF_CODIGO = Ado_datos.Recordset!EDIF_CODIGO
                rs_aux1!venta_codigo = var_cod
                rs_aux1!beneficiario_codigo = VAR_AUX
                If Ado_datos.Recordset!cotiza_cantidad = 0 Then
                    rs_aux1!venta_cantidad_total = 1
                Else
                    rs_aux1!venta_cantidad_total = Ado_datos.Recordset!cotiza_cantidad
                End If
                rs_aux1!tipo_moneda = Ado_datos.Recordset!tipo_moneda
                rs_aux1!venta_monto_total_bs = Ado_datos.Recordset!cotiza_precio_total_bs * rs_aux1!venta_cantidad_total
                rs_aux1!venta_monto_total_dol = Ado_datos.Recordset!cotiza_precio_total_dol * rs_aux1!venta_cantidad_total
                'rs_aux1!venta_monto_total_bs = Ado_datos.Recordset!cotiza_fob_seg_bs * Ado_datos.Recordset!cotiza_cantidad
                'rs_aux1!venta_monto_total_dol = Ado_datos.Recordset!cotiza_fob_seg_dol * Ado_datos.Recordset!cotiza_cantidad
                rs_aux1!venta_monto_cobrado_bs = 0
                rs_aux1!venta_monto_cobrado_dol = 0
                rs_aux1!venta_saldo_p_cobrar_bs = Ado_datos.Recordset!cotiza_fob_seg_bs * Ado_datos.Recordset!cotiza_cantidad
                rs_aux1!venta_saldo_p_cobrar_dol = Ado_datos.Recordset!cotiza_fob_seg_dol * Ado_datos.Recordset!cotiza_cantidad
                
                rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                rs_aux1!unimed_codigo = "EQP"
                rs_aux1!estado_codigo = "REG"
                rs_aux1!fecha_registro = Date
                rs_aux1!usr_codigo = glusuario
                rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
                VAR_SW2 = 0
            End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
            'GRABA VENTA DETALLE
            If var_cod = "" Then
                var_cod = rs_aux1!venta_codigo
            End If
            Set rs_aux3 = New ADODB.Recordset
            If rs_aux3.State = 1 Then rs_aux3.Close
            'rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & "  and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
            rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux3.RecordCount > 0 Then
                VAR_LOCAL = Ado_datos.Recordset!cotiza_cantidad - rs_aux3.RecordCount
                If VAR_LOCAL < 0 Then
                    sino = MsgBox("Desea eliminar los registros anteriores ? SI = Elimina los anteriores y genera otros nuevos. NO = Cancela el proceso y luego elimine el detalle de los equipos en VENTAS NUEVAS. ", vbYesNo + vbQuestion, "Atención")
                    If sino = vbYes Then
                        db.Execute "Delete ao_ventas_detalle Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                        VAR_LOCAL = Val(dtc_desc15)
                    Else
                        Exit Sub
                    End If
                End If
                If VAR_LOCAL > 0 Then
                    rs_aux3.MoveFirst
                    VAR_EQP = 0
                    While VAR_LOCAL > VAR_EQP
                        VAR_EQP = VAR_EQP + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = Year(Date)
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                        rs_aux3!cotiza_codigo = Ado_datos.Recordset!cotiza_codigo          'VAR_AUX
                        'rs_aux3!bien_codigo = "NA" + Trim(Str(rs_aux3.RecordCount))      'Ado_datos.Recordset!bien_codigo
                        rs_aux3!bien_codigo = "NA" + Trim(Str(Ado_datos.Recordset!cotiza_codigo))           'Trim(Str(rs_aux3.RecordCount))      'Ado_datos.Recordset!bien_codigo
                        rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                        rs_aux3!venta_precio_unitario_bs = Ado_datos.Recordset!cotiza_precio_fob_bs         'cotiza_fob_seg_bs
                        rs_aux3!venta_descuento_bs = Ado_datos.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_bs = Ado_datos.Recordset!cotiza_precio_fob_bs            'cotiza_fob_seg_bs
                        rs_aux3!venta_precio_unitario_dol = Ado_datos.Recordset!cotiza_precio_fob_dol       'cotiza_fob_seg_dol
                        rs_aux3!venta_descuento_dol = Ado_datos.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_dol = Ado_datos.Recordset!cotiza_precio_fob_dol          'cotiza_fob_seg_dol
                        rs_aux3!concepto_venta = "Codigo: " + Ado_datos.Recordset!bien_codigo + " Modelo: " + Ado_datos.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datos.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                        rs_aux3!observaciones = "Codigo: " + Ado_datos.Recordset!bien_codigo + " Modelo: " + Ado_datos.Recordset!modelo_codigo + " Paradas: " + dtc_desc10   '
                        'ok
                        rs_aux3!grupo_codigo = "40000"
                        rs_aux3!subgrupo_codigo = "43000"
                        rs_aux3!par_codigo = "43340"
                        'ok
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo = Ado_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                        'PROVEEDOR OTIS JQA 22-FEB-2017
                        rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                        'rs_aux3!modelo_codigo_x = Ado_datos0.Recordset!beneficiario_codigo
                        '"S/M"    'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!pais_codigo = Ado_datos.Recordset!pais_codigo
                        'rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                    Wend
                Else
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datos.Recordset!cotiza_fob_seg_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datos.Recordset!cotiza_fob_seg_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datos.Recordset!cotiza_precio_fob_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datos.Recordset!cotiza_precio_fob_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set usr_codigo  = '" & glusuario & "' Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  "
                End If
'                    If Left(rs_aux3!bien_codigo, 4) = "AO36" Or Left(rs_aux3!bien_codigo, 4) = "36NO" Then
'                    Else
'                    End If
'                'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
            Else
                'VAR_AUX = rs_aux3.RecordCount + 1
                VAR_EQP = 0
                While Ado_datos.Recordset!cotiza_cantidad > VAR_EQP
                    VAR_EQP = VAR_EQP + 1
                    rs_aux3.AddNew
                    rs_aux3!ges_gestion = Year(Date)
                    rs_aux3!venta_codigo = var_cod
                    rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                    rs_aux3!cotiza_codigo = Ado_datos.Recordset!cotiza_codigo          'VAR_AUX
                    rs_aux3!bien_codigo = "NA" + Trim(Str(VAR_EQP))      'Ado_datos.Recordset!bien_codigo
                    rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                    rs_aux3!venta_precio_unitario_bs = Ado_datos.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_descuento_bs = Ado_datos.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_bs = Ado_datos.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_precio_unitario_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
                    rs_aux3!venta_descuento_dol = Ado_datos.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
                    rs_aux3!concepto_venta = "Codigo: " + Ado_datos.Recordset!bien_codigo + "Modelo: " + Ado_datos.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datos.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                    'ok
                    rs_aux3!grupo_codigo = "40000"
                    rs_aux3!subgrupo_codigo = "43000"
                    rs_aux3!par_codigo = "43340"
                    'ok
                    rs_aux3!tipo_descuento = 0
                    rs_aux3!almacen_codigo = 0
                    rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
                    rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                    'PROVEEDOR OTIS JQA 22-FEB-2017
                    rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                    'rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
                    rs_aux3!modelo_elegido = "S"
'                    rs_aux3!modelo_elegido_h = "N"
'                    rs_aux3!modelo_elegido_x = "N"
                    rs_aux3!pais_codigo = Ado_datos.Recordset!pais_codigo
                    'rs_aux3!estado_codigo = "REG"
                    rs_aux3!fecha_registro = Date
                    rs_aux3!usr_codigo = glusuario
                    rs_aux3.Update
                Wend
            End If
'            VAR_NO2 = VAR_NO2 + rs_datos!h_nro_total_equipos - 1
'            VAR_NO3 = "36NO-" + Trim(Str(VAR_NO2))
'            If rs_datos!h_nro_total_equipos > 1 Then
'                'If Right(VAR_NO3, 1) = 0 Then
'                    rs_datos!unidad_codigo_ant = VAR_NO1 + "-" + Right(VAR_NO3, 2)
'                'Else
'                '    rs_datos!unidad_codigo_ant = VAR_NO1 + "/" + Right(VAR_NO3, 1)
'                'End If
'            Else
'                rs_datos!unidad_codigo_ant = VAR_NO1
'            End If
'            rs_datos!unidad_codigo_ant = rs_datos!unidad_codigo + Trim(Str(rs_datos!solicitud_codigo))
    '        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
    '        db.Execute "Update ao_solicitud_cotiza_venta Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
    '        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
                'AAAAAAAAAAAAAAAQQQQQQQQQQQQQQUUUUUUUUUUUUUUUUUUIIIIIIIIIIII
                'VAR_COD3 = "NA" + Trim(Str(i))
                'rs_aux1!bien_codigo = VAR_COD3  '"NA" + Trim(Str(VAR_COD2))
                '
            'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
            If VAR_SW2 = 0 Then
                Set rs_aux2 = New ADODB.Recordset
                If rs_aux2.State = 1 Then rs_aux2.Close
                SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
                rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux2.RecordCount > 0 Then
                    rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                    rs_datos!doc_numero = rs_aux2!correl_doc
                    'Txt_campo1.Caption = rs_aux2!correl_doc
                    rs_aux2.Update
                End If
                'rs_datos!doc_numero = Txt_campo1.Caption
                'REVISAR !!! JQA 2014_07_08
                'VAR_ARCH = RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
                If IsNull(RTrim(rs_datos!doc_codigo)) Or RTrim(rs_datos!doc_codigo) = "" Then
                    VAR_ARCH = "COM_R-223-0"
                Else
                    VAR_ARCH = "COM_" + RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
                End If
                Ado_datos.Recordset!archivo_respaldo = VAR_ARCH + ".PDF"
                Ado_datos.Recordset!archivo_respaldo_cargado = "N"
                'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
    '            Set rs_aux2 = New ADODB.Recordset
    '            If rs_aux2.State = 1 Then rs_aux2.Close
    '            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos1.Recordset!doc_codigo2 & "'  "
    '            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
    '            If rs_aux2.RecordCount > 0 Then
    '                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
    '                rs_datos!doc_numero2 = rs_aux2!correl_doc
    '                rs_aux2.Update
    '            End If
    '            VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos1!doc_codigo2) + "-") + LTrim(Str(rs_datos1!doc_numero2))
    '            rs_datos!archivo_respaldo = VAR_ARCH2 + ".PDF"
    '            rs_datos!archivo_respaldo_cargado = "N"
            End If
            Ado_datos.Recordset!estado_codigo_verif = "APR"
            Ado_datos.Recordset!fecha_registro = Date
            Ado_datos.Recordset!usr_codigo = glusuario
            Ado_datos.Recordset.UpdateBatch adAffectAll
        
            Ado_datos.Recordset.MoveNext
         Wend
         
         'ACTUALIZA PROVEEDOR BRASIL
         'db.Execute "update ao_solicitud_cotiza_modelo set beneficiario_codigo = '101853029' where pais_codigo = 'BRA' and unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " "
         'ACTUALIZA PROVEEDOR ARGENTINA
         'db.Execute "update ao_solicitud_cotiza_modelo set beneficiario_codigo = '30-51662787-1' where pais_codigo = 'ARG' and unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " "
          db.Execute "UPDATE ao_solicitud_cotiza_modelo set ao_solicitud_cotiza_modelo.beneficiario_codigo  = gc_beneficiario.beneficiario_codigo FROM ao_solicitud_cotiza_modelo INNER JOIN gc_beneficiario ON ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.pais_codigo  AND ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.depto_sigla "
      End If
      
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub BtnAprobarA_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
   
   On Error GoTo UpdateErr
   Set rs_aux2 = New ADODB.Recordset
   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datosA.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datosA.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux2.RecordCount > 0 Then
        VAR_CONT2 = rs_aux2.RecordCount
   Else
        MsgBox "No se puede APROBAR debe registrar el Detalle de Costos ...", vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   VAR_SW = "MOD"
   If Ado_datosA.Recordset!estado_codigo = "REG" Then       'And Ado_datos.Recordset!correl_edificacion > 0
   'If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de VERIFICAR y enviar datos para el Registro del Contrato ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datosA.Recordset.MoveFirst
         While Not Ado_datosA.Recordset.EOF
'                db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_precio_total_bs = cotiza_precio_fob_bs Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
'                db.Execute "Update ao_solicitud_cotiza_venta Set cotiza_precio_total_dol = cotiza_precio_fob_dol Where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
                'AQUIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII WWWWWWWWWWWW
'                If Ado_datosA.Recordset!pais_continente = "AMERICA" Then
'                    'VAR_DOLCLI = Ado_datosA.Recordset!cotiza_precio_total_dol - Ado_datosA.Recordset!cotiza_precio_fob_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol
'                    'VAR_BSCLI = Ado_datosA.Recordset!cotiza_precio_total_bs - Ado_datosA.Recordset!cotiza_precio_fob_bs - Ado_datosA.Recordset!cotiza_precio_seg_bs
'                End If
                If Ado_datosA.Recordset!pais_continente = "ASIA" Then
                    'VAR_DOLCLI = Ado_datosA.Recordset!cotiza_precio_total_dol - Ado_datosA.Recordset!cotiza_precio_fob_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = Ado_datosA.Recordset!cotiza_precio_total_bs - Ado_datosA.Recordset!cotiza_precio_fob_bs - Ado_datosA.Recordset!cotiza_precio_seg_bs
                End If
                If Ado_datosA.Recordset!pais_continente = "EUROPA" Then
                    'VAR_DOLCLI = Ado_datos1.Recordset!cotiza_precio_total_dol - Ado_datos1.Recordset!cotiza_precio_fob_dol - Ado_datos1.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = Ado_datos1.Recordset!cotiza_precio_total_bs - Ado_datos1.Recordset!cotiza_precio_fob_bs - Ado_datos1.Recordset!cotiza_precio_seg_bs
                End If
                'WWWWWWWWWWWWWWWWWWWWW
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "    "
                SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datosA.Recordset!solicitud_codigo & "    "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue procesada, los datos de este Registro actualizarán al que fue registrado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datosA.Recordset!cotiza_precio_fob_bs      'cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datosA.Recordset!cotiza_precio_fob_dol       'cotiza_precio_total_dol
                    VAR_SW2 = 1
                Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datosA.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datosA.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datosA.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = IIf(IsNull(rs_aux2!Codigo), "0", rs_aux2!Codigo)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datosA.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datosA.Recordset!solicitud_codigo
                    rs_aux1!EDIF_CODIGO = Ado_datosA.Recordset!EDIF_CODIGO
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = VAR_AUX
                    If Ado_datosA.Recordset!cotiza_cantidad = 0 Then
                        rs_aux1!venta_cantidad_total = 1
                    Else
                        rs_aux1!venta_cantidad_total = Ado_datosA.Recordset!cotiza_cantidad
                    End If
                    rs_aux1!venta_monto_total_bs = Ado_datosA.Recordset!cotiza_precio_total_bs * rs_aux1!venta_cantidad_total
                    rs_aux1!venta_monto_total_dol = Ado_datosA.Recordset!cotiza_precio_total_dol * rs_aux1!venta_cantidad_total
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    'jqa 2015-06-01 revisar calculos
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux1!venta_monto_total_bs          'Ado_datosA.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux1!venta_monto_total_dol        'Ado_datosA.Recordset!cotiza_precio_total_dol
                    rs_aux1!unidad_codigo_ant = Ado_datosA.Recordset!unidad_codigo_ant
                    rs_aux1!unimed_codigo = "MES"
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
                    VAR_SW2 = 0
                End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
'            Case "4"
'        End Select
        'GRABA VENTA DETALLE
        If var_cod = "" Then
            var_cod = rs_aux1!venta_codigo
        End If
        
'        Set rs_aux3 = New ADODB.Recordset
'        If rs_aux3.State = 1 Then rs_aux3.Close
'        rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'        'rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & "  and bien_codigo = '" & Ado_datosA.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux3.RecordCount > 0 Then
'            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'            rs_aux3!venta_precio_unitario_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs
'            rs_aux3!venta_descuento_bs = 0
'            rs_aux3!venta_precio_total_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs
'            rs_aux3!venta_precio_unitario_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol
'            rs_aux3!venta_descuento_dol = 0
'            rs_aux3!venta_precio_total_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol
'            rs_aux3!modelo_codigo1 = Ado_datosA.Recordset!modelo_codigo
'            rs_aux3!modelo_codigo_h = Ado_datosA.Recordset!modelo_codigo_h
'            rs_aux3!modelo_codigo_x = Ado_datosA.Recordset!modelo_codigo_x
'            rs_aux3!fecha_registro = Date
'            rs_aux3!usr_codigo = glusuario
'            rs_aux3.Update
'        Else
'            VAR_AUX = rs_aux3.RecordCount + 1
'            rs_aux3.AddNew
'            rs_aux3!ges_gestion = Year(Date)
'            rs_aux3!venta_codigo = var_cod
'            rs_aux3!venta_codigo_det = Ado_datosA.Recordset!cotiza_codigo      'VAR_AUX
'            rs_aux3!bien_codigo = Ado_datosA.Recordset!bien_codigo
'            rs_aux3!venta_det_cantidad = Ado_datosA.Recordset!cotiza_cantidad
'            rs_aux3!venta_precio_unitario_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs
'            rs_aux3!venta_descuento_bs = 0
'            rs_aux3!venta_precio_total_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs
'            rs_aux3!venta_precio_unitario_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol
'            rs_aux3!venta_descuento_dol = 0
'            rs_aux3!venta_precio_total_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol
'            rs_aux3!concepto_venta = "Codigo: " + Ado_datosA.Recordset!bien_codigo + "Modelo: " + Ado_datosA.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datosA.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
'            'ok
'            rs_aux3!grupo_codigo = "40000"
'            rs_aux3!subgrupo_codigo = "43000"
'            rs_aux3!par_codigo = "43340"
'            'ok
'            rs_aux3!tipo_descuento = 0
'            rs_aux3!almacen_codigo = 0
'            rs_aux3!modelo_codigo1 = Ado_datosA.Recordset!modelo_codigo
'            rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
'            rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
'            rs_aux3!modelo_elegido = "N"
'            rs_aux3!modelo_elegido_h = "N"
'            rs_aux3!modelo_elegido_x = "N"
'            'rs_aux3!estado_codigo = "REG"
'            rs_aux3!fecha_registro = Date
'            rs_aux3!usr_codigo = glusuario
'            rs_aux3.Update
'        End If
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux3 = New ADODB.Recordset
            If rs_aux3.State = 1 Then rs_aux3.Close
            'rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & "  and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
            rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux3.RecordCount > 0 Then
                VAR_LOCAL = Ado_datosA.Recordset!cotiza_cantidad - rs_aux3.RecordCount
                If VAR_LOCAL < 0 Then
                    sino = MsgBox("Desea eliminar los registros anteriores ? SI = Elimina los anteriores y genera otros nuevos. NO = Cancela el proceso y luego elimine el detalle de los equipos en VENTAS NUEVAS. ", vbYesNo + vbQuestion, "Atención")
                    If sino = vbYes Then
                        db.Execute "Delete ao_ventas_detalle Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                        VAR_LOCAL = Val(dtc_desc15)
                    Else
                        Exit Sub
                    End If
                End If
                If VAR_LOCAL > 0 Then
                    rs_aux3.MoveFirst
                    VAR_EQP = 0
                    While VAR_LOCAL > VAR_EQP
                        VAR_EQP = VAR_EQP + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = Year(Date)
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                        rs_aux3!cotiza_codigo = Ado_datosA.Recordset!cotiza_codigo          'VAR_AUX
                        rs_aux3!bien_codigo = "NA" + Trim(Str(Ado_datosA.Recordset!cotiza_codigo))           'Trim(Str(rs_aux3.RecordCount))      'Ado_datos.Recordset!bien_codigo
                        rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                        rs_aux3!venta_precio_unitario_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs        'cotiza_fob_seg_bs
                        rs_aux3!venta_descuento_bs = Ado_datosA.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_bs = Ado_datosA.Recordset!cotiza_precio_fob_bs           'cotiza_fob_seg_bs
                        rs_aux3!venta_precio_unitario_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol      'cotiza_fob_seg_dol
                        rs_aux3!venta_descuento_dol = Ado_datosA.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_dol = Ado_datosA.Recordset!cotiza_precio_fob_dol         'cotiza_fob_seg_dol
                        rs_aux3!concepto_venta = "Codigo: " + Ado_datosA.Recordset!bien_codigo + " Modelo: " + Ado_datosA.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datosA.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                        'ok
                        rs_aux3!grupo_codigo = "40000"
                        rs_aux3!subgrupo_codigo = "43000"
                        rs_aux3!par_codigo = "43340"
                        'ok
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo = Ado_datosA.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo1 = Ado_datosA.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                        'rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!pais_codigo = Ado_datosA.Recordset!pais_codigo
                        'rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                    Wend
                Else
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datosA.Recordset!cotiza_fob_seg_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datosA.Recordset!cotiza_fob_seg_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datosA.Recordset!cotiza_precio_fob_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datosA.Recordset!cotiza_precio_fob_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set usr_codigo  = '" & glusuario & "' Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  "
                End If
'                    If Left(rs_aux3!bien_codigo, 4) = "AO36" Or Left(rs_aux3!bien_codigo, 4) = "36NO" Then
'                    Else
'                    End If
'                'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                rs_aux3!venta_precio_unitario_bs = Ado_datos.Recordset!cotiza_fob_seg_bs    'cotiza_precio_fob_bs
'                rs_aux3!venta_descuento_bs = 0
'                rs_aux3!venta_precio_total_bs = Ado_datos.Recordset!cotiza_fob_seg_bs
'                rs_aux3!venta_precio_unitario_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
'                rs_aux3!venta_descuento_dol = 0
'                rs_aux3!venta_precio_total_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
'                rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
'                rs_aux3!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h
'                rs_aux3!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x
'                rs_aux3!fecha_registro = Date
'                rs_aux3!usr_codigo = glusuario
'                rs_aux3.Update
            Else
                'VAR_AUX = rs_aux3.RecordCount + 1
                VAR_EQP = 0
                Set rs_aux8 = New ADODB.Recordset
                If rs_aux8.State = 1 Then rs_aux8.Close
                rs_aux8.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " AND par_codigo = '43340'  ", db, adOpenKeyset, adLockOptimistic
                If rs_aux8.RecordCount > 0 Then
                    VAR_AUX = rs_aux8.RecordCount
                Else
                    VAR_AUX = 0
                End If
                While Ado_datosA.Recordset!cotiza_cantidad > VAR_EQP
                    VAR_EQP = VAR_EQP + 1
                    rs_aux3.AddNew
                    rs_aux3!ges_gestion = Year(Date)
                    rs_aux3!venta_codigo = var_cod
                    rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                    rs_aux3!cotiza_codigo = Ado_datosA.Recordset!cotiza_codigo          'VAR_AUX
                    VAR_AUX = VAR_AUX + 1
                    rs_aux3!bien_codigo = "NA" + Trim(Str(VAR_AUX))      'Ado_datos.Recordset!bien_codigo
                    rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                    rs_aux3!venta_precio_unitario_bs = Ado_datosA.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_descuento_bs = Ado_datosA.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_bs = Ado_datosA.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_precio_unitario_dol = Ado_datosA.Recordset!cotiza_fob_seg_dol
                    rs_aux3!venta_descuento_dol = Ado_datosA.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_dol = Ado_datosA.Recordset!cotiza_fob_seg_dol
                    rs_aux3!concepto_venta = "Codigo: " + Ado_datosA.Recordset!bien_codigo + " Modelo: " + Ado_datosA.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datosA.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                    'ok
                    rs_aux3!grupo_codigo = "40000"
                    rs_aux3!subgrupo_codigo = "43000"
                    rs_aux3!par_codigo = "43340"
                    'ok
                    rs_aux3!tipo_descuento = 0
                    rs_aux3!almacen_codigo = 0
                    rs_aux3!modelo_codigo = Ado_datosA.Recordset!modelo_codigo
                    rs_aux3!modelo_codigo1 = Ado_datosA.Recordset!modelo_codigo
                    rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                    rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                    'rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
                    rs_aux3!modelo_elegido = "N"
'                    rs_aux3!modelo_elegido_h = "N"
'                    rs_aux3!modelo_elegido_x = "N"
                    rs_aux3!pais_codigo = Ado_datos0.Recordset!pais_codigo
                    'rs_aux3!estado_codigo = "REG"
                    rs_aux3!fecha_registro = Date
                    rs_aux3!usr_codigo = glusuario
                    rs_aux3.Update
                Wend
            End If
        'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
        If VAR_SW2 = 0 Then
            If IsNull(Ado_datosA.Recordset!doc_codigo) Then
                Ado_datosA.Recordset!doc_codigo = "R-222"
            End If
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datosA.Recordset!doc_codigo & "'  "
            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                rs_datosA!doc_numero = rs_aux2!correl_doc
                'Txt_campo1.Caption = rs_aux2!correl_doc
                rs_aux2.Update
            End If
            'rs_datos!doc_numero = Txt_campo1.Caption
            'REVISAR !!! JQA 2014_07_08
            'VAR_ARCH = RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
            VAR_ARCH = "COM_" + RTrim(RTrim(rs_datosA!doc_codigo) + "-") + LTrim(Str(rs_datosA!doc_numero))
            Ado_datosA.Recordset!archivo_respaldo = VAR_ARCH + ".PDF"
            Ado_datosA.Recordset!archivo_respaldo_cargado = "N"
            'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
'            Set rs_aux2 = New ADODB.Recordset
'            If rs_aux2.State = 1 Then rs_aux2.Close
'            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos1.Recordset!doc_codigo2 & "'  "
'            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux2.RecordCount > 0 Then
'                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                rs_datosA!doc_numero2 = rs_aux2!correl_doc
'                rs_aux2.Update
'            End If
'            VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos1!doc_codigo2) + "-") + LTrim(Str(rs_datos1!doc_numero2))
'            rs_datosA!archivo_respaldo = VAR_ARCH2 + ".PDF"
'            rs_datosA!archivo_respaldo_cargado = "N"
        End If
            Ado_datosA.Recordset!estado_codigo_verif = "APR"
            Ado_datosA.Recordset!fecha_registro = Date
            Ado_datosA.Recordset!usr_codigo = glusuario
            Ado_datosA.Recordset.UpdateBatch adAffectAll
        
            Ado_datosA.Recordset.MoveNext
         Wend
      End If
      'ACTUALIZA PROVEEDOR CHINA
      'db.Execute "update ao_solicitud_cotiza_modelo set beneficiario_codigo = '212391920010' where pais_codigo = 'CHN' and unidad_codigo = '" & Ado_datosA.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datosA.Recordset!solicitud_codigo & " "
      db.Execute "UPDATE ao_solicitud_cotiza_modelo set ao_solicitud_cotiza_modelo.beneficiario_codigo  = gc_beneficiario.beneficiario_codigo FROM ao_solicitud_cotiza_modelo INNER JOIN gc_beneficiario ON ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.pais_codigo  AND ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.depto_sigla "
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobarE_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    On Error GoTo UpdateErr
   Set rs_aux2 = New ADODB.Recordset
   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datosE.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datosE.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux2.RecordCount > 0 Then
        VAR_CONT2 = rs_aux2.RecordCount
   Else
        MsgBox "No se puede APROBAR debe registrar el Detalle de Costos ...", vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   VAR_SW = "MOD"
   If Ado_datosE.Recordset!estado_codigo = "REG" Then       'And Ado_datos.Recordset!correl_edificacion > 0
   'If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de VERIFICAR y enviar datos para el Registro del Contrato ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         Ado_datosE.Recordset.MoveFirst
         While Not Ado_datosE.Recordset.EOF
                If Ado_datosE.Recordset!pais_continente = "ASIA" Then
                    'VAR_DOLCLI = Ado_datosE.Recordset!cotiza_precio_total_dol - Ado_datosE.Recordset!cotiza_precio_fob_dol - Ado_datosE.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = Ado_datosE.Recordset!cotiza_precio_total_bs - Ado_datosE.Recordset!cotiza_precio_fob_bs - Ado_datosE.Recordset!cotiza_precio_seg_bs
                End If
                If Ado_datosE.Recordset!pais_continente = "EUROPA" Then
                    'VAR_DOLCLI = Ado_datos1.Recordset!cotiza_precio_total_dol - Ado_datos1.Recordset!cotiza_precio_fob_dol - Ado_datos1.Recordset!cotiza_precio_seg_dol
                    'VAR_BSCLI = Ado_datos1.Recordset!cotiza_precio_total_bs - Ado_datos1.Recordset!cotiza_precio_fob_bs - Ado_datos1.Recordset!cotiza_precio_seg_bs
                End If
                'WWWWWWWWWWWWWWWWWWWWW
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "    "
                SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datosE.Recordset!solicitud_codigo & "    "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue procesada, los datos de este Registro actualizarán al que fue registrado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datosE.Recordset!cotiza_precio_fob_bs      'cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datosE.Recordset!cotiza_precio_fob_dol       'cotiza_precio_total_dol
                    VAR_SW2 = 1
                Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datosE.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datosE.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & Ado_datosE.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datosE.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datosE.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datosE.Recordset!solicitud_codigo
                    rs_aux1!EDIF_CODIGO = Ado_datosE.Recordset!EDIF_CODIGO
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = VAR_AUX
                    If Ado_datosE.Recordset!cotiza_cantidad = 0 Then
                        rs_aux1!venta_cantidad_total = 1
                    Else
                        rs_aux1!venta_cantidad_total = Ado_datosE.Recordset!cotiza_cantidad
                    End If
                    rs_aux1!venta_monto_total_bs = Ado_datosE.Recordset!cotiza_precio_total_bs * rs_aux1!venta_cantidad_total
                    rs_aux1!venta_monto_total_dol = Ado_datosE.Recordset!cotiza_precio_total_dol * rs_aux1!venta_cantidad_total
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    'jqa 2015-06-01 revisar calculos
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux1!venta_monto_total_bs          'Ado_datosA.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux1!venta_monto_total_dol        'Ado_datosA.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_total_bs = Ado_datosE.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = Ado_datosE.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    'jqa 2015-06-01 revisar calculos
'                    rs_aux1!venta_saldo_p_cobrar_bs = Ado_datosE.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = Ado_datosE.Recordset!cotiza_precio_total_dol
                    rs_aux1!unidad_codigo_ant = Ado_datosE.Recordset!unidad_codigo_ant
                    rs_aux1!unimed_codigo = "MES"
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
                    VAR_SW2 = 0
                End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos1.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos1.Recordset!solicitud_codigo & "  "
'            Case "4"
'        End Select
        'GRABA VENTA DETALLE
        If var_cod = "" Then
            var_cod = rs_aux1!venta_codigo
        End If
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
            Set rs_aux3 = New ADODB.Recordset
            If rs_aux3.State = 1 Then rs_aux3.Close
            'rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & "  and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
            rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux3.RecordCount > 0 Then
                VAR_LOCAL = Ado_datosE.Recordset!cotiza_cantidad - rs_aux3.RecordCount
                If VAR_LOCAL < 0 Then
                    sino = MsgBox("Desea eliminar los registros anteriores ? SI = Elimina los anteriores y genera otros nuevos. NO = Cancela el proceso y luego elimine el detalle de los equipos en VENTAS NUEVAS. ", vbYesNo + vbQuestion, "Atención")
                    If sino = vbYes Then
                        db.Execute "Delete ao_ventas_detalle Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                        VAR_LOCAL = Val(dtc_desc15)
                    Else
                        Exit Sub
                    End If
                End If
                If VAR_LOCAL > 0 Then
                    rs_aux3.MoveFirst
                    VAR_EQP = 0
                    While VAR_LOCAL > VAR_EQP
                        VAR_EQP = VAR_EQP + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = Year(Date)
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                        rs_aux3!cotiza_codigo = Ado_datosE.Recordset!cotiza_codigo          'VAR_AUX
                        'rs_aux3!bien_codigo = "NA" + Trim(Str(rs_aux3.RecordCount))      'Ado_datos.Recordset!bien_codigo
                        rs_aux3!bien_codigo = "NA" + Trim(Str(Ado_datosE.Recordset!cotiza_codigo))           'Trim(Str(rs_aux3.RecordCount))      'Ado_datos.Recordset!bien_codigo
                        rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                        rs_aux3!venta_precio_unitario_bs = Ado_datosE.Recordset!cotiza_precio_fob_bs        'cotiza_fob_seg_bs
                        rs_aux3!venta_descuento_bs = Ado_datosE.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_bs = Ado_datosE.Recordset!cotiza_precio_fob_bs           'cotiza_fob_seg_bs
                        rs_aux3!venta_precio_unitario_dol = Ado_datosE.Recordset!cotiza_precio_fob_dol      'cotiza_fob_seg_dol
                        rs_aux3!venta_descuento_dol = Ado_datosE.Recordset!cotiza_precio_fob_me               'EUR
                        rs_aux3!venta_precio_total_dol = Ado_datosE.Recordset!cotiza_precio_fob_dol         'cotiza_fob_seg_dol
                        rs_aux3!concepto_venta = "Codigo: " + Ado_datosE.Recordset!bien_codigo + " Modelo: " + Ado_datosE.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datosE.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                        'ok
                        rs_aux3!grupo_codigo = "40000"
                        rs_aux3!subgrupo_codigo = "43000"
                        rs_aux3!par_codigo = "43340"
                        'ok
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo = Ado_datosE.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo1 = Ado_datosE.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                        'rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
'                        rs_aux3!modelo_elegido_h = "N"
'                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!pais_codigo = Ado_datosE.Recordset!pais_codigo
                        'rs_aux3!estado_codigo = "REG"
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                    Wend
                Else
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datosE.Recordset!cotiza_fob_seg_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                    'db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datosE.Recordset!cotiza_fob_seg_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_bs = " & Ado_datosE.Recordset!cotiza_precio_fob_bs & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set venta_precio_unitario_dol = " & Ado_datosE.Recordset!cotiza_precio_fob_dol & " Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                    db.Execute "Update ao_ventas_detalle Set usr_codigo  = '" & glusuario & "' Where venta_codigo = " & var_cod & " AND cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  "
                End If
'                    If Left(rs_aux3!bien_codigo, 4) = "AO36" Or Left(rs_aux3!bien_codigo, 4) = "36NO" Then
'                    Else
'                    End If
'                'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                rs_aux3!venta_precio_unitario_bs = Ado_datos.Recordset!cotiza_fob_seg_bs    'cotiza_precio_fob_bs
'                rs_aux3!venta_descuento_bs = 0
'                rs_aux3!venta_precio_total_bs = Ado_datos.Recordset!cotiza_fob_seg_bs
'                rs_aux3!venta_precio_unitario_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
'                rs_aux3!venta_descuento_dol = 0
'                rs_aux3!venta_precio_total_dol = Ado_datos.Recordset!cotiza_fob_seg_dol
'                rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
'                rs_aux3!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h
'                rs_aux3!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x
'                rs_aux3!fecha_registro = Date
'                rs_aux3!usr_codigo = glusuario
'                rs_aux3.Update
            Else
                'VAR_AUX = rs_aux3.RecordCount + 1
                VAR_EQP = 0
                While Ado_datosE.Recordset!cotiza_cantidad > VAR_EQP
                    VAR_EQP = VAR_EQP + 1
                    rs_aux3.AddNew
                    rs_aux3!ges_gestion = Year(Date)
                    rs_aux3!venta_codigo = var_cod
                    rs_aux3!venta_codigo_det = rs_aux3.RecordCount      'Ado_datos.Recordset!cotiza_codigo      'VAR_AUX
                    rs_aux3!cotiza_codigo = Ado_datosE.Recordset!cotiza_codigo          'VAR_AUX
                    rs_aux3!bien_codigo = "NA" + Trim(Str(VAR_EQP))      'Ado_datos.Recordset!bien_codigo
                    rs_aux3!venta_det_cantidad = 1      'Ado_datos.Recordset!cotiza_cantidad
                    rs_aux3!venta_precio_unitario_bs = Ado_datosE.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_descuento_bs = Ado_datosE.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_bs = Ado_datosE.Recordset!cotiza_fob_seg_bs
                    rs_aux3!venta_precio_unitario_dol = Ado_datosE.Recordset!cotiza_fob_seg_dol
                    rs_aux3!venta_descuento_dol = Ado_datosE.Recordset!cotiza_precio_fob_me               'EUR
                    rs_aux3!venta_precio_total_dol = Ado_datosE.Recordset!cotiza_fob_seg_dol
                    rs_aux3!concepto_venta = "Codigo: " + Ado_datosE.Recordset!bien_codigo + " Modelo: " + Ado_datosE.Recordset!modelo_codigo     'aw_p_ao_solicitud_cotiza_datos.dtc_desc21.Text           '+ " - " + Ado_datosE.Recordset!bien_codigo  '"PAGO POR VENTAS NUEVAS"
                    'ok
                    rs_aux3!grupo_codigo = "40000"
                    rs_aux3!subgrupo_codigo = "43000"
                    rs_aux3!par_codigo = "43340"
                    'ok
                    rs_aux3!tipo_descuento = 0
                    rs_aux3!almacen_codigo = 0
                    rs_aux3!modelo_codigo = Ado_datosE.Recordset!modelo_codigo
                    rs_aux3!modelo_codigo1 = Ado_datosE.Recordset!modelo_codigo
                    rs_aux3!modelo_codigo_h = "S/M"    'Ado_datos.Recordset!modelo_codigo_h
                    rs_aux3!modelo_codigo_x = IIf(IsNull(Ado_datos0.Recordset!beneficiario_codigo), "0", Ado_datos0.Recordset!beneficiario_codigo)
                    'rs_aux3!modelo_codigo_x = "S/M"    'Ado_datos.Recordset!modelo_codigo_x
                    rs_aux3!modelo_elegido = "N"
'                    rs_aux3!modelo_elegido_h = "N"
'                    rs_aux3!modelo_elegido_x = "N"
                    rs_aux3!pais_codigo = Ado_datos.Recordset!pais_codigo
                    'rs_aux3!estado_codigo = "REG"
                    rs_aux3!fecha_registro = Date
                    rs_aux3!usr_codigo = glusuario
                    rs_aux3.Update
                Wend
            End If
         'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
        If VAR_SW2 = 0 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datosE.Recordset!doc_codigo & "'  "
            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                rs_datosA!doc_numero = rs_aux2!correl_doc
                'Txt_campo1.Caption = rs_aux2!correl_doc
                rs_aux2.Update
            End If
            'rs_datos!doc_numero = Txt_campo1.Caption
            'REVISAR !!! JQA 2014_07_08
            'VAR_ARCH = RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
            VAR_ARCH = "COM_" + RTrim(RTrim(rs_datosA!doc_codigo) + "-") + LTrim(Str(rs_datosA!doc_numero))
            Ado_datosE.Recordset!archivo_respaldo = VAR_ARCH + ".PDF"
            Ado_datosE.Recordset!archivo_respaldo_cargado = "N"
            'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
'            Set rs_aux2 = New ADODB.Recordset
'            If rs_aux2.State = 1 Then rs_aux2.Close
'            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos1.Recordset!doc_codigo2 & "'  "
'            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'            If rs_aux2.RecordCount > 0 Then
'                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'                rs_datosA!doc_numero2 = rs_aux2!correl_doc
'                rs_aux2.Update
'            End If
'            VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos1!doc_codigo2) + "-") + LTrim(Str(rs_datos1!doc_numero2))
'            rs_datosA!archivo_respaldo = VAR_ARCH2 + ".PDF"
'            rs_datosA!archivo_respaldo_cargado = "N"
        End If
        Ado_datosE.Recordset!estado_codigo_verif = "APR"
        Ado_datosE.Recordset!fecha_registro = Date
        Ado_datosE.Recordset!usr_codigo = glusuario
        Ado_datosE.Recordset.UpdateBatch adAffectAll
        
        Ado_datosE.Recordset.MoveNext
        Wend
      End If
      'ACTUALIZA PROVEEDOR ESPAÑA
      'db.Execute "update ao_solicitud_cotiza_modelo set beneficiario_codigo = 'ES-A-41043449' where pais_codigo = 'ESP' and unidad_codigo = '" & Ado_datosE.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datosE.Recordset!solicitud_codigo & " "
      db.Execute "UPDATE ao_solicitud_cotiza_modelo set ao_solicitud_cotiza_modelo.beneficiario_codigo  = gc_beneficiario.beneficiario_codigo FROM ao_solicitud_cotiza_modelo INNER JOIN gc_beneficiario ON ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.pais_codigo  AND ao_solicitud_cotiza_modelo.pais_codigo = gc_beneficiario.depto_sigla "
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub BtnBuscar_Click()
    Set ClBuscaGrid = New ClBuscaEnGridExterno
    Set ClBuscaGrid.Conexión = db
    ClBuscaGrid.EsTdbGrid = False
    Set ClBuscaGrid.GridTrabajo = dg_datos0
    ClBuscaGrid.QueryUtilizado = queryinicial
    Set ClBuscaGrid.RecordsetTrabajo = rs_datos0
    'ClBuscaGrid.CamposVisibles = "11010011"
    ClBuscaGrid.Ejecutar
End Sub

Private Sub BtnGrabar_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    GlCotiza = Ado_datos0.Recordset!cotiza_codigo       'txt_codigo1.Caption
    GlUnidad = Ado_datos0.Recordset!unidad_codigo
    Select Case SSTab1.Tab
        Case 0
            marca1 = Ado_datos.Recordset.Bookmark
            If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
                'FraNavega.Enabled = False
                'FraDet1E.Visible = True
                'FraDet1.Visible = False
                FraDet1.Enabled = False
                dg_det1.Enabled = False
                dg_det1.AllowUpdate = False
                VARCTRL = 1
                VAR_CONTI = "AMERICA"
                db.Execute "update ao_solicitud_costos set costo_monto = costo_monto_usd * " & GlTipoCambioOficial & " where unidad_codigo = '" & GlUnidad & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & " AND costo_monto_usd > 0  "
            End If
        Case 1
            marca1 = Ado_datosA.Recordset.Bookmark
            If rs_datosA.RecordCount > 0 And rs_datosA!estado_codigo = "REG" Then
                'FraNavegaA.Enabled = False
                'FraDet1E.Visible = True
                'FraDet1.Visible = False
                FraDet1.Enabled = False
                dg_det1.Enabled = False
                dg_det1.AllowUpdate = False
                VARCTRL = 3
                VAR_CONTI = "ASIA"
                db.Execute "update ao_solicitud_costos set costo_monto = costo_monto_usd * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & " AND costo_monto_usd > 0  "
            End If
        Case 2
            marca1 = Ado_datosE.Recordset.Bookmark
            If rs_datosE.RecordCount > 0 And rs_datosE!estado_codigo = "REG" Then
                'FraNavegaE.Enabled = False
                'FraDet1.Visible = True
                'FraDet1E.Enabled = False
                dg_det1E.Enabled = False
                dg_det1E.AllowUpdate = False
                VARCTRL = 2
                VAR_CONTI = "EUROPA"
                '
                db.Execute "update ao_solicitud_costos set costo_monto = costo_monto2 * " & GlTipoCambioEuro & " where unidad_codigo = '" & txt_campo1.Text & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & " AND costo_monto2 > 0  "
                db.Execute "update ao_solicitud_costos set costo_monto_usd= costo_monto / " & GlTipoCambioOficial & " where unidad_codigo = '" & txt_campo1.Text & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & " AND costo_monto > 0  "
            End If
    End Select
    'VAR_CONTI = Ado_datos.Recordset!pais_continente
    swnuevo = 0
    fraOpciones.Visible = True
    Fra_datos2.Enabled = True
    
    FraNavega0.Enabled = True
    'FrmABMDet.Enabled = False
    SSTab1.Enabled = True
    BtnGrabar.Visible = False
    BtnModDetalle.Visible = True
    BtnAddDetalle2.Visible = True
    
    db.Execute "update ao_solicitud_cotiza_venta set ao_solicitud_cotiza_venta.cotiza_gasto_local_dol = av_solicitud_costo1_acum.costo_montoDol FROM ao_solicitud_cotiza_venta INNER JOIN av_solicitud_costo1_acum " & _
    " ON ao_solicitud_cotiza_venta.unidad_codigo= av_solicitud_costo1_acum.unidad_codigo AND ao_solicitud_cotiza_venta.solicitud_codigo= av_solicitud_costo1_acum.solicitud_codigo AND ao_solicitud_cotiza_venta.pais_continente= av_solicitud_costo1_acum.pais_continente AND ao_solicitud_cotiza_venta.cotiza_codigo = av_solicitud_costo1_acum.cotiza_codigo " & _
    " where ao_solicitud_cotiza_venta.unidad_codigo = '" & parametro & "' and ao_solicitud_cotiza_venta.solicitud_codigo = " & GlSolicitud & " and ao_solicitud_cotiza_venta.pais_continente = '" & VAR_CONTI & "' and ao_solicitud_cotiza_venta.cotiza_codigo = " & txt_codigo1.Caption & " "
    
    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = cotiza_gasto_local_dol * " & CDbl(GlTipoCambioOficial) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
    
    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = cotiza_gasto_local_dol + cotiza_precio_cif_dol   where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = cotiza_precio_total_dol * " & CDbl(GlTipoCambioOficial) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
End Sub

Private Sub BtnImprimirE_Click()
If Ado_datosE.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    'CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_cotizacion_equipos.rpt"
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R222_ar_cotiza_venta_cliente_europa.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datosE.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datosE.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datosE.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datosE.Recordset!EDIF_CODIGO
    CR01.StoredProcParam(4) = Me.Ado_datosE.Recordset!cotiza_codigo
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized

End Sub

Private Sub BtnModificar0_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    Select Case Ado_datos0.Recordset!pais_continente_cot
        Case "AMERICA"
            Call BtnModificar_Click
        Case "ASIA"
            Call BtnModificarA_Click
        Case "EUROPA"
            Call BtnModificarE_Click
        Case Else
            MsgBox "Verifique el Continente en la etapa anterior PARAMETROS DE CALCULO y vuelva a intentar ... ", vbCritical + vbExclamation, "Validación de datos"
    End Select

'Select Case Ado_datos0.Recordset!pais_codigo
'    Case "BRL"
'            Call BtnModificar_Click
'        Case "CHN"
'            Call BtnModificarA_Click
'        Case "ESP"
'            Call BtnModificarE_Click
'        Case Else
'            MsgBox "Verifique el Continente en la etapa anterior PARAMETROS DE CALCULO y vuelva a intentar ... ", vbCritical + vbExclamation, "Validación de datos"
'    End Select
End Sub

'Private Sub BtnCancelar_Click()
'  On Error Resume Next
'   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
'   If sino = vbYes Then
'        rs_datos.CancelUpdate
'        Call ABRIR_TABLA
'        rs_datos.MoveFirst
'        'mbDataChanged = False
'        FraModeloCosto.Visible = False
'        FraGrabarCancelar.Visible = False
'        Fra_datos2.Enabled = False
'        fraOpciones2.Visible = True
'        fraOpciones1.Visible = True
'        FrmABMDet.Visible = True
'        FraDet1.Enabled = True
'        dg_datos.Enabled = True
'        dg_datos1.Enabled = True
'        VAR_SW = ""
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = True
'        SSTab1.TabEnabled(2) = True
'    End If
'End Sub

'Private Sub BtnCancelarA_Click()
'  On Error Resume Next
'   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
'   If sino = vbYes Then
'        rs_datosA.CancelUpdate
'        Call ABRIR_TABLA
'        rs_datosA.MoveFirst
'        'mbDataChanged = False
''        Fra_datos.Enabled = False
'        FraModeloCostoA.Visible = False
''        FraGrabarCancelar.Visible = False
'        Fra_datos2.Enabled = False
'        fraOpciones2A.Visible = True
'        fraOpciones1A.Visible = True
'        FrmABMDet.Visible = True
'        FraDet1.Enabled = True
'        dg_datosA.Enabled = True
'        VAR_SW = ""
'        SSTab1.Tab = 1
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = True
'        SSTab1.TabEnabled(2) = True
'    End If
'End Sub

'Private Sub BtnGrabar_Click()
'  On Error GoTo UpdateErr
'  VAR_VAL = "OK"
'  Call valida_campos
'  If VAR_VAL = "OK" Then
'    VAR_CONTI = "AMERICA"
'    Set rs_datos10 = New ADODB.Recordset
'    If rs_datos10.State = 1 Then rs_datos10.Close
'    rs_datos10.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  ", db, adOpenKeyset, adLockBatchOptimistic
'    'Set Ado_datos3.Recordset = rs_datos6
'    If rs_datos10.RecordCount > 0 Then
''       sino = MsgBox("SI (Graba todos los Registros) - NO (Graba SOLO el Registro Activo) ... ", vbYesNo + vbQuestion, "Atención")
''       If sino = vbYes Then
'           'TODOS LOS REGISTROS - 'Clonar todos los registros
''           Ado_datos.Recordset.MoveFirst
''           While Not Ado_datos.Recordset.EOF
''             Set Ado_datos.Recordset = rs_datos10
''             txt_codigo1.Caption = Ado_datos.Recordset!cotiza_codigo
''             If Val(txt_codigo1.Caption) = 1 Then
''                 Ado_datos.Recordset!cotiza_dec = cmd_dec.Text
''                 Ado_datos.Recordset!tipo_moneda = cmd_moneda.Text
''                 If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
''                    txt_tdc.Text = GlTipoCambioOficial
''                 End If
''                 Ado_datos.Recordset!cotiza_tdc_bol = txt_tdc.Text
''                 Ado_datos.Recordset!costo_monto = txt_montobase.Text
''                 Ado_datos.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me1 = "", "0", Round(txt_fob_me1, Val(cmd_dec)))
''                 Ado_datos.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))  'Txt_campo6.Text
''                 Ado_datos.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me1 = "", "0", Round(txt_dcto_me1, Val(cmd_dec)))
''                 Ado_datos.Recordset!cotiza_precio_dcto_bs = Round(CDbl(txt_dcto_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
''                 Ado_datos.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me1 = "", "0", Round(txt_seguro_me1, Val(cmd_dec)))
''                 Ado_datos.Recordset!cotiza_precio_seg_bs = Round(CDbl(txt_seguro_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
''
''                 Ado_datos.Recordset!cotiza_fob_seg_dol = Round(CDbl(txt_fob_me1) - CDbl(txt_dcto_me1) + CDbl(txt_seguro_me1), Val(cmd_dec))
''                 Ado_datos.Recordset!cotiza_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
''
''                 Ado_datos.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me1 = "", "0", Round(txt_fletefrontera_me1, Val(cmd_dec)))
''                 Ado_datos.Recordset!cotiza_precio_flete_bs = Round(CDbl(txt_fletefrontera_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
''
''                 Ado_datos.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_me1) - CDbl(txt_dcto_me1.Text) + CDbl(txt_seguro_me1.Text) + CDbl(txt_fletefrontera_me1.Text), Val(cmd_dec))
''                 Ado_datos.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
''
''                 Ado_datos.Recordset!fecha_registro = Date     'no cambia
''                 Ado_datos.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
''                 Ado_datos.Recordset.Update    'Batch 'adAffectAll
''                 db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''             Else
'                'CLONAR TODOS LOS REGISTROS
''                Set rs_aux7 = New ADODB.Recordset
''                If rs_aux7.State = 1 Then rs_aux7.Close
''                rs_aux7.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = 1  ", db, adOpenStatic
''                'Set Ado_datos1.Recordset = rs_aux7
''                If rs_aux7.RecordCount > 0 Then
''                    'WWWWWWWWWWWWWWWWWWWWWW
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_dec = " & rs_aux7!cotiza_dec & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set tipo_moneda= '" & rs_aux7!tipo_moneda & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_tdc_bol = " & rs_aux7!cotiza_tdc_bol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set costo_monto = " & rs_aux7!costo_monto & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_dol = " & rs_aux7!cotiza_precio_fob_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_bs = " & Round(CDbl(rs_aux7!cotiza_precio_fob_bs), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_dol = " & rs_aux7!cotiza_precio_dcto_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_bs = " & CDbl(rs_aux7!cotiza_precio_dcto_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_dol = " & rs_aux7!cotiza_precio_seg_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_bs = " & CDbl(rs_aux7!cotiza_precio_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_dol = " & CDbl(rs_aux7!cotiza_fob_seg_dol) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_bs = " & CDbl(rs_aux7!cotiza_fob_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_dol = " & rs_aux7!cotiza_precio_flete_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_bs = " & CDbl(rs_aux7!cotiza_precio_flete_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_dol = " & Round(CDbl(rs_aux7!cotiza_precio_cif_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_bs = " & Round(rs_aux7!cotiza_precio_cif_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(rs_aux7!cotiza_precio_total_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cli), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(rs_aux7!cotiza_precio_total_bs_cli, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cge), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(rs_aux7!cotiza_precio_total_bs_cge, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(CDbl(rs_aux7!cotiza_gasto_local_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(rs_aux7!cotiza_gasto_local_bs, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IT_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & (rs_aux7!cotiza_saldo_local_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IVA_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & rs_aux7!cotiza_saldo_local_IVA_bs & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IT_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & rs_aux7!cotiza_saldo_cge_IT_bs & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IVA_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & rs_aux7!cotiza_saldo_cge_IVA_bs & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_tac_billing_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & rs_aux7!cotiza_saldo_tac_billing_bs & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
''                    'WWWWWWWWWWWWWWWWWWWWWW
''                End If
''             End If
''             'rs_datos10!cotiza_precio_seg_dol = IIf(txt_seguro_me1 = "", "0", txt_seguro_me1)
''             'rs_datos1!cotiza_precio_seg_bs = CDbl(txt_seguro_me1) * CDbl(GlTipoCambioOficial)
''
''    '         'rs_datos!Foto = Date
''    '         'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
''    '         'rs_datos!archivo_foto_cargado = "N"
''    '         'hora_registro
''             'MsgBox Str(rs_datos10.RecordCount)
''
''             'GRABA COSTOS
''             Set rs_aux5 = New ADODB.Recordset
''             If rs_aux5.State = 1 Then rs_aux5.Close
''             rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
''             If rs_aux5.RecordCount = 0 Then
''                Call GRABA_COSTOS
''             Else
''                sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
''
''                If sino = vbYes Then
''                    'OJO BORRAR ao_solicitud_costos
''                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
''                    'corrprog = 0
''                    Call GRABA_COSTOS
''                Else
''                    Set rs_aux6 = New ADODB.Recordset
''                    If rs_aux6.State = 1 Then rs_aux6.Close
''                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux6.RecordCount > 0 Then
''                        VAR_NAC = rs_aux6!costo_monto_usd
''                    End If
''                    Set rs_aux6 = New ADODB.Recordset
''                    If rs_aux6.State = 1 Then rs_aux6.Close
''                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux6.RecordCount > 0 Then
''                        VAR_ALM = rs_aux6!costo_monto_usd
''                    End If
''                    Set rs_aux6 = New ADODB.Recordset
''                    If rs_aux6.State = 1 Then rs_aux6.Close
''                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux6.RecordCount > 0 Then
''                        VAR_AGE = rs_aux6!costo_monto_usd
''                    End If
''                    Set rs_aux6 = New ADODB.Recordset
''                    If rs_aux6.State = 1 Then rs_aux6.Close
''                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1' and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux6.RecordCount > 0 Then
''                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
''                    End If
''                    Set rs_aux6 = New ADODB.Recordset
''                    If rs_aux6.State = 1 Then rs_aux6.Close
''                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
''                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux6.RecordCount > 0 Then
''                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
''                    End If
''                End If
''
''             End If
''             If Ado_datos.Recordset!pais_continente = "AMERICA" And Val(txt_codigo1.Caption) = 1 Then
''                    Set rs_aux4 = New ADODB.Recordset
''                    If rs_aux4.State = 1 Then rs_aux4.Close
''                    rs_aux4.Open "select sum(costo_monto) as totbs, sum(costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "'  AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
''                    If rs_aux4.RecordCount > 0 Then
''                        SUBTOTD = Round(rs_aux4!totdl + Ado_datos.Recordset!cotiza_precio_cif_dol - Ado_datos.Recordset!cotiza_precio_flete_dol, Val(cmd_dec))
''                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    End If
''                    'Importaion Cliente
''                    VAR_LOCAL = Round(rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE, Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(VAR_LOCAL, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(VAR_LOCAL * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    txt_local_IT_me.Text = Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    txt_local_IVA_me = Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec))
''
''                    VAR_DOLCLI2 = Round(SUBTOTD + CDbl(txt_local_IT_me) + CDbl(txt_local_IVA_me), Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(VAR_DOLCLI2, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(VAR_DOLCLI2 * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''
''                    VAR_DOLCLI = Round(rs_aux4!totdl + Ado_datos.Recordset!cotiza_precio_cif_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
''                    'VAR_BSCLI = rs_aux4!totbs + Ado_datos.Recordset!cotiza_precio_cif_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
''                    VAR_BSCLI = Round(VAR_DOLCLI * GlTipoCambioOficial, Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''
''                    VAR_SUBD = Round(SUBTOTD - Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
''                    VAR_SUBB = Round(VAR_SUBD * GlTipoCambioOficial, Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    txt_cge_IT_me = Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec))
''
''                    'IMPORTACION CGE
''                    txt_cge_IVA_me = Round((VAR_SUBD * CDbl(txt_cge_IVA_bs)) - ((Ado_datos.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_me, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    txt_tac_billing_me = Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec))
''
''                    VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_me) + CDbl(txt_cge_IVA_me) + CDbl(txt_tac_billing_me), Val(cmd_dec))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''             End If
''           Ado_datos.Recordset.MoveNext
''           Wend
''       Else
'           '- SOLO EL REGISTRO ACTIVO
'             Ado_datos.Recordset!cotiza_dec = IIf(cmd_dec.Text = "", "2", cmd_dec.Text)
'             Ado_datos.Recordset!tipo_moneda = IIf(cmd_moneda.Text = "", "BOB", cmd_moneda.Text)
'             If txt_tdc.Text = "0" Or txt_tdc.Text = "" Then
'                txt_tdc.Text = GlTipoCambioOficial
'             End If
'             Ado_datos.Recordset!cotiza_tdc_bol = txt_tdc.Text
'             Ado_datos.Recordset!costo_monto = IIf(txt_montobase.Text = "", "0", Round(txt_montobase.Text, Val(cmd_dec)))
'             Ado_datos.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me1 = "", "0", Round(txt_fob_me1, Val(cmd_dec)))
'             Ado_datos.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))  'Txt_campo6.Text
'             Ado_datos.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me1 = "", "0", Round(txt_dcto_me1, Val(cmd_dec)))
'             Ado_datos.Recordset!cotiza_precio_dcto_bs = Round(CDbl(txt_dcto_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'             Ado_datos.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me1 = "", "0", Round(txt_seguro_me1, Val(cmd_dec)))
'             Ado_datos.Recordset!cotiza_precio_seg_bs = Round(CDbl(txt_seguro_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'             Ado_datos.Recordset!cotiza_fob_seg_dol = Round(CDbl(txt_fob_me1) - CDbl(txt_dcto_me1) + CDbl(txt_seguro_me1), Val(cmd_dec))
'             Ado_datos.Recordset!cotiza_fob_seg_bs = Round(CDbl(txt_fob_seg_dol) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'             Ado_datos.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me1 = "", "0", Round(txt_fletefrontera_me1, Val(cmd_dec)))
'             Ado_datos.Recordset!cotiza_precio_flete_bs = Round(CDbl(txt_fletefrontera_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec))
'
'             Ado_datos.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_me1) - CDbl(txt_dcto_me1.Text) + CDbl(txt_seguro_me1.Text) + CDbl(txt_fletefrontera_me1.Text), Val(cmd_dec))
'             Ado_datos.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_me1) * CDbl(GlTipoCambioOficial), Val(cmd_dec)) '
'    '         'rs_datos!Foto = Date
'    '         'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
'    '         'rs_datos!archivo_foto_cargado = "N"
'    '         'hora_registro
'             Ado_datos.Recordset!fecha_registro = Date     'no cambia
'             Ado_datos.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'             Ado_datos.Recordset.Update    'Batch 'adAffectAll
'             db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'NO' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'             'GRABA COSTOS
'             Set rs_aux5 = New ADODB.Recordset
'             If rs_aux5.State = 1 Then rs_aux5.Close
'             rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
'             If rs_aux5.RecordCount = 0 Then
'                Call GRABA_COSTOS
'             Else
'                sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
'                If sino = vbYes Then
'                    'OJO BORRAR ao_solicitud_costos
'                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
'                    'corrprog = 0
'                    Call GRABA_COSTOS
'                Else
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_NAC = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_ALM = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_AGE = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1' and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = '1'  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
'                    'rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                End If
'
'             End If
'             If Ado_datos.Recordset!pais_continente = "AMERICA" And Val(txt_codigo1.Caption) = 1 Then
'                    Set rs_aux4 = New ADODB.Recordset
'                    If rs_aux4.State = 1 Then rs_aux4.Close
'                    rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "'  AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux4.RecordCount > 0 Then
'                        SUBTOTD = Round(rs_aux4!totdl + Ado_datos.Recordset!cotiza_precio_cif_dol - Ado_datos.Recordset!cotiza_precio_flete_dol, Val(cmd_dec))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    End If
'                    'Importaion Cliente
'                    VAR_LOCAL = Round(rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE, Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(VAR_LOCAL, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(VAR_LOCAL * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bs.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    txt_local_IT_me.Text = Round(VAR_LOCAL * CDbl(txt_local_IT_bs.Text), Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bs.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    txt_local_IVA_me = Round(VAR_LOCAL * CDbl(txt_local_IVA_bs.Text), Val(cmd_dec))
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = ROUND(cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol, Val(cmd_dec)) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = ROUND(cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & ", Val(cmd_dec)) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'                    VAR_DOLCLI2 = Round(SUBTOTD + CDbl(txt_local_IT_me) + CDbl(txt_local_IVA_me), Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(VAR_DOLCLI2, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(VAR_DOLCLI2 * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                    VAR_DOLCLI = Round(rs_aux4!totdl + Ado_datos.Recordset!cotiza_precio_cif_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
'                    'VAR_BSCLI = rs_aux4!totbs + Ado_datos.Recordset!cotiza_precio_cif_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
'                    VAR_BSCLI = Round(VAR_DOLCLI * GlTipoCambioOficial, Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                    'VAR_SUBD = IIf(IsNull(Ado_datos.Recordset!cotiza_precio_total_dol), SUBTOTD, Ado_datos.Recordset!cotiza_precio_total_dol) - Ado_datos.Recordset!cotiza_precio_seg_dol
'                    VAR_SUBD = Round(SUBTOTD - Ado_datos.Recordset!cotiza_precio_seg_dol, Val(cmd_dec))
'                    VAR_SUBB = Round(VAR_SUBD * GlTipoCambioOficial, Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    txt_cge_IT_me = Round(VAR_SUBD * CDbl(txt_cge_IT_bs), Val(cmd_dec))
'                    'IMPORTACION CGE
'                    txt_cge_IVA_me = Round((VAR_SUBD * CDbl(txt_cge_IVA_bs)) - ((Ado_datos.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_me, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bs) & ") -((cotiza_precio_cif_dol * 0.1498) * " & CDbl(dtc_desc15) & ")-((" & CDbl(VAR_AGE) & " * 0.13)* " & CDbl(dtc_desc15) & ")  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    txt_tac_billing_me = Round((VAR_SUBD * CDbl(txt_tac_billing_bs)), Val(cmd_dec))
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_SUBD + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol), Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_me) + CDbl(txt_cge_IVA_me) + CDbl(txt_tac_billing_me), Val(cmd_dec))
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_dec)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'             End If
''       End If
'    End If
'     SSTab1.Tab = 0
'     Call ABRIR_TABLA
'     If Ado_datos.Recordset!pais_continente = "AMERICA" Then
'         SSTab1.TabEnabled(0) = True
'     Else
'        SSTab1.TabEnabled(0) = False
'     End If
'     If Ado_datos.Recordset!pais_continente = "ASIA" Then
'        SSTab1.TabEnabled(1) = True
'     Else
'        SSTab1.TabEnabled(1) = False
'     End If
'     If Ado_datos.Recordset!pais_continente = "EUROPA" Then
'        SSTab1.TabEnabled(2) = True
'     Else
'        SSTab1.TabEnabled(2) = False
'     End If
'
''     Call ABRIR_TABLA
'     'Ado_datos.Recordset.MoveLast
''     mbDataChanged = False
''        Fra_datos.Enabled = False
'        FraModeloCosto.Visible = False
'        FraGrabarCancelar.Visible = False
'        Fra_datos2.Enabled = False
'        fraOpciones2.Visible = True
'        fraOpciones1.Visible = True
'        FrmABMDet.Visible = True
'        FraDet1.Enabled = True
'        dg_datos.Enabled = True
'        dg_datos1.Enabled = True
'        VAR_SW = ""
'
''     dtc_codigo9.Enabled = True
'    'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            'VAR_VAL,
''        VAR_NO2 = VAR_NO2 + Val(dtc_desc11.Text) - 1
''        VAR_NO3 = "36NO-" + Trim(Str(VAR_NO2))
''        If rs_datos!h_nro_total_equipos > 1 Then
''            'If Right(VAR_NO3, 1) = 0 Then
''                rs_datos!unidad_codigo_ant = VAR_NO1 + "-" + Right(VAR_NO3, 2)
''            'Else
''            '    rs_datos!unidad_codigo_ant = VAR_NO1 + "/" + Right(VAR_NO3, 1)
''            'End If
''        Else
''            rs_datos!unidad_codigo_ant = Txt_campo11    'VAR_NO1
''        End If
'        'WC2015
''        db.Execute "Update ao_solicitud Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
''        db.Execute "Update ao_solicitud_calculo_trafico Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
''        db.Execute "Update ao_negociacion_cabecera Set unidad_codigo_ant = '" & rs_datos!unidad_codigo_ant & "' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and negocia_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
''     Call GRABA_COSTOS
'
'    'WWWWWWWWWWWWWWWWWWWWWWWWWWWW
'  End If
''  dtc_desc1.Visible = True
''  lbl_aux1.Visible = False
'  Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'
'End Sub

'Private Sub GRABA_COSTOS()
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    'VAR_CONTI = "AMERICA"
'    If VAR_CONTI = "AMERICA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipo= 'B' ", db, adOpenStatic
'    End If
'    If VAR_CONTI = "ASIA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoA= 'B' ", db, adOpenStatic
'    End If
'    If VAR_CONTI = "EUROPA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoE= 'B' ", db, adOpenStatic
'    End If
'    Set Ado_datos3.Recordset = rs_datos6
'    If Ado_datos3.Recordset.RecordCount > 0 Then
'        Ado_datos3.Recordset.MoveFirst
'        While Not Ado_datos3.Recordset.EOF
'            'codigo_costo
'            'costo_descripcion
'            'costo_monto
'            'costo_porcentaje
'            'costo_tipo
'            Set rs_aux5 = New ADODB.Recordset
'            If rs_aux5.State = 1 Then rs_aux5.Close
'            rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and cotiza_codigo = " & CDbl(txt_codigo1) & " ", db, adOpenKeyset, adLockOptimistic      'AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "
'            'If rs_aux5.RecordCount = 0 Then
'                rs_aux5.AddNew
'                rs_aux5!ges_gestion = Year(Date)
'                rs_aux5!unidad_codigo = parametro           'Txt_campo1.Caption
'                rs_aux5!solicitud_codigo = GlSolicitud      'Ado_datos.Recordset!solicitud_codigo
'                rs_aux5!edif_codigo = GlEdificio            'Ado_datos.Recordset!edif_codigo
'                rs_aux5!cotiza_codigo = txt_codigo1         'Ado_datos.Recordset!cotiza_codigo
'
'                rs_aux5!pais_continente = VAR_CONTI
'                rs_aux5!estado_codigo = "REG"
'                rs_aux5!codigo_costo = Ado_datos3.Recordset!codigo_costo
'                rs_aux5!costo_porcentaje = Ado_datos3.Recordset!costo_porcentaje
'                If Ado_datos3.Recordset!costo_porcentaje > 0 Then
'                    If VAR_CONTI = "AMERICA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datos.Recordset!cotiza_precio_fob_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
'                        Else
'                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datos.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            rs_aux5!costo_monto = Round(CDbl(rs_aux5!costo_monto_usd * CDbl(GlTipoCambioOficial)), CDbl(cmd_dec))
'                        End If
'                    End If
'                    If VAR_CONTI = "ASIA" Then
'                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
'                            If IsNull(Ado_datosA.Recordset!cotiza_precio_spread_bs) Then
'                                rs_aux5!costo_monto = Round(CDbl((Ado_datos1A.Recordset!cotiza_precio_fob_bs + Ado_datos1A.Recordset!cotiza_precio_spread_bs) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl((Ado_datos1A.Recordset!cotiza_precio_fob_dol + Ado_datos1A.Recordset!cotiza_precio_spread_dol) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            Else
'                                rs_aux5!costo_monto = Round(CDbl((Ado_datosA.Recordset!cotiza_precio_fob_bs + Ado_datosA.Recordset!cotiza_precio_spread_bs) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl((Ado_datosA.Recordset!cotiza_precio_fob_dol + Ado_datosA.Recordset!cotiza_precio_spread_dol) * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            End If
'                        Else
'                            'rs_aux5!costo_monto = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_cif_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            'rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
'                            If IsNull(Ado_datosA.Recordset!cotiza_precio_base_bs) Then
'                                rs_aux5!costo_monto = Round(CDbl(Ado_datos1A.Recordset!cotiza_precio_base_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(Ado_datos1A.Recordset!cotiza_precio_base_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            Else
'                                rs_aux5!costo_monto = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_base_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosA.Recordset!cotiza_precio_base_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_decA))
'                            End If
'                        End If
'                    End If
'                    If VAR_CONTI = "EUROPA" Then
''                        If Ado_datos3.Recordset!codigo_costo = 15 Then  ' TRANSFERENCIA BANCARIA
''                            rs_aux5!costo_monto = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_fob_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
''                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_fob_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
''                        Else
''                            rs_aux5!costo_monto = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_cif_bs * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
''                            rs_aux5!costo_monto_usd = Round(CDbl(Ado_datosE.Recordset!cotiza_precio_cif_dol * Ado_datos3.Recordset!costo_porcentaje), CDbl(cmd_dec))
''                        End If
'                    End If
'                    rs_aux5!costo_monto2 = 0    'Round(CDbl(IIf(txt_total_bs1.Text = "", "0", txt_total_bs1.Text)), 2)
'                    rs_aux5!costo_monto_usd2 = 0    'Round(CDbl(txt_total_me1.Text), 2)
'                    rs_aux5!costo_monto3 = 0    'Round(CDbl(IIf(txt_dcto_bs1.Text = "", "0", txt_dcto_bs1.Text)), 2)
'                    rs_aux5!costo_monto_usd3 = 0    'Round(CDbl(txt_dcto_me1.Text), 2)
'                Else
'                    'abrir tabla costos_paradas
'                    Set rs_datos9 = New ADODB.Recordset
'                    If rs_datos9.State = 1 Then rs_datos9.Close
'                    rs_datos9.Open "SELECT * FROM ac_costos_paradas where trafico_num_paradas = " & VAR_PRDA & " ", db, adOpenStatic
'                    Set Ado_datos9.Recordset = rs_datos9
'                    If Ado_datos9.Recordset.RecordCount > 0 Then
'                        If Ado_datos3.Recordset!codigo_costo = 9 Then
'                            If VAR_CONTI = "AMERICA" Then
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_instal_pintura), CDbl(cmd_dec))
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_instal_pintura * GlTipoCambioOficial), CDbl(cmd_dec))
'                            End If
'                            If VAR_CONTI = "ASIA" Then
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_instal_pintura), CDbl(cmd_decA))
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_instal_pintura * GlTipoCambioOficial), CDbl(cmd_decA))
'                            End If
'                        End If
'                        If Ado_datos3.Recordset!codigo_costo = 11 Then
'                            If VAR_CONTI = "AMERICA" Then
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs) * CDbl(Txt_campo5.Text), CDbl(cmd_dec))
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd) * CDbl(Txt_campo5.Text), CDbl(cmd_dec))
'                            End If
'                            If VAR_CONTI = "ASIA" Then
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs) * CDbl(Txt_campo5A.Text), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd) * CDbl(Txt_campo5A.Text), CDbl(cmd_decA))
'                            End If
'                            If VAR_CONTI = "EUROPA" Then
''                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_install_bs), 2) * CDbl(Txt_campo5E.Text)
''                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_install_usd), 2) * CDbl(Txt_campo5E.Text)
'                            End If
'                        End If
'                        If Ado_datos3.Recordset!codigo_costo = 12 Then
'                            If VAR_CONTI = "AMERICA" Then
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_ajuste_bs), CDbl(cmd_dec))
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_ajuste_usd), CDbl(cmd_dec))
'                            End If
'                            If VAR_CONTI = "ASIA" Then
'                                rs_aux5!costo_monto = Round(CDbl(rs_datos9!costo_ajuste_bs), CDbl(cmd_decA))
'                                rs_aux5!costo_monto_usd = Round(CDbl(rs_datos9!costo_ajuste_usd), CDbl(cmd_decA))
'                            End If
'                        End If
'                    End If
'                End If
'                If Ado_datos3.Recordset!codigo_costo = 3 Then   'NACIONALIZACION
'                    VAR_NAC = rs_aux5!costo_monto_usd
'                End If
'                If Ado_datos3.Recordset!codigo_costo = 5 Then   'ALMACENAJE
'                    VAR_ALM = rs_aux5!costo_monto_usd
'                End If
'                If Ado_datos3.Recordset!codigo_costo = 6 Then   'COMISION AGENCIA
'                    VAR_AGE = rs_aux5!costo_monto_usd
'                End If
'                If Ado_datos3.Recordset!codigo_costo = 8 Then   'TOTAL FLETES
'                    VAR_FLE = IIf(IsNull(rs_aux5!costo_monto_usd), "0", rs_aux5!costo_monto_usd)
'                End If
'                If VAR_CONTI = "AMERICA" Then
'                    'VAR_DOLCLI = Ado_datos.Recordset!cotiza_precio_total_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol
'                    'VAR_BSCLI = Ado_datos.Recordset!cotiza_precio_total_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
'                End If
'                If VAR_CONTI = "ASIA" Then
'                    'VAR_DOLCLI = Ado_datos.Recordset!cotiza_precio_total_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol
'                    'VAR_BSCLI = Ado_datos.Recordset!cotiza_precio_total_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
'                End If
'                If VAR_CONTI = "EUROPA" Then
'                    'VAR_DOLCLI = Ado_datos.Recordset!cotiza_precio_total_dol - Ado_datos.Recordset!cotiza_precio_fob_dol - Ado_datos.Recordset!cotiza_precio_seg_dol
'                    'VAR_BSCLI = Ado_datos.Recordset!cotiza_precio_total_bs - Ado_datos.Recordset!cotiza_precio_fob_bs - Ado_datos.Recordset!cotiza_precio_seg_bs
'                End If
'                rs_aux5!costo_observaciones = Trim(Ado_datos3.Recordset!costo_descripcion)
'
'                rs_aux5!fecha_registro = Date
'                'aw_p_ao_negociacion_cabecera.Ado_detalle1.Recordset("hora_registro").Value = Date
'                rs_aux5!usr_codigo = glusuario
'                rs_aux5.Update
'            'End If
'            Ado_datos3.Recordset.MoveNext
'        Wend
'    End If
'End Sub

'Private Sub GRABA_COSTOS_CLON()
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    'VAR_CONTI = "AMERICA"
'    If VAR_CONTI = "AMERICA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipo= 'B' ", db, adOpenStatic
'    End If
'    If VAR_CONTI = "ASIA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoA= 'B' ", db, adOpenStatic
'    End If
'    If VAR_CONTI = "EUROPA" Then
'        rs_datos6.Open "select * from ac_costos_comercializacion where costo_tipoE= 'B' ", db, adOpenStatic
'    End If
'    Set Ado_datos3.Recordset = rs_datos6
'    If Ado_datos3.Recordset.RecordCount > 0 Then
'        Ado_datos3.Recordset.MoveFirst
'        While Not Ado_datos3.Recordset.EOF
'            'codigo_costo
'            'costo_descripcion
'            'costo_monto
'            'costo_porcentaje
'            'costo_tipo
'            Set rs_aux5 = New ADODB.Recordset
'            If rs_aux5.State = 1 Then rs_aux5.Close
'            'rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and cotiza_codigo = " & CDbl(txt_codigo1) & " ", db, adOpenKeyset, adLockOptimistic      'AND cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "
'            rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and codigo_costo = " & Ado_datos3.Recordset!codigo_costo & " AND cotiza_codigo = '1' ", db, adOpenKeyset, adLockOptimistic
'            If rs_aux5.RecordCount > 0 Then
'                'If VAR_CONTI = "AMERICA" Then
'                    'db.Execute "INSERT INTO ao_solicitud_costos (cuenta,subcta1,subcta2, aux1, AUX2, aux3, denominacion_aux1, denominacion_aux2, denominacion_aux3, NombreCta, DebeSaldoIBs, DebeSaldoISus, HaberSaldoIBs, HaberSaldoISus, Cod_Anterior, Status, Verificado, Nom_Aux1, Nom_Aux2, Nom_Aux3) VALUES ('0', '0', '0', '0', '0', '0','-', '-', '-', '', 0, 0, 0, 0, '0', 'N', 'N', '', '', '') "
'                    db.Execute "INSERT INTO ao_solicitud_costos (ges_gestion, unidad_codigo, solicitud_codigo, edif_codigo, cotiza_codigo, pais_continente, estado_codigo, codigo_costo, costo_porcentaje, costo_monto, costo_monto_usd, costo_observaciones, usr_codigo) values ('" & glGestion & "', '" & parametro & "', " & GlSolicitud & ", '" & GlEdificio & "', " & Val(txt_codigo1) & ", '" & VAR_CONTI & "', 'REG', " & rs_aux5!codigo_costo & ", " & rs_aux5!costo_porcentaje & ", " & rs_aux5!costo_monto & ", " & rs_aux5!costo_monto_usd & ", '" & Ado_datos3.Recordset!costo_descripcion & "', '" & rs_aux5!usr_codigo & "')"
'                    '(ges_gestion, unidad_codigo, solicitud_codigo, edif_codigo, cotiza_codigo, pais_continente, estado_codigo, codigo_costo, costo_porcentaje, costo_monto_usd) values
'                    '('" & glGestion & "', '" & parametro & "', " & GlSolicitud & ", '" & GlEdificio & "', " & Val(txt_codigo1) & ", '" & VAR_CONTI & "', 'REG', " & rs_aux5!codigo_costo & ", " & rs_aux5!costo_porcentaje & ", " & rs_aux5!costo_monto_usd & ")"
'                'End If
'                'If VAR_CONTI = "ASIA" Then
'                'End If
'            End If
'            Ado_datos3.Recordset.MoveNext
'        Wend
'    End If
'End Sub

'Private Sub valida_campos()
'  '
'  If (dtc_codigo11 = "") Then
'    MsgBox "Debe registrar Parámetros de Cálculo, Consulte con el Administrador ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fob_me1 = "") Or (txt_fob_me1 = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_seguro_me1 = "") Or (txt_seguro_me1 = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fletefrontera_me1 = "") Or (txt_fletefrontera_me1 = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'End Sub

'Private Sub valida_camposA()
'  '
'  If (dtc_codigo11 = "") Then
'    MsgBox "Debe registrar Parámetros de Cálculo, Consulte con el Administrador ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fob_me1A = "") Or (txt_fob_me1A = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo2A.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_seguro_me1A = "") Or (txt_seguro_me1A = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo4A.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_fletefrontera_me1A = "") Or (txt_fletefrontera_me1A = "0") Then
'    MsgBox "Debe registrar ... " + lbl_campo3A.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'
'  If (txt_tacb1 = "") Then
'    MsgBox "Debe registrar % TAC Billing(Global) ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_tacb2 = "") Then
'    MsgBox "Debe registrar TAC Billing(Global) ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_spread1 = "") Then
'    MsgBox "Debe registrar % Spread Global ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_spread2 = "") Then
'    MsgBox "Debe registrar Spread Global ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (txt_GAC_dol = "") Then
'    MsgBox "Debe registrar GAC ... ", vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'
'End Sub

'Private Sub BtnGrabarA_Click()
''WWWWWWWWWWWWWWWWWWWWWWWW
'  On Error GoTo UpdateErr
'  VAR_VAL = "OK"
'  VAR_CONTI = "ASIA"
'  Call valida_camposA
'  If VAR_VAL = "OK" Then
'    Set rs_datos10 = New ADODB.Recordset
'    If rs_datos10.State = 1 Then rs_datos10.Close
'    rs_datos10.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "'   ", db, adOpenKeyset, adLockOptimistic
'    'Set Ado_datos3.Recordset = rs_datos6
'    If rs_datos10.RecordCount > 0 Then
'       sino = MsgBox("SI (Graba todos los Registros) - NO (Graba SOLO el Registro Activo) ... ", vbYesNo + vbQuestion, "Atención")
'       If sino = vbYes Then
'           'TODOS LOS REGISTROS
'           'Set Ado_datosA.Recordset = rs_datos10
'           'Ado_datosA.Recordset.MoveFirst
'           rs_datos10.MoveFirst
'           'While Not Ado_datosA.Recordset.EOF
'           While Not rs_datos10.EOF
''         'MsgBox "codigo: " + Str(rs_datos10!cotiza_codigo)
'             ''Set Ado_datos1A.Recordset = rs_datos10
'             'txt_codigo1.Caption = Ado_datosA.Recordset!cotiza_codigo
'             txt_codigo1.Caption = rs_datos10!cotiza_codigo
'             If Val(txt_codigo1.Caption) = 1 Then
'                 'WWWWWWWWWWWWWWWW
'                 If txt_tdcA.Text = "0" Or txt_tdcA.Text = "" Then
'                    txt_tdcA.Text = GlTipoCambioOficial
'                 End If
'                 If txt_local_IT_bsA.Text = "" Then
'                    txt_local_IT_bsA.Text = "0.0309"
'                 End If
'                 If txt_local_IVA_bsA.Text = "" Then
'                    txt_local_IVA_bsA.Text = "0.1491"
'                 End If
'                 If txt_cge_IT_bsA.Text = "" Then
'                    txt_cge_IT_bsA = "0.0416"
'                 End If
'                 If txt_cge_IVA_bsA.Text = "" Then
'                    txt_cge_IVA_bsA = "0.151"
'                 End If
'                 If txt_tac_billing_bsA.Text = "" Then
'                    txt_tac_billing_bsA = "0.035"
'                 End If
'                 If txt_GAC_bs = "" Then
'                    txt_GAC_bs = "0.05"
'                 End If
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_dec = " & Val(cmd_decA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set tipo_moneda= '" & cmd_monedaA.Text & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_tdc_bol = " & CDbl(txt_tdcA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'
'                 db.Execute "update ao_solicitud_cotiza_venta set costo_monto = " & CDbl(txt_montobaseA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_dol = " & Round(CDbl(txt_fob_me1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_bs = " & Round(CDbl(txt_fob_bs1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_dol = " & Round(CDbl(txt_dcto_me1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_bs = " & Round(CDbl(txt_dcto_bs1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_dol = " & Round(CDbl(txt_seguro_me1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_bs = " & Round(CDbl(txt_seguro_bs1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_dol = " & Round(CDbl(txt_fob_seg_dolA.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_bs = " & Round(CDbl(txt_fob_seg_bsA.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_dol = " & Round(CDbl(txt_fletefrontera_me1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_bs = " & Round(CDbl(txt_fletefrontera_bs1A.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_dol = " & Round(CDbl(txt_tacb2.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_bs = " & Round(CDbl(txt_tacb1.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_dol  = " & Round(CDbl(txt_spread2.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_bs  = " & Round(CDbl(txt_spread1.Text), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_dol = " & Round(CDbl(txt_cif_me1A), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_bs = " & Round(CDbl(txt_cif_bs1A), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_dol = " & Round(CDbl(txt_GAC_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_bs  = " & Round(CDbl(txt_GAC_bs), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_dol  = " & Round(CDbl(txt_base_imp_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_bs  = " & Round(CDbl(txt_base_imp_bs), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'
'                 'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(CDbl(txt_total_me1A), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(CDbl(txt_total_bs1A), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                 db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    'WWWWWWWWWWWWWWWWWWWWWW
'
'                 'Ado_datosA.Recordset!cotiza_dec = cmd_decA.Text
'                 'Ado_datosA.Recordset!tipo_moneda = cmd_monedaA.Text
'                 'Ado_datosA.Recordset!cotiza_tdc_bol = txt_tdcA.Text
'                 'Ado_datosA.Recordset!costo_monto = CDbl(txt_montobaseA.Text)
'                 'Ado_datosA.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me1A = "", "0", txt_fob_me1A)
'                 'Ado_datosA.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA))  'Txt_campo6.Text
'                 'Ado_datosA.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me1A = "", "0", txt_dcto_me1A)
'                 'Ado_datosA.Recordset!cotiza_precio_dcto_bs = Round(CDbl(txt_dcto_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me1A = "", "0", txt_seguro_me1A)
'                 'Ado_datosA.Recordset!cotiza_precio_seg_bs = Round(CDbl(txt_seguro_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_fob_seg_dol = Round(CDbl(txt_fob_me1A) - CDbl(txt_dcto_me1A) + CDbl(txt_seguro_me1A) + CDbl(txt_tacb2) + CDbl(txt_spread2), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_fob_seg_bs = Round(CDbl(txt_fob_seg_dolA) * CDbl(GlTipoCambioOficial), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me1A = "", 0, txt_fletefrontera_me1A)
'                 'Ado_datosA.Recordset!cotiza_precio_flete_bs = Round(CDbl(txt_fletefrontera_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_precio_tacb_dol = IIf(txt_tacb2 = "", 0, CDbl(txt_tacb2))
'                 'Ado_datosA.Recordset!cotiza_precio_tacb_bs = IIf(txt_tacb1 = "", "0.035", CDbl(txt_tacb1))
'                 'Ado_datosA.Recordset!cotiza_precio_spread_dol = IIf(txt_spread2 = "", "0", CDbl(txt_spread2))
'                 'If txt_spread1.Text = "" Then
'                 '   txt_spread1.Text = "0.08"
'                 'End If
'                 'Ado_datosA.Recordset!cotiza_precio_spread_bs = CDbl(txt_spread1.Text)       'IIf(txt_spread1.Text = "", 0.08, CDbl(txt_spread1))
'                 'Ado_datosA.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_me1A) - CDbl(txt_dcto_me1A.Text) + CDbl(txt_seguro_me1A.Text) + CDbl(txt_fletefrontera_me1A.Text) + CDbl(txt_tacb2) + CDbl(txt_spread2), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA)) '
'                 'Ado_datosA.Recordset!cotiza_precio_GAC_dol = IIf(txt_GAC_dol = "", "0", CDbl(txt_GAC_dol))
'                 'If txt_GAC_bs = "" Then
'                 '   txt_GAC_bs = "0.05"
'                 'End If
'                 'Ado_datosA.Recordset!cotiza_precio_GAC_bs = CDbl(txt_GAC_bs)  'IIf(txt_gac_bs = "", "0.05", CDbl(txt_gac_bs))
'                 'Ado_datosA.Recordset!cotiza_precio_base_dol = Round(CDbl(txt_cif_me1A) + CDbl(txt_GAC_dol.Text), Val(cmd_decA))
'                 'Ado_datosA.Recordset!cotiza_precio_base_bs = Round(CDbl(txt_base_imp_dol) * CDbl(GlTipoCambioOficial), Val(cmd_decA)) '
'                 'Ado_datosA.Recordset!fecha_registro = Date     'no cambia
'                 'Ado_datosA.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'                 'Ado_datosA.Recordset.Update    'Batch 'adAffectAll
'                 'db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                  'WWWWWWWWWWWWWWWWWWWWWW
'                 ' REGISTRO ACTIVO        'GRABA costo_monto
'                 Set rs_aux5 = New ADODB.Recordset
'                 If rs_aux5.State = 1 Then rs_aux5.Close
'                 rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
'                 If rs_aux5.RecordCount = 0 Then
''                    Call GRABA_COSTOS
'                 Else
'                    sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
'                    If sino = vbYes Then
'                        'OJO BORRAR ao_solicitud_costos
'                        db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
'                        'corrprog = 0
''                        Call GRABA_COSTOS
'                    Else
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            VAR_NAC = rs_aux6!costo_monto_usd
'                        End If
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            VAR_ALM = rs_aux6!costo_monto_usd
'                        End If
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            VAR_AGE = rs_aux6!costo_monto_usd
'                        End If
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                        End If
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'AMERICA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux6.RecordCount > 0 Then
'                            VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                        End If
'                    End If
'
'                 End If
'                 'WWWWWWWWWWWWWWWWWWWWWWWW
'                 If txt_tdcA.Text = "0" Or txt_tdcA.Text = "" Then
'                    txt_tdcA.Text = GlTipoCambioOficial
'                 End If
'                 If txt_local_IT_bsA.Text = "" Then
'                    txt_local_IT_bsA.Text = "0.0309"
'                 End If
'                 If txt_local_IVA_bsA.Text = "" Then
'                    txt_local_IVA_bsA.Text = "0.1491"
'                 End If
'                 If txt_cge_IT_bsA.Text = "" Then
'                    txt_cge_IT_bsA = "0.0416"
'                 End If
'                 If txt_cge_IVA_bsA.Text = "" Then
'                    txt_cge_IVA_bsA = "0.151"
'                 End If
'                 If txt_tac_billing_bsA.Text = "" Then
'                    txt_tac_billing_bsA = "0.035"
'                 End If
'                 If txt_GAC_bs = "" Then
'                    txt_GAC_bs = "0.05"
'                 End If
'                 If Ado_datosA.Recordset!pais_continente = "ASIA" And Val(txt_codigo1.Caption) = 1 Then
'                        'txt_local_IT_bsA.Text = "0.0309"
'                        'txt_local_IVA_bsA.Text = "0.1491"
'                        'txt_cge_IT_bsA = "0.0416"
'                        'txt_cge_IVA_bsA = "0.151"
'                        'txt_tac_billing_bsA = "0.035"
'                        Set rs_aux4 = New ADODB.Recordset
'                        If rs_aux4.State = 1 Then rs_aux4.Close
'                        rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  ", db, adOpenKeyset, adLockOptimistic
'                        If rs_aux4.RecordCount > 0 Then
'                            SUBTOTD = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        Else
'                            'SUBTOTD = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            SUBTOTD = Round(Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        End If
'                        'Importaion Cliente
'                        VAR_LOCAL = Round(rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(VAR_LOCAL, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(VAR_LOCAL * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bsA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IT_bsA.Text), Val(cmd_decA)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        txt_local_IT_dolA.Text = Round(VAR_LOCAL * CDbl(txt_local_IT_bsA.Text), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bsA.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IVA_bsA.Text), Val(cmd_decA)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        txt_local_IVA_dolA = Round(VAR_LOCAL * CDbl(txt_local_IVA_bsA.Text), Val(cmd_decA))
'
'                        VAR_DOLCLI2 = Round(SUBTOTD + CDbl(txt_local_IVA_dolA) + CDbl(txt_local_IVA_dolA), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(VAR_DOLCLI2, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(VAR_DOLCLI2 * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                        VAR_DOLCLI = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_cif_dol - Ado_datosA.Recordset!cotiza_precio_fob_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol, Val(cmd_decA))
'                        VAR_BSCLI = Round(VAR_DOLCLI * GlTipoCambioOficial, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                        'VAR_SUBD = Round(SUBTOTD - Ado_datosA.Recordset!cotiza_precio_seg_dol, Val(cmd_decA))       'Sin Seguro
'                        VAR_SUBD = Round(SUBTOTD, Val(cmd_decA))        'Con Seguro
'                        VAR_SUBB = Round(VAR_SUBD * GlTipoCambioOficial, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(VAR_SUBD * CDbl(txt_cge_IT_bsA), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "       'Con Seguro
'                        txt_cge_IT_dolA = Round(VAR_SUBD * CDbl(txt_cge_IT_bsA), Val(cmd_decA))
'
'                        'IMPORTACION CGE
'                        txt_cge_IVA_dolA = Round((VAR_SUBD * CDbl(txt_cge_IVA_bsA)) - ((Ado_datosA.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_decA))        'Sin Seguro
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_dolA, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "    'Con Seguro
'                        txt_tac_billing_dolA = Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA))
'
'                        VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_dolA) + CDbl(txt_cge_IVA_dolA) + CDbl(txt_tac_billing_dolA), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                 End If
'
'             Else
'                'CLONA REGISTROS
'                Set rs_aux7 = New ADODB.Recordset
'                If rs_aux7.State = 1 Then rs_aux7.Close
'                rs_aux7.Open "ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = 1  ", db, adOpenStatic
'                'Set Ado_datos11.Recordset = rs_aux7
'                If rs_aux7.RecordCount > 0 Then
'                     'WWWWWWWWWWWWWWWWWWWWWW
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_dec = " & rs_aux7!cotiza_dec & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set tipo_moneda= '" & rs_aux7!tipo_moneda & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_tdc_bol = " & rs_aux7!cotiza_tdc_bol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set costo_monto = " & rs_aux7!costo_monto & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_dol = " & rs_aux7!cotiza_precio_fob_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_fob_bs = " & Round(CDbl(rs_aux7!cotiza_precio_fob_bs), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_dol = " & rs_aux7!cotiza_precio_dcto_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_dcto_bs = " & CDbl(rs_aux7!cotiza_precio_dcto_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_dol = " & rs_aux7!cotiza_precio_seg_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_seg_bs = " & CDbl(rs_aux7!cotiza_precio_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_dol = " & CDbl(rs_aux7!cotiza_fob_seg_dol) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_fob_seg_bs = " & CDbl(rs_aux7!cotiza_fob_seg_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_dol = " & rs_aux7!cotiza_precio_flete_dol & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_flete_bs = " & CDbl(rs_aux7!cotiza_precio_flete_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_dol = " & Round(rs_aux7!cotiza_precio_tacb_dol, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_tacb_bs = " & CDbl(rs_aux7!cotiza_precio_tacb_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_dol  = " & Round(rs_aux7!cotiza_precio_spread_dol, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_spread_bs  = " & CDbl(rs_aux7!cotiza_precio_spread_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_dol = " & Round(CDbl(rs_aux7!cotiza_precio_cif_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_cif_bs = " & Round(rs_aux7!cotiza_precio_cif_bs, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_dol = " & Round(rs_aux7!cotiza_precio_GAC_dol, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_GAC_bs  = " & Round(rs_aux7!cotiza_precio_GAC_bs, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_dol  = " & Round(rs_aux7!cotiza_precio_base_dol, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_base_bs  = " & Round(rs_aux7!cotiza_precio_base_bs, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(rs_aux7!cotiza_precio_total_bs, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cli), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(rs_aux7!cotiza_precio_total_bs_cli, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(CDbl(rs_aux7!cotiza_precio_total_dol_cge), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(rs_aux7!cotiza_precio_total_bs_cge, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(CDbl(rs_aux7!cotiza_gasto_local_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(rs_aux7!cotiza_gasto_local_bs, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IT_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(rs_aux7!cotiza_saldo_local_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_local_IVA_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(rs_aux7!cotiza_saldo_local_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IT_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(rs_aux7!cotiza_saldo_cge_IT_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_cge_IVA_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(rs_aux7!cotiza_saldo_cge_IVA_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round(CDbl(rs_aux7!cotiza_saldo_tac_billing_dol), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(rs_aux7!cotiza_saldo_tac_billing_bs) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set fecha_registro = '" & Date & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set usr_codigo = '" & glusuario & "' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'SI' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'                    'WWWWWWWWWWWWWWWWWWWWWW
'                End If
'                Set rs_aux5 = New ADODB.Recordset
'                If rs_aux5.State = 1 Then rs_aux5.Close
'                rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
'                If rs_aux5.RecordCount = 0 Then
''                   Call GRABA_COSTOS_CLON
'                Else
'                   sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
'                   If sino = vbYes Then
'                       'OJO BORRAR ao_solicitud_costos
'                       db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                       'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
'                       'corrprog = 0
''                       Call GRABA_COSTOS_CLON
'                   End If
'                End If
'
'             End If
'
'           rs_datos10.MoveNext
'           'Ado_datosA.Recordset.MoveNext
'           Wend
'       Else
'             '- SOLO EL REGISTRO ACTIVO
'               Ado_datosA.Recordset!cotiza_dec = cmd_decA.Text
'             Ado_datosA.Recordset!tipo_moneda = cmd_monedaA.Text
'             If txt_tdcA.Text = "0" Or txt_tdcA.Text = "" Then
'                txt_tdcA.Text = GlTipoCambioOficial
'             End If
'             Ado_datosA.Recordset!cotiza_tdc_bol = txt_tdcA.Text
'             Ado_datosA.Recordset!costo_monto = txt_montobaseA.Text
'             Ado_datosA.Recordset!cotiza_precio_fob_dol = IIf(txt_fob_me1A = "", "0", txt_fob_me1A)
'             Ado_datosA.Recordset!cotiza_precio_fob_bs = Round(CDbl(txt_fob_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA))  'Txt_campo6.Text
'             Ado_datosA.Recordset!cotiza_precio_dcto_dol = IIf(txt_dcto_me1A = "", "0", txt_dcto_me1A)
'             Ado_datosA.Recordset!cotiza_precio_dcto_bs = CDbl(txt_dcto_me1A) * CDbl(GlTipoCambioOficial)
'             Ado_datosA.Recordset!cotiza_precio_seg_dol = IIf(txt_seguro_me1A = "", "0", txt_seguro_me1A)
'             Ado_datosA.Recordset!cotiza_precio_seg_bs = CDbl(txt_seguro_me1A) * CDbl(GlTipoCambioOficial)
'
'             Ado_datosA.Recordset!cotiza_fob_seg_dol = CDbl(txt_fob_me1A) - CDbl(txt_dcto_me1A) + CDbl(txt_seguro_me1A) + CDbl(txt_tacb2) + CDbl(txt_spread2)
'             Ado_datosA.Recordset!cotiza_fob_seg_bs = CDbl(txt_fob_seg_dolA) * CDbl(GlTipoCambioOficial)
'
'             Ado_datosA.Recordset!cotiza_precio_flete_dol = IIf(txt_fletefrontera_me1A = "", "0", txt_fletefrontera_me1A)
'             Ado_datosA.Recordset!cotiza_precio_flete_bs = CDbl(txt_fletefrontera_me1A) * CDbl(GlTipoCambioOficial)
'
'             Ado_datosA.Recordset!cotiza_precio_tacb_dol = IIf(txt_tacb2 = "", "0", CDbl(txt_tacb2))
'             Ado_datosA.Recordset!cotiza_precio_tacb_bs = IIf(txt_tacb1 = "", "0.035", CDbl(txt_tacb1))
'             Ado_datosA.Recordset!cotiza_precio_spread_dol = IIf(txt_spread2 = "", "0", CDbl(txt_spread2))
'             Ado_datosA.Recordset!cotiza_precio_spread_bs = IIf(txt_spread1 = "", "0.08", CDbl(txt_spread1))
'
'             'Ado_datosA.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_me1A) - CDbl(txt_dcto_me1A.Text) + CDbl(txt_seguro_me1A.Text) + CDbl(txt_fletefrontera_me1A.Text) + CDbl(txt_tacb2) + CDbl(txt_spread2), Val(cmd_decA))
'             Ado_datosA.Recordset!cotiza_precio_cif_dol = Round(CDbl(txt_fob_seg_dolA) + CDbl(txt_fletefrontera_me1A.Text), Val(cmd_decA))
'             Ado_datosA.Recordset!cotiza_precio_cif_bs = Round(CDbl(txt_cif_me1A) * CDbl(GlTipoCambioOficial), Val(cmd_decA)) '
'
'             Ado_datosA.Recordset!cotiza_precio_GAC_dol = IIf(txt_GAC_dol = "", "0", CDbl(txt_GAC_dol))
'             Ado_datosA.Recordset!cotiza_precio_GAC_bs = IIf(txt_GAC_bs = "", "0.05", CDbl(txt_GAC_bs))
'             Ado_datosA.Recordset!cotiza_precio_base_dol = Round(CDbl(txt_cif_me1A) + CDbl(txt_GAC_dol.Text), Val(cmd_decA))
'             Ado_datosA.Recordset!cotiza_precio_base_bs = Round(CDbl(txt_base_imp_dol) * CDbl(GlTipoCambioOficial), Val(cmd_decA)) '
'             Ado_datosA.Recordset!fecha_registro = Date     'no cambia
'             Ado_datosA.Recordset!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
'             Ado_datosA.Recordset.Update    'Batch 'adAffectAll
'             db.Execute "update ao_solicitud_cotiza_venta set agrupado = 'NO' where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_CONTI & "' and cotiza_codigo = " & txt_codigo1.Caption & "  "
'             'GRABA COSTOS
'             Set rs_aux5 = New ADODB.Recordset
'             If rs_aux5.State = 1 Then rs_aux5.Close
'             rs_aux5.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   ", db, adOpenKeyset, adLockOptimistic
'             If rs_aux5.RecordCount = 0 Then
''                Call GRABA_COSTOS
'             Else
'                sino = MsgBox("La Hoja de Costos ya existe, desea volver a Generarla ? ...", vbYesNo + vbQuestion, "Atención ...")
'                If sino = vbYes Then
'                    'OJO BORRAR ao_solicitud_costos
'                    db.Execute "DELETE ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    'db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '0' where venta_codigo= " & var_cod5 & " "
'                    'corrprog = 0
''                    Call GRABA_COSTOS
'                Else
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '3' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_NAC = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '5' ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_ALM = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '6'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_AGE = rs_aux6!costo_monto_usd
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '8'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_FLE = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                    Set rs_aux6 = New ADODB.Recordset
'                    If rs_aux6.State = 1 Then rs_aux6.Close
'                    rs_aux6.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  and codigo_costo = '14'  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux6.RecordCount > 0 Then
'                        VAR_UTIL = IIf(IsNull(rs_aux6!costo_monto_usd), "0", rs_aux6!costo_monto_usd)
'                    End If
'                End If
'
'             End If
'             If Ado_datosA.Recordset!pais_continente = "ASIA" Then
'                    Set rs_aux4 = New ADODB.Recordset
'                    If rs_aux4.State = 1 Then rs_aux4.Close
'                    'rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA'   ", db, adOpenKeyset, adLockOptimistic
'                    rs_aux4.Open "select sum(costo_monto) as totbs, sum (costo_monto_usd) as totdl from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  ", db, adOpenKeyset, adLockOptimistic
'                    If rs_aux4.RecordCount > 0 Then
'                            SUBTOTD = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        Else
'                            'SUBTOTD = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            SUBTOTD = Round(Ado_datosA.Recordset!cotiza_precio_base_dol - Ado_datosA.Recordset!cotiza_precio_flete_dol, Val(cmd_decA))
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol = " & Round(SUBTOTD, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                            db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs = " & Round(SUBTOTD * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                    End If
''                    'Importaion Cliente
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & rs_aux4!totdl & " - " & VAR_NAC & " - " & VAR_ALM & " - " & VAR_AGE & " - " & VAR_FLE & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & rs_aux4!totbs & " - " & VAR_NAC & " - " & VAR_ALM & " - " & VAR_AGE & " - " & VAR_FLE & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    If txt_local_IT_bsA.Text = "" Then
''                    End If
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bsA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IT_bsA.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bsA.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = cotiza_gasto_local_dol * " & CDbl(txt_local_IVA_bsA.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
''
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = cotiza_precio_total_dol + cotiza_saldo_local_IT_dol + cotiza_saldo_local_IVA_dol where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = cotiza_precio_total_dol_cli * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''
''                    VAR_DOLCLI = rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_cif_dol - Ado_datosA.Recordset!cotiza_precio_fob_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol
''                    VAR_BSCLI = rs_aux4!totbs + Ado_datosA.Recordset!cotiza_precio_cif_bs - Ado_datosA.Recordset!cotiza_precio_fob_bs - Ado_datosA.Recordset!cotiza_precio_seg_bs
''
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
''                    'VAR_SUBD = Ado_datosA.Recordset!cotiza_precio_total_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol    'Sin Seguro
''                    VAR_SUBD = Ado_datosA.Recordset!cotiza_precio_total_dol                                                 'Con Seguro
''                    VAR_SUBB = VAR_SUBD * GlTipoCambioOficial
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IT_bsA) & ") where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    txt_cge_IT_dolA = Round(VAR_SUBD * CDbl(txt_cge_IT_bsA), Val(cmd_decA))
''
''                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = (" & VAR_SUBD & " * " & CDbl(txt_cge_IVA_bsA) & ") -((cotiza_precio_cif_dol * 0.1498) )-((" & CDbl(VAR_AGE) & " * 0.13))  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''
''                    txt_cge_IVA_dolA = Round((VAR_SUBD * CDbl(txt_cge_IVA_bsA)) - ((Ado_datosA.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_decA))        'Sin Seguro
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_dolA, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "    'Con Seguro
''                    txt_tac_billing_dolA = Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA))
''
''                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & VAR_SUBD & "  + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IT_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_cge_IVA_dol) + (ao_solicitud_cotiza_venta.cotiza_saldo_tac_billing_dol) where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    'db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = ao_solicitud_cotiza_venta.cotiza_precio_total_dol_cge * " & GlTipoCambioOficial & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = 'ASIA' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''
''                    VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_dolA) + CDbl(txt_cge_IVA_dolA) + CDbl(txt_tac_billing_dolA), Val(cmd_decA))
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
''                    db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                    'Importaion Cliente
'                        VAR_LOCAL = Round(rs_aux4!totdl - VAR_NAC - VAR_ALM - VAR_AGE - VAR_FLE, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_dol = " & Round(VAR_LOCAL, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_gasto_local_bs = " & Round(VAR_LOCAL * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_bs = " & CDbl(txt_local_IT_bsA.Text) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IT_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IT_bsA.Text), Val(cmd_decA)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        txt_local_IT_dolA.Text = Round(VAR_LOCAL * CDbl(txt_local_IT_bsA.Text), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_bs = " & CDbl(txt_local_IVA_bsA.Text) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_local_IVA_dol = " & Round(VAR_LOCAL * CDbl(txt_local_IVA_bsA.Text), Val(cmd_decA)) & "  where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & "  AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "   "
'                        txt_local_IVA_dolA = Round(VAR_LOCAL * CDbl(txt_local_IVA_bsA.Text), Val(cmd_decA))
'
'                        VAR_DOLCLI2 = Round(SUBTOTD + CDbl(txt_local_IT_dolA) + CDbl(txt_local_IVA_dolA), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cli = " & Round(VAR_DOLCLI2, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cli = " & Round(VAR_DOLCLI2 * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                        VAR_DOLCLI = Round(rs_aux4!totdl + Ado_datosA.Recordset!cotiza_precio_cif_dol - Ado_datosA.Recordset!cotiza_precio_fob_dol - Ado_datosA.Recordset!cotiza_precio_seg_dol, Val(cmd_decA))
'                        VAR_BSCLI = Round(VAR_DOLCLI * GlTipoCambioOficial, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_totusd_menos_seguro = " & VAR_DOLCLI & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & " "
'
'                        'VAR_SUBD = Round(SUBTOTD - Ado_datosA.Recordset!cotiza_precio_seg_dol, Val(cmd_decA))       'Sin Seguro
'                        VAR_SUBD = Round(SUBTOTD, Val(cmd_decA))        'Con Seguro
'                        VAR_SUBB = Round(VAR_SUBD * GlTipoCambioOficial, Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_bs = " & CDbl(txt_cge_IT_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IT_dol = " & Round(VAR_SUBD * CDbl(txt_cge_IT_bsA), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "       'Con Seguro
'                        txt_cge_IT_dolA = Round(VAR_SUBD * CDbl(txt_cge_IT_bsA), Val(cmd_decA))
'
'                        'IMPORTACION CGE
'                        txt_cge_IVA_dolA = Round((VAR_SUBD * CDbl(txt_cge_IVA_bsA)) - ((Ado_datosA.Recordset!cotiza_precio_cif_dol * 0.1498)) - ((CDbl(VAR_AGE) * 0.13)), Val(cmd_decA))        'Sin Seguro
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_bs = " & CDbl(txt_cge_IVA_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_cge_IVA_dol = " & Round(txt_cge_IVA_dolA, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_bs = " & CDbl(txt_tac_billing_bsA) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_saldo_tac_billing_dol = " & Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "    'Con Seguro
'                        txt_tac_billing_dolA = Round((VAR_SUBD * CDbl(txt_tac_billing_bsA)), Val(cmd_decA))
'
'                        VAR_DOLCGE = Round(VAR_SUBD + CDbl(txt_cge_IT_dolA) + CDbl(txt_cge_IVA_dolA) + CDbl(txt_tac_billing_dolA), Val(cmd_decA))
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_dol_cge = " & Round(VAR_DOLCGE, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'                        db.Execute "update ao_solicitud_cotiza_venta set cotiza_precio_total_bs_cge = " & Round(VAR_DOLCGE * GlTipoCambioOficial, Val(cmd_decA)) & " where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " AND pais_continente = '" & VAR_CONTI & "' AND cotiza_codigo = " & CDbl(txt_codigo1) & "  "
'
'             End If
'         End If
'
'     End If
'     SSTab1.Tab = 1
''     If Ado_datosA.Recordset!pais_continente = "AMERICA" Then
'     If VAR_CONTI = "AMERICA" Then
'         SSTab1.TabEnabled(0) = True
'     Else
'        SSTab1.TabEnabled(0) = False
'     End If
'     If VAR_CONTI = "ASIA" Then
'        SSTab1.TabEnabled(1) = True
'     Else
'        SSTab1.TabEnabled(1) = False
'     End If
'     If VAR_CONTI = "EUROPA" Then
'        SSTab1.TabEnabled(2) = True
'     Else
'        SSTab1.TabEnabled(2) = False
'     End If
'     Call ABRIR_TABLA
''     rs_datosA.MoveLast
''     mbDataChanged = False
''        Fra_datos.Enabled = False
'        FraModeloCostoA.Visible = False
'        FraGrabarCancelarA.Visible = False
'        Fra_datos2.Enabled = False
'        fraOpciones2A.Visible = True
'        fraOpciones1A.Visible = True
'        FrmABMDet.Visible = True
'        FraDet1.Enabled = True
'        dg_datosA.Enabled = True
'        dg_datos1A.Enabled = True
'        VAR_SW = ""
''        SSTab1.Tab = 1
''        SSTab1.TabEnabled(0) = False
''        SSTab1.TabEnabled(1) = True
''        SSTab1.TabEnabled(2) = False
''     dtc_codigo9.Enabled = True
'  End If
''  dtc_desc1.Visible = True
''  lbl_aux1.Visible = False
'  Exit Sub
'UpdateErr:
'  MsgBox Err.Description
'End Sub

Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    'CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_cotizacion_equipos.rpt"
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R222_ar_cotiza_venta_cliente_ame.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!EDIF_CODIGO
    CR01.StoredProcParam(4) = Me.Ado_datos.Recordset!cotiza_codigo
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_R224_Ame.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!EDIF_CODIGO
    CR01.StoredProcParam(4) = Me.Ado_datos.Recordset!cotiza_codigo
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimirA_Click()
If Ado_datosA.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    'CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_cotizacion_equipos.rpt"
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R222_ar_cotiza_venta_cliente_asia.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datosA.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datosA.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datosA.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datosA.Recordset!EDIF_CODIGO
    CR01.StoredProcParam(4) = Me.Ado_datosA.Recordset!cotiza_codigo
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized

End Sub

Private Sub BtnModDetalle_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        VARCTRL = 0
        Select Case SSTab1.Tab
            Case 0
                marca1 = Ado_datos.Recordset.Bookmark
                If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
                    'FraNavega.Enabled = False
                    FraDet1E.Visible = False
                    FraDet1.Visible = True
                    FraDet1.Enabled = True
                    dg_det1.Enabled = True
                    dg_det1.AllowUpdate = True
                    VARCTRL = 1
                    VAR_CONTI = "AMERICA"
                End If
            Case 1
                marca1 = Ado_datosA.Recordset.Bookmark
                If rs_datosA.RecordCount > 0 And rs_datosA!estado_codigo = "REG" Then
                    'FraNavegaA.Enabled = False
                    FraDet1E.Visible = False
                    FraDet1.Visible = True
                    FraDet1.Enabled = True
                    dg_det1.Enabled = True
                    dg_det1.AllowUpdate = True
                    VARCTRL = 3
                    VAR_CONTI = "ASIA"
                End If
            Case 2
                marca1 = Ado_datosE.Recordset.Bookmark
                If rs_datosE.RecordCount > 0 And rs_datosE!estado_codigo = "REG" Then
                    'FraNavegaE.Enabled = False
                    FraDet1.Visible = False
                    FraDet1E.Enabled = True
                    dg_det1E.Enabled = True
                    dg_det1E.AllowUpdate = True
                    VARCTRL = 2
                    VAR_CONTI = "EUROPA"
                End If
        End Select
        swnuevo = 2
        fraOpciones.Visible = False
'        Fra_datos2.Enabled = False
        FraNavega0.Enabled = False
        'FrmABMDet.Enabled = False
        SSTab1.Enabled = False
        BtnGrabar.Visible = True
        BtnModDetalle.Visible = False
        BtnAddDetalle2.Visible = False
        BtnAnlDetalle.Visible = False
    Else
        MsgBox "No se puede Modificar, Verifique y vuelva a intentar !! ", vbExclamation
    End If
'    Select Case SSTab1.Tab
'        Case 0
'            marca1 = Ado_datos.Recordset.Bookmark
'            If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
'                FraNavega.Enabled = False
''                FraNavega1.Enabled = False
''                FraModeloCosto.Enabled = False
'                VARCTRL = 1
'                If VARCTRL = 1 Then
'                    aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")    ' Codigo Unidad
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("cotiza_codigo")    ' Nro. Cotización
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = Me.Ado_detalle1.Recordset("edif_codigo")      ' Codigo Edificio
'
'                    aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("codigo_costo")     ' Codigo Costo
'                    aw_p_ao_solicitud_cotiza_detalle.dtc_desc1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
'                    aw_p_ao_solicitud_cotiza_detalle.dtc_aux1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
'                    aw_p_ao_solicitud_cotiza_detalle.dtc_aux2.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
'
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_campo5.Caption = Me.Ado_detalle1.Recordset("pais_continente")    ' Continente
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_campo3.Text = Me.Ado_detalle1.Recordset("costo_porcentaje")    ' % Costo
'                    If Ado_datos.Recordset!cotiza_precio_fob_dol = "0" Or IsNull(Ado_datos.Recordset!cotiza_precio_fob_dol) Then
'                    'If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'                        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
'                    Else
'                        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Ado_datos.Recordset!cotiza_precio_fob_dol   ' Monto Modelo1(ME)
'                    End If
'                    aw_p_ao_solicitud_cotiza_detalle.Txt_campo4.Text = Me.Ado_detalle1.Recordset("costo_observaciones") ' Observaciones
'                    aw_p_ao_solicitud_cotiza_detalle.Show vbModal
'            '    Else
'            '        MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
'                End If
'            End If
'
'            FraNavega.Enabled = True
'        Case 1
'            marca1 = Ado_datosA.Recordset.Bookmark
'            If rs_datosA.RecordCount > 0 And rs_datosA!estado_codigo = "REG" Then
'                FraNavegaA.Enabled = False
''                FraNavega1A.Enabled = False
''                FraModeloCostoA.Enabled = False
'                VARCTRL = 3
'                    'ASIA
'                If VARCTRL = 3 Then
'                    aw_p_ao_solicitud_cotiza_det_asia.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")    ' Codigo Unidad
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("cotiza_codigo")    ' Nro. Cotización
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo2.Caption = Me.Ado_detalle1.Recordset("edif_codigo")      ' Codigo Edificio
'
'                    aw_p_ao_solicitud_cotiza_det_asia.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("codigo_costo")     ' Codigo Costo
'                    aw_p_ao_solicitud_cotiza_det_asia.dtc_desc1.BoundText = aw_p_ao_solicitud_cotiza_det_asia.dtc_codigo1.BoundText
'                    aw_p_ao_solicitud_cotiza_det_asia.dtc_aux1.BoundText = aw_p_ao_solicitud_cotiza_det_asia.dtc_codigo1.BoundText
'                    aw_p_ao_solicitud_cotiza_det_asia.dtc_aux2.BoundText = aw_p_ao_solicitud_cotiza_det_asia.dtc_codigo1.BoundText
'
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo5.Caption = Me.Ado_detalle1.Recordset("pais_continente")    ' Continente
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo3.Text = Me.Ado_detalle1.Recordset("costo_porcentaje")    ' % Costo
'
'                    aw_p_ao_solicitud_cotiza_det_asia.lbl_decA.Caption = Ado_datosA.Recordset!cotiza_dec     'cmd_decA.Text      ' # Decimales
'                    If Ado_datosA.Recordset!cotiza_precio_fob_dol = "0" Or IsNull(Ado_datosA.Recordset!cotiza_precio_fob_dol) Then
'                    'If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'                        aw_p_ao_solicitud_cotiza_det_asia.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
'                    Else
'                        aw_p_ao_solicitud_cotiza_det_asia.txt_monto01.Caption = Ado_datosA.Recordset!cotiza_precio_fob_dol   ' Monto Modelo1(ME)
'                    End If
'                    aw_p_ao_solicitud_cotiza_det_asia.Txt_campo4.Text = Me.Ado_detalle1.Recordset("costo_observaciones") ' Observaciones
'                    aw_p_ao_solicitud_cotiza_det_asia.Show vbModal
'            '    Else
'            '        MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
'                End If
'
'            End If
'            FraNavegaA.Enabled = True
'        Case 2
'        marca1 = Ado_datosE.Recordset.Bookmark
'        If rs_datosE.RecordCount > 0 And rs_datosE!estado_codigo = "REG" Then
'            FraNavegaE.Enabled = False
''            FraModeloCostoE.Enabled = False
''            FraNavega1E.Enabled = False
'            VARCTRL = 2
'            'EUROPA
'            If VARCTRL = 2 Then
'                aw_p_ao_solicitud_cotiza_det_eur.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")    ' Codigo Unidad
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("cotiza_codigo")    ' Nro. Cotización
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo2.Caption = Me.Ado_detalle1.Recordset("edif_codigo")      ' Codigo Edificio
'
'                aw_p_ao_solicitud_cotiza_det_eur.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("codigo_costo")     ' Codigo Costo
'                aw_p_ao_solicitud_cotiza_det_eur.dtc_desc1.BoundText = aw_p_ao_solicitud_cotiza_det_eur.dtc_codigo1.BoundText
'                aw_p_ao_solicitud_cotiza_det_eur.dtc_aux1.BoundText = aw_p_ao_solicitud_cotiza_det_eur.dtc_codigo1.BoundText
'                aw_p_ao_solicitud_cotiza_det_eur.dtc_aux2.BoundText = aw_p_ao_solicitud_cotiza_det_eur.dtc_codigo1.BoundText
'
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo3.Text = Me.Ado_detalle1.Recordset("costo_porcentaje")    ' % Costo
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo6 = Ado_datosE.Recordset!cotiza_tdc_me
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo7 = Ado_datosE.Recordset!cotiza_dec
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo8 = Ado_datosE.Recordset!cotiza_tdc_bol
'
'        '        If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'        '            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
'        '        Else
'        '            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Me.txt_fob_me1.Text   ' Monto Modelo1(ME)
'        '        End If
'                aw_p_ao_solicitud_cotiza_det_eur.txt_monto01.Caption = IIf(IsNull(Ado_datosE.Recordset!cotiza_precio_base_me), "0", Ado_datosE.Recordset!cotiza_precio_base_me)
'
'                aw_p_ao_solicitud_cotiza_det_eur.Txt_campo4.Text = Me.Ado_detalle1.Recordset("costo_observaciones") ' Observaciones
'                aw_p_ao_solicitud_cotiza_det_eur.Show vbModal
'            End If
'
'        End If
'        FraNavegaE.Enabled = True
'    End Select
''    Select Case SSTab1.Tab
''        Case 0
''
'''            FraNavega1.Enabled = True
'''            FraModeloCosto.Enabled = True
''        Case 1
''            FraNavegaA.Enabled = True
'''            FraNavega1A.Enabled = True
'''            FraModeloCostoA.Enabled = True
''        Case 2
''            FraNavegaE.Enabled = False
'''            FraModeloCostoE.Enabled = False
'''            FraNavega1E.Enabled = False
''    End Select
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
  
End Sub

Private Sub BtnModificar_Click()
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos0.RecordCount > 0 Then
    If Ado_datos0.Recordset!estado_codigo_cot = "REG" Then
        If Ado_datos0.Recordset!pais_continente_cot = "AMERICA" Then
            'swnuevo = 2
'            VAR_SW = Ado_datos0.Recordset!SOLICITUD_codigo
'            VAR_SW = Ado_datos0.Recordset!unidad_codigo
            fraOpciones1.Visible = False
            FraNavega.Enabled = False
            FraDet1.Enabled = False
            FrmABMDet.Visible = False
'            FraModeloCosto.Enabled = False
            Fra_datos2.Enabled = False
            VAR_SW = "MOD"
        '    txt_fob_me1.SetFocus
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
    '    Select Case dtc_codigo2.Text
    '        Case "1"
    '        Case "2"
    '        Case "3"
            'Call ABRIR_TABLA_DET
            VAR_CONTI = Ado_datos0.Recordset!pais_continente_cot
            GlConti = VAR_CONTI
            GlSolicitud = Me.Ado_datos0.Recordset!solicitud_codigo           ' Nro. Negociacion (Cod.solicitud)
            GlUnidad = Me.Ado_datos0.Recordset!unidad_codigo                 ' Codigo Unidad
            GlNombFor = Me.txt_campo12                                          ' Descripcion Unidad
            GlCotiza = Me.Ado_datos0.Recordset!cotiza_codigo                ' Nro. Cotización
            GlEdificio = Me.Ado_datos0.Recordset!EDIF_CODIGO                ' Codigo Edificio
            aw_solicitud_cotiza_datos.Show vbModal
            
            VAR_CONTI = "AMERICA"
'            frm_ao_solicitud_cotiza_datos.txt_conti.Caption = VAR_CONTI
'            frm_ao_solicitud_cotiza_datos.txt_codigo.Caption = Me.Ado_datos0.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'            frm_ao_solicitud_cotiza_datos.Txt_campo1.Caption = Me.Ado_datos0.Recordset("unidad_codigo")    ' Codigo Unidad
'            frm_ao_solicitud_cotiza_datos.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'            frm_ao_solicitud_cotiza_datos.Txt_Correl.Caption = Me.Ado_datos0.Recordset("cotiza_codigo")    ' Nro. Cotización
'            frm_ao_solicitud_cotiza_datos.Txt_campo2A.Caption = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'            GlEdificio = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'
'            frm_ao_solicitud_cotiza_datos.Txt_campo4.Text = Me.Ado_datos0.Recordset("modelo_codigo") ' Modelo
'            frm_ao_solicitud_cotiza_datos.Txt_campo5.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_nro_montador), "2", Me.Ado_datos0.Recordset!cotiza_nro_montador) ' Montadores
'            frm_ao_solicitud_cotiza_datos.Txt_campo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_energia), "0", Me.Ado_datos0.Recordset!cotiza_energia) ' Energia
'            frm_ao_solicitud_cotiza_datos.Txt_campo3.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_luz), "0", Me.Ado_datos0.Recordset!cotiza_luz) ' Luz
'            frm_ao_solicitud_cotiza_datos.Txt_campo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!bien_cotiza_num_accesos), "0", Me.Ado_datos0.Recordset!bien_cotiza_num_accesos) ' Num Accesos
'            frm_ao_solicitud_cotiza_datos.Txt_campo9.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_fondo), "0", Me.Ado_datos0.Recordset!dimension_fosa_fondo) ' Fosa fondo
'            frm_ao_solicitud_cotiza_datos.Txt_campo10.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_frente), "0", Me.Ado_datos0.Recordset!dimension_fosa_frente) ' Fosa Frente
'            frm_ao_solicitud_cotiza_datos.Txt_campo8.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_m), "0", Me.Ado_datos0.Recordset!dimension_fosa_m) ' Espacio Dintel
'            'Equipo
'            frm_ao_solicitud_cotiza_datos.dtc_codigo21.Text = IIf(IsNull(Me.Ado_datos0.Recordset!bien_codigo), "NA1", Me.Ado_datos0.Recordset!bien_codigo)     ' Codigo Equipo
'            frm_ao_solicitud_cotiza_datos.dtc_desc21.BoundText = frm_ao_solicitud_cotiza_datos.dtc_codigo21.BoundText
'            frm_ao_solicitud_cotiza_datos.dtc_desc24.BoundText = frm_ao_solicitud_cotiza_datos.dtc_codigo21.BoundText
'            'Pais
'            frm_ao_solicitud_cotiza_datos.dtc_codigo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!pais_codigo), "BRA", Me.Ado_datos0.Recordset!pais_codigo)    ' Pais
'            frm_ao_solicitud_cotiza_datos.dtc_desc7.BoundText = frm_ao_solicitud_cotiza_datos.dtc_codigo7.BoundText
'            'Tipo de Equipo
'            frm_ao_solicitud_cotiza_datos.dtc_codigo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!tipo_eqp), "A", Me.Ado_datos0.Recordset!tipo_eqp)    ' Tipo Equipo
'            frm_ao_solicitud_cotiza_datos.dtc_desc2.BoundText = frm_ao_solicitud_cotiza_datos.dtc_codigo2.BoundText
'            'Cuarto de Control
'            frm_ao_solicitud_cotiza_datos.dtc_codigo61.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cuadro_ctrl_codigo), "1", Me.Ado_datos0.Recordset!cuadro_ctrl_codigo)    'Cuarto de Control
'            frm_ao_solicitud_cotiza_datos.dtc_desc61.BoundText = frm_ao_solicitud_cotiza_datos.dtc_codigo61.BoundText
'
'            'Dimensión Cabina Frente (mm)                  Dimensión Cabina Lado (mm)                    Dimensión Cabina Alto (mm)
''            Txt_campo11 'Txt_campo12    Txt_campo13
'
'            'Motor
'
'    '        If txt_fob_bs1.Text = "0" Or txt_fob_bs1.Text = "" Then
'    '            frm_ao_solicitud_cotiza_datos.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
'    '        Else
'    '            frm_ao_solicitud_cotiza_datos.txt_monto01.Caption = Me.txt_fob_me1.Text   ' Monto Modelo1(ME)
'    '        End If
'            frm_ao_solicitud_cotiza_datos.Show vbModal
'    '        Case "4"
'    '
'    '    End Select
            fraOpciones1.Visible = True
            FraNavega.Enabled = True
'            FraModeloCosto.Enabled = False
            Fra_datos2.Enabled = False
            FraDet1.Enabled = True
            FrmABMDet.Visible = True
            dg_datos.Enabled = True
            VAR_SW = ""
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            swnuevo = 0
            Call ABRIR_TABLA
        Else
            MsgBox "El registro NO corresponde al continente: AMERICA, verifique los Parámetros de Cálculo por favor ...", vbExclamation, "Validación de Registro"
        End If
    Else
      MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
    End If
  Else
    MsgBox "No existe el Registro para Modificar, Vuelva a intentar...!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar1_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
'  On Error GoTo EditErr
''  lblStatus.Caption = "Modificar registro"
'   If Ado_datos.Recordset!estado_codigo = "REG" Then
'      If Ado_datos.Recordset!pais_continente = "AMERICA" Then
'        If Txt_campo5.Text = "" Then
'            MsgBox "Debe registrar el Número de Montadores, verifique por favor y vuelva a intentar...", vbExclamation, "Validación de Registro"
'            Exit Sub
'        End If
''        Fra_datos.Enabled = True
'        FraModeloCosto.Visible = True
'        FraModeloCosto.Enabled = True
'        FraGrabarCancelar.Visible = True
'        Fra_datos2.Enabled = True
'        'fraOpciones.Enabled = False
'        fraOpciones1.Visible = False
'        fraOpciones2.Visible = False
'        FrmABMDet.Visible = False
'        FraDet1.Visible = True
'        FraDet1.Enabled = False
'        FraDet1E.Visible = False
'        dg_datos.Enabled = False
'        dg_datos1.Enabled = False
'        VAR_SW = "MOD"
'        SSTab1.Tab = 0
'        SSTab1.TabEnabled(0) = True
'        SSTab1.TabEnabled(1) = False
'        SSTab1.TabEnabled(2) = False
'        If txt_fob_me1.Enabled = False Then
'            txt_fob_me1.Enabled = True
''            txt_fob_me1.SetFocus
''        Else
''            txt_fob_me1.SetFocus
'        End If
'        cmd_dec.SetFocus
'        txt_local_IT_bs.Text = "0.0309"
'        txt_local_IVA_bs.Text = "0.1491"
'        txt_cge_IT_bs = "0.0416"
'        txt_cge_IVA_bs = "0.16"
'        txt_tac_billing_bs = "0.035"
'    '    BtnVer.Visible = True
'        'dtc_codigo9.Enabled = False
'      Else
'        MsgBox "El registro NO corresponde al continente: AMERICA, verifique por favor ...", vbExclamation, "Validación de Registro"
'      End If
'   Else
'      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
'   End If
'
'  Exit Sub
'
'EditErr:
'  MsgBox Err.Description
'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    If IsNull(Ado_datos.Recordset!cotiza_nro_montador) Then            'Txt_campo5A.Text = "" Then
        MsgBox "Debe registrar el Número de Montadores, verifique por favor y vuelva a intentar...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    If Ado_datos.Recordset!pais_continente <> "AMERICA" Then
        MsgBox "El registro NO corresponde al continente: AMERICA, verifique por favor y vuelva a intentar ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    'swnuevo = 2
    fraOpciones1.Visible = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Visible = False
    'FraModeloCostoE.Enabled = False
    'Fra_datos2.Enabled = False
'    fraOpciones2.Visible = False
'    FraNavega1.Enabled = False
    VAR_SW = "MOD"
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        'Call ABRIR_TABLA_DET
        VAR_CONTI = "AMERICA"
        aw_p_ao_solicitud_cotiza_costos.txt_conti = VAR_CONTI
        aw_p_ao_solicitud_cotiza_costos.txt_codigo.Caption = Me.Ado_datos.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
        GlSolicitud = Me.Ado_datos.Recordset("solicitud_codigo") ' Nro. Tramite (Cod.solicitud)
        aw_p_ao_solicitud_cotiza_costos.txt_campo1.Caption = parametro 'Me.Ado_datosE.Recordset("unidad_codigo")    ' Codigo Unidad
        aw_p_ao_solicitud_cotiza_costos.Txt_descripcion.Caption = Me.txt_campo12.Text                         ' Descripcion Unidad
        aw_p_ao_solicitud_cotiza_costos.Txt_Correl.Caption = Me.Ado_datos.Recordset("cotiza_codigo")    ' Nro. Cotización
        aw_p_ao_solicitud_cotiza_costos.Txt_campo2.Caption = Me.Ado_datos.Recordset("edif_codigo")      ' Codigo Edificio
        aw_p_ao_solicitud_cotiza_costos.txt_pais.Caption = Me.Ado_datos.Recordset("pais_codigo")      ' Pais
        aw_p_ao_solicitud_cotiza_costos.txt_campo5.Text = Me.Ado_datos.Recordset!cotiza_nro_montador     ' #Montadores
        aw_p_ao_solicitud_cotiza_costos.txt_paradas = dtc_desc10    'paradas
        aw_p_ao_solicitud_cotiza_costos.cmd_dec.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_dec), "2", Me.Ado_datos.Recordset!cotiza_dec)               ' NRO.Decimales
        aw_p_ao_solicitud_cotiza_costos.cmd_moneda.Text = IIf(IsNull(Me.Ado_datos.Recordset!tipo_moneda), "BRL", Me.Ado_datos.Recordset!tipo_moneda)        ' Tipo de Moneda
        aw_p_ao_solicitud_cotiza_costos.Txt_tdc.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_tdc_bol), "6.96", Me.Ado_datos.Recordset!cotiza_tdc_bol)    ' Tipo de Cambio
        aw_p_ao_solicitud_cotiza_costos.txt_montobase.Text = IIf(IsNull(Me.Ado_datos.Recordset!costo_monto), "0", Me.Ado_datos.Recordset!costo_monto)       ' Monto Moneda Base
        
        aw_p_ao_solicitud_cotiza_costos.txt_fob_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_fob_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_fob_dol)  ' FOB ME
        aw_p_ao_solicitud_cotiza_costos.txt_fob_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_fob_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_fob_bs)    ' FOB BOB
        
        aw_p_ao_solicitud_cotiza_costos.txt_dcto_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_dcto_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_dcto_dol)   ' Dcto ME
        aw_p_ao_solicitud_cotiza_costos.txt_dcto_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_dcto_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_dcto_bs)     ' Dcto Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_seguro_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_seg_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_seg_dol)   ' Seguro ME
        aw_p_ao_solicitud_cotiza_costos.txt_seguro_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_seg_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_seg_bs)     ' Seguro Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_fob_seg_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_fob_seg_dol), "0", Me.Ado_datos.Recordset!cotiza_fob_seg_dol)   ' FOB+Seguro ME
        aw_p_ao_solicitud_cotiza_costos.txt_fob_seg_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_fob_seg_bs), "0", Me.Ado_datos.Recordset!cotiza_fob_seg_bs)      'FOB+Seguro Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_fletefrontera_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_flete_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_flete_dol)    'Flete ME
        aw_p_ao_solicitud_cotiza_costos.txt_fletefrontera_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_flete_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_flete_bs)      'Flete Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_cif_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_cif_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_cif_dol) ' CIF ME
        aw_p_ao_solicitud_cotiza_costos.txt_cif_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_cif_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_cif_bs) ' CIF Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_gastos_locales_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_gasto_local_dol), "0", Me.Ado_datos.Recordset!cotiza_gasto_local_dol) ' GastoLocal ME
        aw_p_ao_solicitud_cotiza_costos.txt_gastos_locales_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_gasto_local_bs), "0", Me.Ado_datos.Recordset!cotiza_gasto_local_bs) ' GastoLocal Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_total_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_dol), "0", Me.Ado_datos.Recordset!cotiza_precio_total_dol) ' SUB-TOTAL ME
        aw_p_ao_solicitud_cotiza_costos.txt_total_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_bs), "0", Me.Ado_datos.Recordset!cotiza_precio_total_bs) ' SUB-TOTAL Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_local_IT_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_local_IT_dol), "0", Me.Ado_datos.Recordset!cotiza_saldo_local_IT_dol) ' Local-IT ME
        aw_p_ao_solicitud_cotiza_costos.txt_local_IT_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_local_IT_bs), "0", Me.Ado_datos.Recordset!cotiza_saldo_local_IT_bs) ' Local-IT Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_local_IVA_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_local_IVA_dol), "0", Me.Ado_datos.Recordset!cotiza_saldo_local_IVA_dol) ' Local-IVA ME
        aw_p_ao_solicitud_cotiza_costos.txt_local_IVA_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_local_IVA_bs), "0", Me.Ado_datos.Recordset!cotiza_saldo_local_IVA_bs) ' Local-IVA Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_totalCli_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_dol_cli), "0", Me.Ado_datos.Recordset!cotiza_precio_total_dol_cli) ' TOT-Cli ME
        aw_p_ao_solicitud_cotiza_costos.txt_totalCli_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_bs_cli), "0", Me.Ado_datos.Recordset!cotiza_precio_total_bs_cli) ' TOT-Cli Bs
        
        aw_p_ao_solicitud_cotiza_costos.txt_cge_IT_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_cge_IT_dol), "0", Me.Ado_datos.Recordset!cotiza_saldo_cge_IT_dol)    ' CGE-IT ME
        aw_p_ao_solicitud_cotiza_costos.txt_cge_IT_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_cge_IT_bs), "0", Me.Ado_datos.Recordset!cotiza_saldo_cge_IT_bs)    ' CGE-IT Bs
        aw_p_ao_solicitud_cotiza_costos.txt_cge_IVA_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_cge_IVA_dol), "0", Me.Ado_datos.Recordset!cotiza_saldo_cge_IVA_dol)    ' CGE-IVA ME
        aw_p_ao_solicitud_cotiza_costos.txt_cge_IVA_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_cge_IVA_bs), "0", Me.Ado_datos.Recordset!cotiza_saldo_cge_IVA_bs)    ' CGE-IVA Bs
        aw_p_ao_solicitud_cotiza_costos.txt_tac_billing_dol.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_tac_billing_dol), "0", Me.Ado_datos.Recordset!cotiza_saldo_tac_billing_dol)    'SaldoTacBill ME
        aw_p_ao_solicitud_cotiza_costos.txt_tac_billing_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_saldo_tac_billing_bs), "0", Me.Ado_datos.Recordset!cotiza_saldo_tac_billing_bs)    'SaldoTacBill Bs
        aw_p_ao_solicitud_cotiza_costos.txt_totalCGE_me.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_dol_cge), "0", Me.Ado_datos.Recordset!cotiza_precio_total_dol_cge)    'TOT-CGE ME
        aw_p_ao_solicitud_cotiza_costos.txt_totalCGE_bs.Text = IIf(IsNull(Me.Ado_datos.Recordset!cotiza_precio_total_bs_cge), "0", Me.Ado_datos.Recordset!cotiza_precio_total_bs_cge)    'TOT-CGE Bs
        
        aw_p_ao_solicitud_cotiza_costos.Show vbModal
        
        fraOpciones1.Visible = True
        FraNavega.Enabled = True
'        FraModeloCostoE.Enabled = False
        Fra_datos2.Enabled = False
        FraDet1.Enabled = True
        FrmABMDet.Visible = True
        dg_datos.Enabled = True
'        fraOpciones2.Visible = True
'        FraNavega1.Enabled = True

        VAR_SW = ""
        SSTab1.Tab = 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
'        If VAR_CONTI = "AMERICA" Then
'            SSTab1.TabEnabled(0) = True
'        Else
'           SSTab1.TabEnabled(0) = False
'        End If
'        If VAR_CONTI = "ASIA" Then
'           SSTab1.TabEnabled(1) = True
'           If Ado_datosA.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(1) = False
'        End If
'        If VAR_CONTI = "EUROPA" Then
'           SSTab1.TabEnabled(2) = True
'           If Ado_datosE.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(2) = False
'        End If
        swnuevo = 0
        Call ABRIR_TABLA
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar1A_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
'  On Error GoTo EditErr
''  lblStatus.Caption = "Modificar registro"
'    If Ado_datosA.Recordset!estado_codigo = "REG" Then
'        If Ado_datosA.Recordset!pais_continente = "ASIA" Then
'            If IsNull(Ado_datosA.Recordset!cotiza_nro_montador) Then            'Txt_campo5A.Text = "" Then
'                MsgBox "Debe registrar el Número de Montadores, verifique por favor y vuelva a intentar...", vbExclamation, "Validación de Registro"
'                Exit Sub
'            End If
'    '        Fra_datos.Enabled = True
''            FraModeloCostoA.Visible = True
''            FraModeloCostoA.Enabled = True
''            FraGrabarCancelarA.Visible = True
'            Fra_datos2.Enabled = True
'            'fraOpciones.Enabled = False
'            fraOpciones1A.Visible = False
''            fraOpciones2A.Visible = False
'            FrmABMDet.Visible = False
'            FraDet1.Enabled = False
'            dg_datosA.Enabled = False
''            dg_datos1A.Enabled = False
'            VAR_SW = "MOD"
'            cmd_decA.SetFocus
'            SSTab1.Tab = 1
'            SSTab1.TabEnabled(0) = False
'            SSTab1.TabEnabled(1) = True
'            SSTab1.TabEnabled(2) = False
'            txt_local_IT_bsA.Text = "0.0309"
'            txt_local_IVA_bsA.Text = "0.1491"
'            txt_cge_IT_bsA = "0.0416"
'            txt_cge_IVA_bsA = "0.16"
'            txt_tac_billing_bsA = "0.035"
'            txt_tacb1 = "0.035"
'            txt_spread1 = "0.08"
'            txt_gac_bs = "0.05"
'    '        aw_p_ao_solicitud_cotiza_costosE.Show vbModal
'        '    BtnVer.Visible = True
'            'dtc_codigo9.Enabled = False
'      Else
'        MsgBox "El registro NO corresponde al continente: ASIA, verifique por favor ...", vbExclamation, "Validación de Registro"
'      End If
'    Else
'      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
'    End If
'  Exit Sub
'
'EditErr:
'  MsgBox Err.Description
If rs_datosA.RecordCount > 0 Then
  If rs_datosA!estado_codigo = "REG" Then
    If IsNull(Ado_datosA.Recordset!cotiza_nro_montador) Then            'Txt_campo5A.Text = "" Then
        MsgBox "Debe registrar el Número de Montadores, verifique por favor y vuelva a intentar...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    If Ado_datosA.Recordset!pais_continente <> "ASIA" Then
        MsgBox "El registro NO corresponde al continente: ASIA, verifique por favor y vuelva a intentar ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    'swnuevo = 2
    fraOpciones1A.Visible = False
    FraNavegaA.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Visible = False
    'FraModeloCostoE.Enabled = False
    Fra_datos2.Enabled = False
    'fraOpciones2E.Visible = False
    'FraNavega1E.Enabled = False
    VAR_SW = "MOD"
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
        'Call ABRIR_TABLA_DET
        VAR_CONTI = "ASIA"
        aw_p_ao_solicitud_cotiza_costosA.txt_conti.Caption = VAR_CONTI
        aw_p_ao_solicitud_cotiza_costosA.txt_codigo.Caption = Me.Ado_datosA.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
        aw_p_ao_solicitud_cotiza_costosA.txt_campo1.Caption = parametro 'Me.Ado_datosA.Recordset("unidad_codigo")    ' Codigo Unidad
        aw_p_ao_solicitud_cotiza_costosA.Txt_descripcion.Caption = Me.txt_campo12                        ' Descripcion Unidad
        aw_p_ao_solicitud_cotiza_costosA.Txt_Correl.Caption = Me.Ado_datosA.Recordset("cotiza_codigo")    ' Nro. Cotización
        aw_p_ao_solicitud_cotiza_costosA.Txt_campo2.Caption = Me.Ado_datosA.Recordset("edif_codigo")      ' Codigo Edificio
        aw_p_ao_solicitud_cotiza_costosA.txt_pais.Caption = Me.Ado_datosA.Recordset("pais_codigo")      ' Pais
        aw_p_ao_solicitud_cotiza_costosA.txt_campo5.Text = Me.Ado_datosA.Recordset!cotiza_nro_montador     ' #Montadores
        aw_p_ao_solicitud_cotiza_costosA.txt_paradas = dtc_desc10.Caption    'paradas

        aw_p_ao_solicitud_cotiza_costosA.cmd_dec.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_dec), "2", Me.Ado_datosA.Recordset!cotiza_dec) ' NRO.Decimales
        aw_p_ao_solicitud_cotiza_costosA.cmd_moneda.Text = IIf(IsNull(Me.Ado_datosA.Recordset!tipo_moneda), "EUR", Me.Ado_datosA.Recordset!tipo_moneda) ' Tipo de Moneda
        aw_p_ao_solicitud_cotiza_costosA.Txt_tdc.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_tdc_bol), "6.96", Me.Ado_datosA.Recordset!cotiza_tdc_bol) ' Tipo de Cambio
        aw_p_ao_solicitud_cotiza_costosA.txt_montobase.Text = IIf(IsNull(Me.Ado_datosA.Recordset!costo_monto), "0", Me.Ado_datosA.Recordset!costo_monto) ' Monto Moneda Base
        
        aw_p_ao_solicitud_cotiza_costosA.txt_fob_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_fob_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_fob_dol) ' FOB ME
        aw_p_ao_solicitud_cotiza_costosA.txt_fob_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_fob_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_fob_bs) ' FOB BOB

        aw_p_ao_solicitud_cotiza_costosA.txt_dcto_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_dcto_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_dcto_dol) ' Dcto ME
        aw_p_ao_solicitud_cotiza_costosA.txt_dcto_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_dcto_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_dcto_bs) ' Dcto Bs

        aw_p_ao_solicitud_cotiza_costosA.txt_tacb_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_tacb_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_tacb_dol) ' TacBill ME
        aw_p_ao_solicitud_cotiza_costosA.txt_tacb1.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_tacb_bs), "0.035", Me.Ado_datosA.Recordset!cotiza_precio_tacb_bs)  ' TacBill param

        aw_p_ao_solicitud_cotiza_costosA.txt_spread_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_spread_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_spread_dol)  ' Spread ME
        aw_p_ao_solicitud_cotiza_costosA.txt_spread1.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_spread_bs), "0.08", Me.Ado_datosA.Recordset!cotiza_precio_spread_bs)   ' Spread Param

        aw_p_ao_solicitud_cotiza_costosA.txt_seguro_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_seg_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_seg_dol)    ' Seguro ME
        aw_p_ao_solicitud_cotiza_costosA.txt_seguro_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_seg_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_seg_bs)    ' Seguro Bs

        aw_p_ao_solicitud_cotiza_costosA.txt_fob_seg_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_fob_seg_dol), "0", Me.Ado_datosA.Recordset!cotiza_fob_seg_dol)    ' FOB+Seguro ME
        aw_p_ao_solicitud_cotiza_costosA.txt_fob_seg_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_fob_seg_bs), "0", Me.Ado_datosA.Recordset!cotiza_fob_seg_bs)    'FOB+Seguro Bs

        aw_p_ao_solicitud_cotiza_costosA.txt_fletefrontera_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_flete_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_flete_dol)    'Flete ME
        aw_p_ao_solicitud_cotiza_costosA.txt_fletefrontera_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_flete_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_flete_bs)    'Flete Bs
        
        aw_p_ao_solicitud_cotiza_costosA.txt_cif_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_cif_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_cif_dol) ' CIF ME
        aw_p_ao_solicitud_cotiza_costosA.txt_cif_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_cif_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_cif_bs) ' CIF Bs

        aw_p_ao_solicitud_cotiza_costosA.txt_GAC_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_GAC_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_GAC_dol)    ' GAC ME
        aw_p_ao_solicitud_cotiza_costosA.txt_gac_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_GAC_bs), "0.05", Me.Ado_datosA.Recordset!cotiza_precio_GAC_bs)   ' GAC param
        
        aw_p_ao_solicitud_cotiza_costosA.txt_base_imp_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_base_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_base_dol)   ' Base Imponible Usd
        aw_p_ao_solicitud_cotiza_costosA.txt_base_imp_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_base_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_base_bs)      ' Base Imponible Bs
        
        aw_p_ao_solicitud_cotiza_costosA.txt_gastos_locales_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_gasto_local_dol), "0", Me.Ado_datosA.Recordset!cotiza_gasto_local_dol) ' GastoLocal ME
        aw_p_ao_solicitud_cotiza_costosA.txt_gastos_locales_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_gasto_local_bs), "0", Me.Ado_datosA.Recordset!cotiza_gasto_local_bs)    ' GastoLocal Bs
        
        aw_p_ao_solicitud_cotiza_costosA.txt_total_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_dol), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_dol) ' SUB-TOTAL ME
        aw_p_ao_solicitud_cotiza_costosA.txt_total_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_bs), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_bs)   ' SUB-TOTAL Bs
        
        aw_p_ao_solicitud_cotiza_costosA.txt_local_IT_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_local_IT_dol), "0", Me.Ado_datosA.Recordset!cotiza_saldo_local_IT_dol) ' Local-IT ME
        aw_p_ao_solicitud_cotiza_costosA.txt_local_IT_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_local_IT_bs), "0.0309", Me.Ado_datosA.Recordset!cotiza_saldo_local_IT_bs)    ' Local-IT Bs
        aw_p_ao_solicitud_cotiza_costosA.txt_local_IVA_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_local_IVA_dol), "0", Me.Ado_datosA.Recordset!cotiza_saldo_local_IVA_dol)  ' Local-IVA ME
        aw_p_ao_solicitud_cotiza_costosA.txt_local_IVA_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_local_IVA_bs), "0.1491", Me.Ado_datosA.Recordset!cotiza_saldo_local_IVA_bs)     ' Local-IVA Bs
        aw_p_ao_solicitud_cotiza_costosA.txt_totalCli_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_dol_cli), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_dol_cli)  ' TOT-Cli ME
        aw_p_ao_solicitud_cotiza_costosA.txt_totalCli_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_bs_cli), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_bs_cli)    ' TOT-Cli Bs
        
        aw_p_ao_solicitud_cotiza_costosA.txt_cge_IT_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_cge_IT_dol), "0", Me.Ado_datosA.Recordset!cotiza_saldo_cge_IT_dol)    ' CGE-IT ME
        aw_p_ao_solicitud_cotiza_costosA.txt_cge_IT_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_cge_IT_bs), "0.0416", Me.Ado_datosA.Recordset!cotiza_saldo_cge_IT_bs)    ' CGE-IT Bs
        aw_p_ao_solicitud_cotiza_costosA.txt_cge_IVA_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_cge_IVA_dol), "0", Me.Ado_datosA.Recordset!cotiza_saldo_cge_IVA_dol)    ' CGE-IVA ME
        aw_p_ao_solicitud_cotiza_costosA.txt_cge_IVA_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_cge_IVA_bs), "0.16", Me.Ado_datosA.Recordset!cotiza_saldo_cge_IVA_bs)    ' CGE-IVA Bs
        aw_p_ao_solicitud_cotiza_costosA.txt_tac_billing_dol.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_tac_billing_dol), "0", Me.Ado_datosA.Recordset!cotiza_saldo_tac_billing_dol)    'SaldoTacBill ME
        aw_p_ao_solicitud_cotiza_costosA.txt_tac_billing_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_saldo_tac_billing_bs), "0.035", Me.Ado_datosA.Recordset!cotiza_saldo_tac_billing_bs)    'SaldoTacBill Bs
        aw_p_ao_solicitud_cotiza_costosA.txt_totalCGE_me.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_dol_cge), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_dol_cge)    'TOT-CGE ME
        aw_p_ao_solicitud_cotiza_costosA.txt_totalCGE_bs.Text = IIf(IsNull(Me.Ado_datosA.Recordset!cotiza_precio_total_bs_cge), "0", Me.Ado_datosA.Recordset!cotiza_precio_total_bs_cge)    'TOT-CGE Bs
        
        aw_p_ao_solicitud_cotiza_costosA.Show vbModal
        fraOpciones1A.Visible = True
        FraNavegaA.Enabled = True
'        FraModeloCostoE.Enabled = False
        Fra_datos2.Enabled = False
        FraDet1.Enabled = True
        FrmABMDet.Visible = True
        dg_datosA.Enabled = True

        VAR_SW = ""
        SSTab1.Tab = 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = False
'        If VAR_CONTI = "AMERICA" Then
'            SSTab1.TabEnabled(0) = True
'        Else
'           SSTab1.TabEnabled(0) = False
'        End If
'        If VAR_CONTI = "ASIA" Then
'           SSTab1.TabEnabled(1) = True
'           If Ado_datosA.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(1) = False
'        End If
'        If VAR_CONTI = "EUROPA" Then
'           SSTab1.TabEnabled(2) = True
'           If Ado_datosA.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(2) = False
'        End If
        swnuevo = 0
        Call ABRIR_TABLA
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
Else
    MsgBox "No se puede Modificar el registro, debe completar los Datos para Cotización !! ", vbExclamation
End If
End Sub

Private Sub BtnModificar1E_Click()
    If glusuario = "CCRUZ" Then
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  If rs_datosE.RecordCount > 0 And rs_datosE!estado_codigo = "REG" Then
    If IsNull(Ado_datosE.Recordset!cotiza_nro_montador) Then            'Txt_campo5A.Text = "" Then
        MsgBox "Debe registrar el Número de Montadores, verifique por favor y vuelva a intentar...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    If Ado_datosE.Recordset!pais_continente <> "EUROPA" Then
        MsgBox "El registro NO corresponde al continente: EUROPA, verifique por favor y vuelva a intentar ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    'swnuevo = 2
    fraOpciones1E.Visible = False
    FraNavegaE.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Visible = False
    'FraModeloCostoE.Enabled = False
    Fra_datos2.Enabled = False
'    fraOpciones2E.Visible = False
'    FraNavega1E.Enabled = False
    VAR_SW = "MOD"
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        'Call ABRIR_TABLA_DET
        VAR_CONTI = "EUROPA"
        aw_p_ao_solicitud_cotiza_costosE.txt_conti.Caption = VAR_CONTI
        aw_p_ao_solicitud_cotiza_costosE.txt_codigo.Caption = Me.Ado_datosE.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
        aw_p_ao_solicitud_cotiza_costosE.txt_campo1.Caption = parametro 'Me.Ado_datosE.Recordset("unidad_codigo")    ' Codigo Unidad
        aw_p_ao_solicitud_cotiza_costosE.Txt_descripcion.Caption = Me.txt_campo12                        ' Descripcion Unidad
        aw_p_ao_solicitud_cotiza_costosE.Txt_Correl.Caption = Me.Ado_datosE.Recordset("cotiza_codigo")    ' Nro. Cotización
        aw_p_ao_solicitud_cotiza_costosE.Txt_campo2.Caption = Me.Ado_datosE.Recordset("edif_codigo")      ' Codigo Edificio
        aw_p_ao_solicitud_cotiza_costosE.txt_pais.Caption = Me.Ado_datosE.Recordset("pais_codigo")      ' Pais
        aw_p_ao_solicitud_cotiza_costosE.txt_campo5.Text = Me.Ado_datosE.Recordset!cotiza_nro_montador     ' #Montadores
        aw_p_ao_solicitud_cotiza_costosE.txt_paradas = dtc_desc10.Caption    'paradas
        aw_p_ao_solicitud_cotiza_costosE.txt_tdc_me.Text = IIf(IsNull(GlTipoCambioEuro) Or GlTipoCambioEuro = 0, "10", GlTipoCambioEuro)  'Euro
        aw_p_ao_solicitud_cotiza_costosE.cmd_dec.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_dec), "2", Me.Ado_datosE.Recordset!cotiza_dec) ' NRO.Decimales
        aw_p_ao_solicitud_cotiza_costosE.cmd_moneda.Text = IIf(IsNull(Me.Ado_datosE.Recordset!tipo_moneda), "EUR", Me.Ado_datosE.Recordset!tipo_moneda) ' Tipo de Moneda
        aw_p_ao_solicitud_cotiza_costosE.Txt_tdc.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_tdc_bol), "6.96", Me.Ado_datosE.Recordset!cotiza_tdc_bol) ' Tipo de Cambio
        aw_p_ao_solicitud_cotiza_costosE.txt_montobase.Text = IIf(IsNull(Me.Ado_datosE.Recordset!costo_monto), "0", Me.Ado_datosE.Recordset!costo_monto) ' Monto Moneda Base
        
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_fob_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_fob_me) ' FOB EU
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_fob_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_fob_dol) ' FOB ME
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_fob_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_fob_bs) ' FOB BOB
        aw_p_ao_solicitud_cotiza_costosE.txt_dcto_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_dcto_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_dcto_me) ' Dcto EU
        aw_p_ao_solicitud_cotiza_costosE.txt_dcto_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_dcto_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_dcto_dol) ' Dcto ME
        aw_p_ao_solicitud_cotiza_costosE.txt_dcto_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_dcto_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_dcto_bs) ' Dcto Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_tacb_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_tacb_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_tacb_me) ' TacBill eu
        aw_p_ao_solicitud_cotiza_costosE.txt_tacb_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_tacb_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_tacb_dol) ' TacBill ME
        aw_p_ao_solicitud_cotiza_costosE.txt_tacb_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_tacb_bs), "0.02", Me.Ado_datosE.Recordset!cotiza_precio_tacb_bs) ' TacBill Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_spread_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_spread_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_spread_me) ' Spread eu
        aw_p_ao_solicitud_cotiza_costosE.txt_spread_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_spread_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_spread_dol) ' Spread ME
        aw_p_ao_solicitud_cotiza_costosE.txt_spread_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_spread_bs), "0.02", Me.Ado_datosE.Recordset!cotiza_precio_spread_bs)    ' Spread Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_seguro_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_seg_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_seg_me)    ' Seguro EU
        aw_p_ao_solicitud_cotiza_costosE.txt_seguro_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_seg_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_seg_dol)    ' Seguro ME
        aw_p_ao_solicitud_cotiza_costosE.txt_seguro_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_seg_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_seg_bs)    ' Seguro Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_seg_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_fob_seg_me), "0", Me.Ado_datosE.Recordset!cotiza_fob_seg_me)    ' FOB+Seguro EU
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_seg_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_fob_seg_dol), "0", Me.Ado_datosE.Recordset!cotiza_fob_seg_dol)    ' FOB+Seguro ME
        aw_p_ao_solicitud_cotiza_costosE.txt_fob_seg_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_fob_seg_bs), "0", Me.Ado_datosE.Recordset!cotiza_fob_seg_bs)    'FOB+Seguro Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_fletefrontera_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_flete_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_flete_me)    'Flete EU
        aw_p_ao_solicitud_cotiza_costosE.txt_fletefrontera_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_flete_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_flete_dol)    'Flete ME
        aw_p_ao_solicitud_cotiza_costosE.txt_fletefrontera_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_flete_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_flete_bs)    'Flete Bs
        
        aw_p_ao_solicitud_cotiza_costosE.txt_cif_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_cif_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_cif_me) ' CIF EU
        aw_p_ao_solicitud_cotiza_costosE.txt_cif_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_cif_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_cif_dol) ' CIF ME
        aw_p_ao_solicitud_cotiza_costosE.txt_cif_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_cif_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_cif_bs) ' CIF Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_gastos_locales_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_gasto_local_me), "0", Me.Ado_datosE.Recordset!cotiza_gasto_local_me) ' GastoLocal EU
        aw_p_ao_solicitud_cotiza_costosE.txt_gastos_locales_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_gasto_local_dol), "0", Me.Ado_datosE.Recordset!cotiza_gasto_local_dol) ' GastoLocal ME
        aw_p_ao_solicitud_cotiza_costosE.txt_gastos_locales_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_gasto_local_bs), "0", Me.Ado_datosE.Recordset!cotiza_gasto_local_bs) ' GastoLocal Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_total_eu.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_me), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_me) ' SUB-TOTAL EU
        aw_p_ao_solicitud_cotiza_costosE.txt_total_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_dol), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_dol) ' SUB-TOTAL ME
        aw_p_ao_solicitud_cotiza_costosE.txt_total_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_bs), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_bs) ' SUB-TOTAL Bs
        
        aw_p_ao_solicitud_cotiza_costosE.txt_local_IT_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_local_IT_dol), "0", Me.Ado_datosE.Recordset!cotiza_saldo_local_IT_dol) ' Local-IT ME
        aw_p_ao_solicitud_cotiza_costosE.txt_local_IT_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_local_IT_bs), "0", Me.Ado_datosE.Recordset!cotiza_saldo_local_IT_bs) ' Local-IT Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_local_IVA_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_local_IVA_dol), "0", Me.Ado_datosE.Recordset!cotiza_saldo_local_IVA_dol) ' Local-IVA ME
        aw_p_ao_solicitud_cotiza_costosE.txt_local_IVA_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_local_IVA_bs), "0", Me.Ado_datosE.Recordset!cotiza_saldo_local_IVA_bs) ' Local-IVA Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_totalCli_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_dol_cli), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_dol_cli) ' TOT-Cli ME
        aw_p_ao_solicitud_cotiza_costosE.txt_totalCli_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_bs_cli), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_bs_cli) ' TOT-Cli Bs
        
        aw_p_ao_solicitud_cotiza_costosE.txt_cge_IT_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_cge_IT_dol), "0", Me.Ado_datosE.Recordset!cotiza_saldo_cge_IT_dol)    ' CGE-IT ME
        aw_p_ao_solicitud_cotiza_costosE.txt_cge_IT_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_cge_IT_bs), "0", Me.Ado_datosE.Recordset!cotiza_saldo_cge_IT_bs)    ' CGE-IT Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_cge_IVA_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_cge_IVA_dol), "0", Me.Ado_datosE.Recordset!cotiza_saldo_cge_IVA_dol)    ' CGE-IVA ME
        aw_p_ao_solicitud_cotiza_costosE.txt_cge_IVA_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_cge_IVA_bs), "0", Me.Ado_datosE.Recordset!cotiza_saldo_cge_IVA_bs)    ' CGE-IVA Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_tac_billing_dol.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_tac_billing_dol), "0", Me.Ado_datosE.Recordset!cotiza_saldo_tac_billing_dol)    'SaldoTacBill ME
        aw_p_ao_solicitud_cotiza_costosE.txt_tac_billing_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_saldo_tac_billing_bs), "0", Me.Ado_datosE.Recordset!cotiza_saldo_tac_billing_bs)    'SaldoTacBill Bs
        aw_p_ao_solicitud_cotiza_costosE.txt_totalCGE_me.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_dol_cge), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_dol_cge)    'TOT-CGE ME
        aw_p_ao_solicitud_cotiza_costosE.txt_totalCGE_bs.Text = IIf(IsNull(Me.Ado_datosE.Recordset!cotiza_precio_total_bs_cge), "0", Me.Ado_datosE.Recordset!cotiza_precio_total_bs_cge)    'TOT-CGE Bs
        
        aw_p_ao_solicitud_cotiza_costosE.Show vbModal
        
        fraOpciones1E.Visible = True
        FraNavegaE.Enabled = True
'        FraModeloCostoE.Enabled = False
        Fra_datos2.Enabled = False
        FraDet1.Enabled = True
        FrmABMDet.Visible = True
        dg_datosE.Enabled = True
'        fraOpciones2E.Visible = True
'        FraNavega1E.Enabled = True

        VAR_SW = ""
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
'        If VAR_CONTI = "AMERICA" Then
'            SSTab1.TabEnabled(0) = True
'        Else
'           SSTab1.TabEnabled(0) = False
'        End If
'        If VAR_CONTI = "ASIA" Then
'           SSTab1.TabEnabled(1) = True
'           If Ado_datosA.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(1) = False
'        End If
'        If VAR_CONTI = "EUROPA" Then
'           SSTab1.TabEnabled(2) = True
'           If Ado_datosE.Recordset.RecordCount > 0 Then
'             Call ABRIR_TABLA_DET
'           End If
'        Else
'           SSTab1.TabEnabled(2) = False
'        End If
        swnuevo = 0
        Call ABRIR_TABLA
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificarA_Click()
  If Ado_datos0.Recordset.RecordCount > 0 Then
    If Ado_datos0.Recordset!estado_codigo = "REG" Then
        If Ado_datos0.Recordset!pais_continente_cot = "ASIA" Then
            'swnuevo = 2
            fraOpciones1A.Visible = False
            FraNavegaA.Enabled = False
            FraDet1.Enabled = False
            FrmABMDet.Visible = False
         '   Fra_datos.Enabled = False
'            FraModeloCostoA.Enabled = False
            Fra_datos2.Enabled = False
            VAR_SW = "MOD"
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
        '    Select Case dtc_codigo2.Text
        '        Case "1"
        '        Case "2"
        '        Case "3"
            'Call ABRIR_TABLA_DET
            VAR_CONTI = Ado_datos0.Recordset!pais_continente_cot
            GlConti = VAR_CONTI
            GlSolicitud = Me.Ado_datos0.Recordset!solicitud_codigo           ' Nro. Negociacion (Cod.solicitud)
            GlUnidad = Me.Ado_datos0.Recordset!unidad_codigo                 ' Codigo Unidad
            GlNombFor = Me.txt_campo12                                          ' Descripcion Unidad
            GlCotiza = Me.Ado_datos0.Recordset!cotiza_codigo                ' Nro. Cotización
            GlEdificio = Me.Ado_datos0.Recordset!EDIF_CODIGO                ' Codigo Edificio
            aw_solicitud_cotiza_datos.Show vbModal
            
            VAR_CONTI = "ASIA"
'            frm_ao_solicitud_cotiza_datosA.txt_conti.Caption = VAR_CONTI
'            frm_ao_solicitud_cotiza_datosA.txt_codigo.Caption = Me.Ado_datos0.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'            frm_ao_solicitud_cotiza_datosA.Txt_campo1.Caption = Me.Ado_datos0.Recordset("unidad_codigo")    ' Codigo Unidad
'            frm_ao_solicitud_cotiza_datosA.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'            frm_ao_solicitud_cotiza_datosA.Txt_Correl.Caption = Me.Ado_datos0.Recordset("cotiza_codigo")    ' Nro. Cotización
'            frm_ao_solicitud_cotiza_datosA.Txt_campo2A.Caption = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'            GlEdificio = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'            frm_ao_solicitud_cotiza_datosA.Txt_campo4.Text = Me.Ado_datos0.Recordset("modelo_codigo") ' Modelo
'
'            frm_ao_solicitud_cotiza_datosA.Txt_campo5.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_nro_montador), "2", Me.Ado_datos0.Recordset!cotiza_nro_montador) ' Montadores
'            frm_ao_solicitud_cotiza_datosA.Txt_campo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_energia), "0", Me.Ado_datos0.Recordset!cotiza_energia) ' Energia
'            frm_ao_solicitud_cotiza_datosA.Txt_campo3.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_luz), "0", Me.Ado_datos0.Recordset!cotiza_luz) ' Luz
'            frm_ao_solicitud_cotiza_datosA.Txt_campo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!bien_cotiza_num_accesos), "0", Me.Ado_datos0.Recordset!bien_cotiza_num_accesos) ' Num Accesos
'            frm_ao_solicitud_cotiza_datosA.Txt_campo9.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_fondo), "0", Me.Ado_datos0.Recordset!dimension_fosa_fondo) ' Fosa fondo
'            frm_ao_solicitud_cotiza_datosA.Txt_campo10.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_frente), "0", Me.Ado_datos0.Recordset!dimension_fosa_frente) ' Fosa Frente
'            frm_ao_solicitud_cotiza_datosA.Txt_campo8.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_m), "0", Me.Ado_datos0.Recordset!dimension_fosa_m) ' Espacio Dintel
'            'Equipo
'            frm_ao_solicitud_cotiza_datosA.dtc_codigo21.Text = IIf(IsNull(Me.Ado_datos0.Recordset!bien_codigo), "NA2", Me.Ado_datos0.Recordset!bien_codigo)     ' Codigo Equipo
'            frm_ao_solicitud_cotiza_datosA.dtc_desc24.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo21.BoundText
'            frm_ao_solicitud_cotiza_datosA.dtc_desc21.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo21.BoundText
'            'Pais
'            frm_ao_solicitud_cotiza_datosA.dtc_codigo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!pais_codigo), "CHN", Me.Ado_datos0.Recordset!pais_codigo)    ' Pais
'            frm_ao_solicitud_cotiza_datosA.dtc_desc7.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo7.BoundText
'            'Tipo de Equipo
'            frm_ao_solicitud_cotiza_datosA.dtc_codigo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!tipo_eqp), "A", Me.Ado_datos0.Recordset!tipo_eqp)    ' Tipo Equipo
'            frm_ao_solicitud_cotiza_datosA.dtc_desc2.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo2.BoundText
'            'Cuarto de Control
'            frm_ao_solicitud_cotiza_datosA.dtc_codigo61.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cuadro_ctrl_codigo), "1", Me.Ado_datos0.Recordset!cuadro_ctrl_codigo)    'Cuarto de Control
'            frm_ao_solicitud_cotiza_datosA.dtc_desc61.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo61.BoundText
'            'Marcas
'            frm_ao_solicitud_cotiza_datosA.dtc_codigo3.Text = IIf(IsNull(Me.Ado_datos0.Recordset!marca_codigo), "S/M", Me.Ado_datos0.Recordset!marca_codigo)    'Marca
'            frm_ao_solicitud_cotiza_datosA.dtc_desc3.BoundText = frm_ao_solicitud_cotiza_datosA.dtc_codigo3.BoundText
'
'            frm_ao_solicitud_cotiza_datosA.Show vbModal
'    '        Case "4"
'    '
'    '    End Select
            fraOpciones1A.Visible = True
            FraNavegaA.Enabled = True
'            FraModeloCostoA.Enabled = False
            Fra_datos2.Enabled = False
            FraDet1.Enabled = True
            FrmABMDet.Visible = True
            dg_datosA.Enabled = True
            VAR_SW = ""
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            swnuevo = 0
            Call ABRIR_TABLA
        Else
            MsgBox "El registro NO corresponde al continente: ASIA, verifique por favor ...", vbExclamation, "Validación de Registro"
        End If
    Else
      MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
    End If
  Else
    MsgBox "No existe el Registro para Modificar, Vuelva a intentar...!! ", vbExclamation
  End If

End Sub

Private Sub BtnModificarE_Click()
  If rs_datos0.RecordCount > 0 And rs_datos0!estado_codigo = "REG" Then
    If Ado_datos0.Recordset!pais_continente_cot <> "EUROPA" Then
        MsgBox "El registro NO corresponde al continente: EUROPA, verifique por favor y vuelva a intentar ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    'swnuevo = 2
    fraOpciones1E.Visible = False
    FraNavegaE.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Visible = False
    'FraModeloCostoE.Enabled = False
    Fra_datos2.Enabled = False
    VAR_SW = "MOD"
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
'    Select Case dtc_codigo2.Text
'        Case "1"
'        Case "2"
'        Case "3"
        'Call ABRIR_TABLA_DET
         VAR_CONTI = Ado_datos0.Recordset!pais_continente_cot
            GlConti = VAR_CONTI
            GlSolicitud = Me.Ado_datos0.Recordset!solicitud_codigo           ' Nro. Negociacion (Cod.solicitud)
            GlUnidad = Me.Ado_datos0.Recordset!unidad_codigo                 ' Codigo Unidad
            GlNombFor = Me.txt_campo12                                          ' Descripcion Unidad
            GlCotiza = Me.Ado_datos0.Recordset!cotiza_codigo                ' Nro. Cotización
            GlEdificio = Me.Ado_datos0.Recordset!EDIF_CODIGO                ' Codigo Edificio
            aw_solicitud_cotiza_datos.Show vbModal
            
        VAR_CONTI = "EUROPA"
'        frm_ao_solicitud_cotiza_datosE.txt_conti.Caption = "EUROPA"
'        frm_ao_solicitud_cotiza_datosE.txt_codigo.Caption = Me.Ado_datos0.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
'        frm_ao_solicitud_cotiza_datosE.Txt_campo1.Caption = Me.Ado_datos0.Recordset("unidad_codigo")    ' Codigo Unidad
'        frm_ao_solicitud_cotiza_datosE.Txt_descripcion.Caption = Me.Txt_campo12                        ' Descripcion Unidad
'        frm_ao_solicitud_cotiza_datosE.Txt_Correl.Caption = Me.Ado_datos0.Recordset("cotiza_codigo")    ' Nro. Cotización
'        frm_ao_solicitud_cotiza_datosE.Txt_campo2A.Caption = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'        GlEdificio = Me.Ado_datos0.Recordset("edif_codigo")      ' Codigo Edificio
'        frm_ao_solicitud_cotiza_datosE.Txt_campo4.Text = Me.Ado_datos0.Recordset("modelo_codigo") ' Modelo
'        frm_ao_solicitud_cotiza_datosE.Txt_campo5.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_nro_montador), "2", Me.Ado_datos0.Recordset!cotiza_nro_montador) ' Montadores
'        frm_ao_solicitud_cotiza_datosE.Txt_campo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_energia), "0", Me.Ado_datos0.Recordset!cotiza_energia) ' Energia
'        frm_ao_solicitud_cotiza_datosE.Txt_campo3.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cotiza_luz), "0", Me.Ado_datos0.Recordset!cotiza_luz) ' Luz
'        frm_ao_solicitud_cotiza_datosE.Txt_campo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!bien_cotiza_num_accesos), "0", Me.Ado_datos0.Recordset!bien_cotiza_num_accesos) ' Num Accesos
'        frm_ao_solicitud_cotiza_datosE.Txt_campo9.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_fondo), "0", Me.Ado_datos0.Recordset!dimension_fosa_fondo) ' Fosa fondo
'        frm_ao_solicitud_cotiza_datosE.Txt_campo10.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_frente), "0", Me.Ado_datos0.Recordset!dimension_fosa_frente) ' Fosa Frente
'        frm_ao_solicitud_cotiza_datosE.Txt_campo8.Text = IIf(IsNull(Me.Ado_datos0.Recordset!dimension_fosa_m), "0", Me.Ado_datos0.Recordset!dimension_fosa_m) ' Espacio Dintel
'        'Equipo dtc_codigo21
'        frm_ao_solicitud_cotiza_datosE.dtc_codigo21.Text = Me.Ado_datos0.Recordset("bien_codigo")     ' Codigo Equipo
'        frm_ao_solicitud_cotiza_datosE.dtc_desc24.BoundText = frm_ao_solicitud_cotiza_datosE.dtc_codigo21.BoundText
'        frm_ao_solicitud_cotiza_datosE.dtc_desc21.BoundText = frm_ao_solicitud_cotiza_datosE.dtc_codigo21.BoundText
'
'        'Pais
'        frm_ao_solicitud_cotiza_datosE.dtc_codigo7.Text = IIf(IsNull(Me.Ado_datos0.Recordset!pais_codigo), "NN", Me.Ado_datos0.Recordset!pais_codigo)    ' Pais
'        frm_ao_solicitud_cotiza_datosE.dtc_desc7.BoundText = frm_ao_solicitud_cotiza_datosE.dtc_codigo7.BoundText
'        'Tipo de Equipo
'        'frm_ao_solicitud_cotiza_datosE.dtc_codigo2.Text = Me.Ado_datos0.Recordset("tipo_eqp")    ' Tipo Equipo
'        frm_ao_solicitud_cotiza_datosE.dtc_codigo2.Text = IIf(IsNull(Me.Ado_datos0.Recordset!tipo_eqp), "X", Me.Ado_datos0.Recordset("tipo_eqp"))  ' Tipo Equipo
'        frm_ao_solicitud_cotiza_datosE.dtc_desc2.BoundText = frm_ao_solicitud_cotiza_datosE.dtc_codigo2.BoundText
'        'Cuarto de Control
'        frm_ao_solicitud_cotiza_datosE.dtc_codigo61.Text = IIf(IsNull(Me.Ado_datos0.Recordset!cuadro_ctrl_codigo), "1", Me.Ado_datos0.Recordset!cuadro_ctrl_codigo)    'Cuarto de Control
'        frm_ao_solicitud_cotiza_datosE.dtc_desc61.BoundText = frm_ao_solicitud_cotiza_datosE.dtc_codigo61.BoundText
'
'        frm_ao_solicitud_cotiza_datosE.Show vbModal
''        Case "4"
''
''    End Select
        fraOpciones1E.Visible = True
        FraNavegaE.Enabled = True
'        FraModeloCostoE.Enabled = False
        Fra_datos2.Enabled = False
        FraDet1.Enabled = True
        FrmABMDet.Visible = True
        dg_datosE.Enabled = True
        VAR_SW = ""
        SSTab1.Tab = 2
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = True
        swnuevo = 0
        Call ABRIR_TABLA
        'Call ABRIR_TABLA_DET
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub


Private Sub dtc_desc64_Click(Area As Integer)
'    dtc_codigo64.BoundText = dtc_codigo64.BoundText
End Sub

Private Sub BtnSalir2_Click()
    Fra_datos2.Enabled = False
    fraOpciones.Visible = True
    FrmABMDet.Visible = True
    FraNavega0.Enabled = True
    SSTab1.Enabled = True
    Fra_datos2.Visible = False
    FraDet1.Enabled = True
End Sub

Private Sub BtnVer_Click()
    Fra_datos2.Enabled = True
    Fra_datos2.Visible = True
    fraOpciones.Visible = False
    FrmABMDet.Visible = False
    FraNavega0.Enabled = False
    SSTab1.Enabled = False
    FraDet1.Enabled = False
End Sub

Private Sub Form_Load()
'    Fra_datos.Enabled = False
'    FraModeloCosto.Enabled = False
'    FraModeloCostoA.Enabled = False
'    Fra_datos2.Enabled = False
'    dg_datos.Enabled = True
'    dg_datosA.Enabled = True
    'lbl_aux1.Visible = False

'    lbl_titulo2.Caption = lbl_titulo.Caption

    swnuevo = 0
    VAR_SW = ""
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    VAR_UORIGEN = Aux
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            Aux = "DCOMB"
            VAR_DPTO = "3"
        Case "1.7"    'Santa Cruz
            Aux = "DCOMS"
            VAR_DPTO = "7"
        Case "1.2"    'La Paz - Comercial
            Aux = "DVTA"
            VAR_DPTO = "2"
        Case "1.9"    ' Chuquisaca
            Aux = "DCOMC"
            VAR_DPTO = "1"
        Case "1.3"    'La Paz - Modernizacion
            Aux = "DNMOD"
            VAR_DPTO = "2"
        Case "0"    ' TODO
            If glusuario = "ASANTIVAÑEZ" Then
                Aux = "DNMOD"
                VAR_DPTO = "2"
            Else
                Aux = "DVTA"
                VAR_DPTO = "0"
            End If
            'Aux = "DVTA"
            'VAR_DPTO = "2"
     End Select
    If parametro = "" Then
        parametro = Aux     'txt_campo1.Text
    End If
    VAR_FRA = "HOJA DE COSTOS - "
    'FraNavega0.Caption = lbl_titulo.Caption
'    FraModeloCosto.Visible = False
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
'    Call ABRIR_TABLA
    Fra_datos2.Enabled = False
    VAR_PRDA = IIf(dtc_desc10 = "", 0, dtc_desc10)
''    If VAR_PAISC = "AMERICA" Then
''        FraModeloCosto.Visible = False
''        dg_datos.Enabled = True
''        SSTab1.Tab = 0
''        SSTab1.TabEnabled(0) = True
''        SSTab1.TabEnabled(1) = False
''        SSTab1.TabEnabled(2) = False
''        If Ado_datos.Recordset.RecordCount > 0 Then
''            Call ABRIR_TABLA_DET
''        End If
''    End If
''    If VAR_PAISC = "ASIA" Then
''        FraModeloCostoA.Visible = False
''        dg_datosA.Enabled = True
''        SSTab1.Tab = 1
''        SSTab1.TabEnabled(0) = False
''        SSTab1.TabEnabled(1) = True
''        SSTab1.TabEnabled(2) = False
''        If Ado_datosA.Recordset.RecordCount > 0 Then
''            Call ABRIR_TABLA_DET
''        End If
''    End If
''    If VAR_PAISC = "EUROPA" Then
'''        FraModeloCostoE.Enabled = False
''        dg_datosE.Enabled = True
''        SSTab1.Tab = 2
''        SSTab1.TabEnabled(0) = False
''        SSTab1.TabEnabled(1) = False
''        SSTab1.TabEnabled(2) = True
''        If Ado_datosE.Recordset.RecordCount > 0 Then
''            Call ABRIR_TABLA_DET
''        End If
''    End If
'
'    VAR_PRDA = IIf(dtc_desc10 = "", 0, dtc_desc10)
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        usuario2 = rs_aux6!beneficiario_codigo
        VAR_DA = rs_aux6!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos0 = New Recordset
     If rs_datos0.State = 1 Then rs_datos0.Close
     Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo_cot = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo_cot = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8'  or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3' ) )) "
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo_cot = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' ) )) "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then            'LPZ
                queryinicial = "select * From av_calculo_y_cotiza WHERE (estado_codigo_cot = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC')) "
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo_cot = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
                'queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Then
                queryinicial = "select * From av_calculo_y_cotiza WHERE (estado_codigo_cot = 'REG' AND (unidad_codigo = 'DNMOD') )"
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE (estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "'  AND left(edif_codigo,1) = '" & VAR_DPTO & "' ))"      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo_cot = 'REG' AND unidad_codigo = '" & VAR_UORIGEN & "'  AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6') )) "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC'))) "
                Else
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC'))) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND (unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC'))) "
                Else
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((estado_codigo_cot = 'REG' AND (unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC'))) "
                End If
            End If
     End Select
'    Set rs_datos0 = New Recordset
'    If rs_datos0.State = 1 Then rs_datos0.Close
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Then
'        queryinicial = "select * From av_calculo_y_cotiza WHERE estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "' "
'    Else
'        queryinicial = "select * From av_calculo_y_cotiza WHERE estado_codigo_cot = 'REG' AND unidad_codigo = '" & parametro & "' AND beneficiario_codigo_resp = '" & usuario2 & "'"
'    End If
    rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos0.Recordset = rs_datos0.DataSource
    Set dg_datos0.DataSource = Ado_datos0.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
  Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        usuario2 = rs_aux6!beneficiario_codigo
        VAR_DA = rs_aux6!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.2"
    End If
    Set rs_datos0 = New Recordset
    If rs_datos0.State = 1 Then rs_datos0.Close
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '4' ))) "
        Case "1.7"    'Santa Cruz
            If glusuario = "CURDININEA" Then        'SCZ
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' or left(edif_codigo,1) = '9' or left(edif_codigo,1) = '3' )))  "
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '8' )))  "
            End If
            
        Case "1.2"    'La Paz - Comercial
            If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then            'LPZ
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMC')) "
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE (((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9' )))) "
                'queryinicial = "select * From ao_solicitud_calculo_trafico WHERE ((estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "') OR (estado_codigo = 'REG'  AND unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '1' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' or left(edif_codigo,1) = '9'  ) )) "
            End If
        Case "1.3"    'La Paz - Modernizacion
            If glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "CCOLODRO" Then
                queryinicial = "select * From av_calculo_y_cotiza WHERE (unidad_codigo = 'DNMOD') "
            Else
                queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' ))) "      'AND beneficiario_codigo_resp2 = '" & usuario2 & "'
            End If
        Case "1.9"    ' Chuquisaca
            queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = '" & parametro & "') OR (unidad_codigo = '" & VAR_UORIGEN & "' AND (left(edif_codigo,1) = '" & VAR_DPTO & "' or left(edif_codigo,1) = '5'  or left(edif_codigo,1) = '6' )))  "
        Case "1.4"    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
        Case Else    ' ADMIN
            If glusuario = "ADMIN" Or glusuario = "VPAREDES" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Then
                If VAR_UORIGEN = "DVTA" Then
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = 'DVTA' OR unidad_codigo = 'DCOMS' OR unidad_codigo = 'DCOMB' OR unidad_codigo = 'DCOMC')) "
                    'queryinicial = "select * From ao_solicitud WHERE estado_codigo = 'REG'  "
                Else
                    queryinicial = "select * From av_calculo_y_cotiza WHERE ((unidad_codigo = 'DNMOD' OR unidad_codigo = 'DMODS' OR unidad_codigo = 'DMODB' OR unidad_codigo = 'DMODC')) "
                End If
            End If
     End Select
'    Set rs_datos0 = New Recordset
'    If rs_datos0.State = 1 Then rs_datos0.Close
'    If glusuario = "ADMIN" Or glusuario = "CPLATA" Or glusuario = "DTERCEROS" Or glusuario = "GSOLIZ" Then
'        queryinicial = "Select * from av_calculo_y_cotiza where unidad_codigo = '" & parametro & "' "
'    Else
'        queryinicial = "select * From av_calculo_y_cotiza WHERE unidad_codigo = '" & parametro & "' AND beneficiario_codigo_resp = '" & usuario2 & "'"
'    End If
    rs_datos0.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos0.Recordset = rs_datos0.DataSource
    Set dg_datos0.DataSource = Ado_datos0.Recordset
End Sub

Private Sub Ado_datos0_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos0.Recordset.RecordCount > 0 Then
'    FraDet2.Caption = "COTIZACION PARA EL CLIENTE - " & Ado_datos.Recordset!edif_codigo
'    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
'    '<-- Inicio                Identificación del Cliente                Fin -->    ' PARA el ado.caption
'    'Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'    'Image2 = Img_Foto
''    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
''        'BtnVer.Visible = True
''        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
''        Image2 = Img_Foto
''    Else
''        'BtnVer.Visible = False
''        'chkEstado.Value = vbUnchecked
''    End If

'    If VAR_SW <> "ADD" And VAR_SW <> "" Then
        dtc_desc15.Caption = "0"
        dtc_desc16.Caption = "0"
        dtc_desc17.Caption = "0"
        GlSolicitud = Ado_datos0.Recordset!solicitud_codigo
        GlUnidad = Ado_datos0.Recordset!unidad_codigo
        VAR_PAISC = IIf(IsNull(Ado_datos0.Recordset!pais_continente_cot), "", Ado_datos0.Recordset!pais_continente_cot)
        GlEdificio = Ado_datos0.Recordset!EDIF_CODIGO      ' Codigo Edificio
        GlCotiza = Ado_datos0.Recordset!cotiza_codigo       'txt_codigo1.Caption
        VAR_AME = 0
        VAR_ASI = 0
        VAR_EUR = 0
    Select Case Ado_datos0.Recordset!cotiza_codigo
        Case "1"
            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas), "0", Ado_datos0.Recordset!trafico_num_paradas)
            dtc_desc15.Caption = IIf(IsNull(Ado_datos0.Recordset!pais_continente), "0", Ado_datos0.Recordset!pais_continente)
            dtc_desc16.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_nro_equipos), "0", Ado_datos0.Recordset!trafico_nro_equipos)
        Case "2"
            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas2), "0", Ado_datos0.Recordset!trafico_num_paradas2)
            dtc_desc15.Caption = IIf(IsNull(Ado_datos0.Recordset!pais_continente2), "0", Ado_datos0.Recordset!pais_continente2)
            dtc_desc16.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_nro_equipos2), "0", Ado_datos0.Recordset!trafico_nro_equipos2)
        Case "3"
            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas3), "0", Ado_datos0.Recordset!trafico_num_paradas3)
            dtc_desc15.Caption = IIf(IsNull(Ado_datos0.Recordset!pais_continente3), "0", Ado_datos0.Recordset!pais_continente3)
            dtc_desc16.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_nro_equipos3), "0", Ado_datos0.Recordset!trafico_nro_equipos3)
        Case "4"
            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas4), "0", Ado_datos0.Recordset!trafico_num_paradas4)
            dtc_desc15.Caption = IIf(IsNull(Ado_datos0.Recordset!pais_continente4), "0", Ado_datos0.Recordset!pais_continente4)
            dtc_desc16.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_nro_equipos4), "0", Ado_datos0.Recordset!trafico_nro_equipos4)
    End Select
    
'    Select Case Ado_datos0.Recordset!cotiza_codigo
'        Case "1"
'            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas), "0", Ado_datos0.Recordset!trafico_num_paradas)
'        Case "2"
'            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas2), "0", Ado_datos0.Recordset!trafico_num_paradas2)
'        Case "3"
'            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas2), "0", Ado_datos0.Recordset!trafico_num_paradas3)
'        Case "4"
'            dtc_desc10.Caption = IIf(IsNull(Ado_datos0.Recordset!trafico_num_paradas2), "0", Ado_datos0.Recordset!trafico_num_paradas3)
'    End Select
'
'    Select Case Ado_datos0.Recordset!pais_continente
'        Case "AMERICA"
'            VAR_AME = VAR_AME + Ado_datos0.Recordset!trafico_nro_equipos
'        Case "ASIA"
'            VAR_ASI = VAR_ASI + Ado_datos0.Recordset!trafico_nro_equipos
'        Case "EUROPA"
'            VAR_EUR = VAR_EUR + Ado_datos0.Recordset!trafico_nro_equipos
'    End Select
'    Select Case Ado_datos0.Recordset!pais_continente2
'        Case "AMERICA"
'            VAR_AME = VAR_AME + Ado_datos0.Recordset!trafico_nro_equipos2
'        Case "ASIA"
'            VAR_ASI = VAR_ASI + Ado_datos0.Recordset!trafico_nro_equipos2
'        Case "EUROPA"
'            VAR_EUR = VAR_EUR + Ado_datos0.Recordset!trafico_nro_equipos2
'    End Select
'    Select Case Ado_datos0.Recordset!pais_continente3
'        Case "AMERICA"
'            VAR_AME = VAR_AME + Ado_datos0.Recordset!trafico_nro_equipos3
'        Case "ASIA"
'            VAR_ASI = VAR_ASI + Ado_datos0.Recordset!trafico_nro_equipos3
'        Case "EUROPA"
'            VAR_EUR = VAR_EUR + Ado_datos0.Recordset!trafico_nro_equipos3
'    End Select
'    Select Case Ado_datos0.Recordset!pais_continente4
'        Case "AMERICA"
'            VAR_AME = VAR_AME + Ado_datos0.Recordset!trafico_nro_equipos4
'        Case "ASIA"
'            VAR_ASI = VAR_ASI + Ado_datos0.Recordset!trafico_nro_equipos4
'        Case "EUROPA"
'            VAR_EUR = VAR_EUR + Ado_datos0.Recordset!trafico_nro_equipos4
'    End Select
'        dtc_desc15.Caption = VAR_AME
'        dtc_desc16.Caption = VAR_ASI
'        dtc_desc17.Caption = VAR_EUR
        Call ABRIR_TABLA
        If VAR_PAISC = "AMERICA" Then
'            FraModeloCosto.Visible = False
            dg_datos.Enabled = True
            SSTab1.Tab = 0
            SSTab1.TabEnabled(0) = True
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            If Ado_datos.Recordset.RecordCount > 0 Then
                'dtc_desc15.Caption = Ado_datos0.Recordset!trafico_nro_equipos      'cotiza_cantidad
                Call ABRIR_TABLA_DET
            End If
        Else
            'dtc_desc15.Caption = "0"
            Set dg_det1.DataSource = rsNada
        End If
        If VAR_PAISC = "ASIA" Then
'            FraModeloCostoA.Visible = False
            dg_datosA.Enabled = True
            SSTab1.Tab = 1
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = False
            If Ado_datosA.Recordset.RecordCount > 0 Then
'                dtc_desc16.Caption = Ado_datos0.Recordset!trafico_nro_equipos
                Call ABRIR_TABLA_DET
            End If
        Else
            'dtc_desc16.Caption = "0"
            Set dg_det1.DataSource = rsNada
        End If
        If VAR_PAISC = "EUROPA" Then
    '        FraModeloCostoE.Enabled = False
            dg_datosE.Enabled = True
            SSTab1.Tab = 2
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = True
            If Ado_datosE.Recordset.RecordCount > 0 Then
'                dtc_desc17.Caption = Ado_datos0.Recordset!trafico_nro_equipos
                Call ABRIR_TABLA_DET
            End If
        Else
            'dtc_desc17.Caption = "0"
            Set dg_det1E.DataSource = rsNada
        End If
        
'    If VAR_SW <> "ADD" And VAR_SW <> "" And VAR_PAISC <> "NN" Then
'            Call ABRIR_TABLA_DET
'    Else
'            'Set rs_det1 = New ADODB.Recordset
'            Set dg_det1.DataSource = rsNada
''            Set dg_det2.DataSource = rsNada
'            'Set DtgLaborales.DataSource = rsNada
'    End If
    'txt_aux9.Text = dtc_desc9.Text
    'If Ado_datos0.Recordset!estado_cotiza = "APR" Then
    If Ado_datos0.Recordset!estado_codigo_cot = "REG" Then
            FrmABMDet.Visible = True
            BtnModificar0.Visible = True
'            FrmABMDet2.Visible = False
    Else
            FrmABMDet.Visible = False
            BtnModificar0.Visible = False
'            FrmABMDet2.Visible = True
    End If
    
  End If
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos01 = New ADODB.Recordset
    If rs_datos01.State = 1 Then rs_datos01.Close
    rs_datos01.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    Set Ado_datos01.Recordset = rs_datos01
    txt_campo12.BoundText = txt_campo1.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    'rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos03.Recordset = rs_datos3
    txt_desc3.BoundText = txt_codigo3.BoundText
    txt_aux3.BoundText = txt_codigo3.BoundText
    
    'Cálculo de Tráfico.
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "Select * from ao_solicitud_calculo_trafico where unidad_codigo= '" & parametro & "' and estado_codigo_verif = 'APR' ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
'    dtc_desc11.BoundText = dtc_codigo11.BoundText

    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes ", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
'    dtc_desc21.BoundText = dtc_codigo21.BoundText
    'Modelo 1
    Set rs_datos31 = New ADODB.Recordset
    If rs_datos31.State = 1 Then rs_datos31.Close
    rs_datos31.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo = 'BRA'", db, adOpenStatic
    Set Ado_datos31.Recordset = rs_datos31
'    dtc_desc31.BoundText = dtc_codigo31.BoundText
    'Modelo 2
    Set rs_datos41 = New ADODB.Recordset
    If rs_datos41.State = 1 Then rs_datos41.Close
    rs_datos41.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo =  'ESP'  ", db, adOpenStatic
    Set Ado_datos41.Recordset = rs_datos41
'    dtc_desc41.BoundText = dtc_codigo41.BoundText
    'Modelo 3
    Set rs_datos51 = New ADODB.Recordset
    If rs_datos51.State = 1 Then rs_datos51.Close
    rs_datos51.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo = 'CHN' ", db, adOpenStatic
    Set Ado_datos51.Recordset = rs_datos51
 '   dtc_desc51.BoundText = dtc_codigo51.BoundText
    'Cuadro de Control
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cuadro_ctrl ", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
    'Tipo de Equipo
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from ac_bienes_equipo_tipos ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
'    dtc_desc2.BoundText = dtc_codigo2.BoundText
    'gc_pais
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from gc_pais where pais_continente = 'AMERICA' order by pais_descripcion", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
    'Industria
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from gc_pais where pais_continente = 'ASIA' ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

'Private Sub Maximo_Numerador()
''  TxtCrr.Text = "1"
''  Set RsTmp = New ADODB.Recordset
'''  Set rst_ben = New ADODB.Recordset
'''  rst_ben.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", DB, adOpenStatic
'''  Set AdoTip_ben.Recordset = rst_ben
''  RsTmp.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", db, adOpenStatic
''  'Set RsTmp = DbConex.Execute("Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ;")
''  If Not RsTmp.EOF Then
''     TxtCrr.Text = RsTmp!Codigo
''  End If
'End Sub

'Private Sub Carga_Beneficiario()
''  Set rstbeneficiario = New ADODB.Recordset
''  If rstbeneficiario.State = 1 Then rstbeneficiario.Close
''  sql = "SELECT ges_gestion as gestion,unidad_codigo as Unid_Ejec,solicitud_codigo as Codigo,trafico_codigo,estado_codigo,edif_codigo,trafico_num_paradas,trafico_recorrido," _
''  & " trafico_nro_equipos,vel_equipo_codigo,tipo_puerta,trafico_ancho_puerta,cabina_codigo," _
''  & " tecnologia_codigo , sist_puerta, condicion_ventas " _
''  & " From ao_solicitud_ctrl_trafico WHERE estado_codigo = 'REG'"
'''  SQL = "Select ges_gestion,unidad_codigo,solicitud_codigo,trafico_codigo from ao_solicitud_ctrl_trafico order by unidad_codigo,solicitud_codigo,trafico_codigo"
''  rstbeneficiario.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
''  Set Ado_datos.Recordset = rstbeneficiario
''  'Ado_datos.ConnectionString = sConex
''  'Ado_datos.RecordSource = SQL
''  'Ado_datos.Refresh
''
''  dg_datos.Columns(0).Width = 800 'maxWidth
''  dg_datos.Columns(1).Width = 1556
''  dg_datos.Columns(2).Width = 1556
''  dg_datos.Columns(4).Alignment = dbgRight
'''  dg_datos.Columns(2).Alignment = dbgRight
'''  dg_datos.Columns(3).Alignment = dbgRight
'''  dg_datos.Columns(4).Alignment = dbgCenter
'''  dg_datos.Columns(2).NumberFormat = ("###0.00")
'''  dg_datos.Columns(3).NumberFormat = ("###0.00")
''
''  'LblReg.Caption = "Total Registros --> " & Ado_datos.Recordset.RecordCount
'End Sub

'Function Llena_Combos()
''  CmbReco.Clear
''  sql = " SELECT recorrido_descripcion From ac_bienes_equipo_recorrido; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbReco.AddItem RsTmp!recorrido_descripcion
''         RsTmp.MoveNext
''     Wend
''  End If
'''---
''  CmbNroPasaj.Clear
''  sql = " SELECT pasajeros_descripcion From ac_bienes_equipo_nro_pasajeros; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbNroPasaj.AddItem RsTmp!pasajeros_descripcion
''         RsTmp.MoveNext
''     Wend
''  End If
'''---
'''  CmbVelEq.Clear
'''  SQL = " SELECT vel_equipo_descripcion From ac_bienes_equipo_velocidad WHERE vel_equipo_codigo = " & nCod & "; "
'''  If RsTmp.State = 1 Then RsTmp.Close
'''  RsTmp.Open SQL, DB, adOpenStatic
'''  If Not RsTmp.EOF Then
'''     While Not (RsTmp.EOF)
'''           CmbVelEq.AddItem RsTmp!pasajeros_descripcion
'''         RsTmp.MoveNext
'''     Wend
'''  End If
'End Function
'
'Function Llena_Clientes1()
''  CmbCodCli1.Clear
''  CmbCliente.Clear
''  Call ABRE_CONECCION
''  Set RsTmp = DbConex.Execute("select * from CLIENTES order by nomBRECLI ;")
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbCodCli1.AddItem RsTmp!CodCli
''           CmbCliente.AddItem RsTmp!nombrecli
''         RsTmp.MoveNext
''     Wend
''  End If
''  Call CERRAR_CONECCION
'End Function
'
'Private Sub CmbCliente_Click()
'' If CmbCliente.ListIndex = -1 Then Exit Sub
'' CmbCodCli1.ListIndex = CmbCliente.ListIndex
'End Sub
'
'Private Sub dg_datos_Click()
''  MsgBox "sss"
''   Call Llena_Varios
''  txtDescrip = dg_datos.Columns(1).Text
'End Sub

'Private Sub dg_datos_KeyDown(KeyCode As Integer, Shift As Integer)
'  Call Llena_Varios
''  txtDescrip = dg_datos.Columns(1).Text
'End Sub
'Function Llena_Varios()
''  If RsTmp.State = 1 Then RsTmp.Close
''  'If DB.State = adStateOpen Then DB.Close
''  sql = " SELECT unidad_descripcion FROM gc_unidad_ejecutora " & _
''        "  WHERE unidad_codigo = '" & TxtUEjec & "';"
''  RsTmp.Open sql, db, adCmdText 'adOpenStatic
''  If Not RsTmp.EOF Then
''     txtDescrip.Text = RsTmp!unidad_descripcion
''  End If
'''--
''  sql = " SELECT edif_denominacion FROM gc_edificaciones " & _
''              "  WHERE edif_codigo = '" & Txtedif & "';"
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     txtDesEdif.Text = RsTmp!edif_denominacion
''  End If
'''-------
''  CmbVelEq.Clear
''  sql = " SELECT vel_equipo_descripcion From ac_bienes_equipo_velocidad WHERE vel_equipo_codigo = " & TxtCodVel & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbVelEq.AddItem RsTmp!vel_equipo_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbVelEq.ListIndex = 0
''  End If
'''-------
''  CmbTipoPuerta.Clear
''  sql = " SELECT tipo_puerta_descripcion From ac_bienes_equipo_tipo_puerta_piso WHERE tipo_puerta = " & Txttip & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbTipoPuerta.AddItem RsTmp!tipo_puerta_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbTipoPuerta.ListIndex = 0
''  End If
'''-------cabina_codigo
''  CmbEstat.Clear
''  sql = " SELECT cabina_descripcion From ac_bienes_equipo_cabina_estetica WHERE cabina_codigo = '" & TxtEst & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbEstat.AddItem RsTmp!cabina_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbEstat.ListIndex = 0
''  End If
'''-------
''  CmbTecno.Clear
''  sql = " SELECT tecnologia_descripcion From ac_bienes_equipo_tecnologia WHERE tecnologia_codigo = '" & TxtTecno & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbTecno.AddItem RsTmp!tecnologia_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbTecno.ListIndex = 0
''  End If
''        'FALTA sist_puerta
'''-------
''  CmbCondVenta.Clear
''  sql = " SELECT condicion_ventas_descripcion From ac_bienes_equipo_condicion_ventas WHERE condicion_ventas = '" & TxtCondVenta & "'; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open sql, db, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbCondVenta.AddItem RsTmp!condicion_ventas_descripcion
''         RsTmp.MoveNext
''     Wend
''     CmbCondVenta.ListIndex = 0
''  End If
''
'End Function

'Private Sub dtc_aux1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_aux1.BoundText
'    dtc_codigo1.BoundText = dtc_aux1.BoundText
'End Sub
'
'Private Sub dtc_aux3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_aux3.BoundText
'    dtc_desc3.BoundText = dtc_aux3.BoundText
'End Sub

'Private Sub dtc_codigo1_Click(Area As Integer)
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
''    dtc_aux1.BoundText = dtc_codigo1.BoundText
'End Sub

'Private Sub dtc_codigo21_Click(Area As Integer)
'    dtc_desc21.BoundText = dtc_codigo21.BoundText
''    dtc_desc22.BoundText = dtc_codigo21.BoundText
''    dtc_desc23.BoundText = dtc_codigo21.BoundText
'    dtc_desc24.BoundText = dtc_codigo21.BoundText
'End Sub

Private Sub dtc_codigo22_Click(Area As Integer)
'    dtc_desc22.BoundText = dtc_codigo22.BoundText
End Sub

Private Sub dtc_codigo23_Click(Area As Integer)
    'dtc_desc23.BoundText = dtc_codigo23.BoundText
End Sub

'Private Sub dtc_codigo3_Click(Area As Integer)
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
'    dtc_aux3.BoundText = dtc_codigo3.BoundText
'End Sub

'Private Sub dtc_codigo61_Click(Area As Integer)
'    dtc_desc61.BoundText = dtc_codigo61.BoundText
'End Sub

'Private Sub dtc_desc1_Click(Area As Integer)
'    dtc_codigo1.BoundText = dtc_desc1.BoundText
''    dtc_aux1.BoundText = dtc_desc1.BoundText
''    Call pnivel1(dtc_codigo1.BoundText)
''    dtc_desc10.Enabled = True
''    Call pnivel11(dtc_codigo1.BoundText)
''    dtc_desc11.Enabled = True
'End Sub

'Private Sub pnivel1(codigo1 As String)
''   Dim strConsultaF As String
''   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'
'   Set dtc_codigo10.RowSource = Nothing
''   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo10.ReFill
'   dtc_codigo10.BoundText = Empty
'
'   Set dtc_desc10.RowSource = Nothing
'   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc10.ReFill
'   dtc_desc10.BoundText = Empty
'End Sub

'Private Sub pnivel11(codigo1 As String)
'   Dim strConsultaF As String
'   'strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
'   strConsultaF = "Select * from gv_personal_contratado where unidad_codigo = '" & codigo1 & "' order by beneficiario_denominacion"
'
'   Set dtc_codigo11.RowSource = Nothing
'   Set dtc_codigo11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_codigo11.ReFill
'   dtc_codigo11.BoundText = Empty
'
'   Set dtc_desc11.RowSource = Nothing
'   Set dtc_desc11.RowSource = db.Execute(strConsultaF, , adCmdText)
'   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
'   dtc_desc11.ReFill
'   dtc_desc11.BoundText = Empty
'End Sub

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

'Private Sub dtc_desc21_Click(Area As Integer)
'    dtc_codigo21.BoundText = dtc_desc21.BoundText
''    dtc_desc22.BoundText = dtc_desc21.BoundText
''    dtc_desc23.BoundText = dtc_desc21.BoundText
'    dtc_desc24.BoundText = dtc_desc21.BoundText
'End Sub

'Private Sub dtc_desc3_Click(Area As Integer)
'    dtc_codigo3.BoundText = dtc_desc3.BoundText
'    dtc_aux3.BoundText = dtc_desc3.BoundText
'End Sub

'Private Sub dtc_desc61_Click(Area As Integer)
'    dtc_codigo61.BoundText = dtc_desc61.BoundText
'End Sub

'Private Sub FraModelo_Click()
''    FraModelo.Visible = False
'    FraModeloCosto.Visible = True
'End Sub

Private Sub ABRIR_TABLA()
    
    Select Case VAR_PAISC
        Case "AMERICA"
            'Cotiza AMERICA (Brasil, EE.UU, y otros)
            Set rs_datos = New Recordset
            If rs_datos.State = 1 Then rs_datos.Close
            rs_datos.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & GlUnidad & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'AMERICA'  and cotiza_codigo = " & GlCotiza & " ", db, adOpenKeyset, adLockOptimistic
            'queryinicial = "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Text & "' and solicitud_codigo = " & Ado_datos0.Recordset!solicitud_codigo & " and pais_continente = 'AMERICA'  "            'txt_codigo1.Caption
            'queryinicial = "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & GlCotiza & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'AMERICA'  "
            'rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
            Set Ado_datos.Recordset = rs_datos.DataSource
            Set dg_datos.DataSource = Ado_datos.Recordset
            If Ado_datos.Recordset.RecordCount > 0 Then
                'VAR_PAISC = "AMERICA"
                'txt_codigo1.Caption = Ado_datos.Recordset!cotiza_codigo
            End If
        Case "ASIA"
            'Cotiza ASIA (China, Japon)
            Set rs_datosA = New ADODB.Recordset
            If rs_datosA.State = 1 Then rs_datosA.Close
            rs_datosA.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & GlUnidad & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'ASIA' and cotiza_codigo = " & GlCotiza & "  ", db, adOpenKeyset, adLockOptimistic
            'rs_datosA.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Txt_campo1.Text & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'ASIA'  ", db, adOpenKeyset, adLockOptimistic         'txt_codigo1.Caption
            Set Ado_datosA.Recordset = rs_datosA
            Set dg_datosA.DataSource = Ado_datosA.Recordset
            If Ado_datosA.Recordset.RecordCount > 0 Then
                'VAR_PAISC = "ASIA"
                'txt_codigo1.Caption = Ado_datosA.Recordset!cotiza_codigo
            End If
        Case "EUROPA"
            'Cotiza EUROPA (España, Francia, etc.)
            Set rs_datosE = New ADODB.Recordset
            If rs_datosE.State = 1 Then rs_datosE.Close
            'rs_datosE.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'EUROPA' and cotiza_codigo = " & txt_codigo1.Caption & " ", db, adOpenKeyset, adLockOptimistic
            rs_datosE.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & GlUnidad & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'EUROPA'  and cotiza_codigo = " & GlCotiza & " ", db, adOpenKeyset, adLockOptimistic
            Set Ado_datosE.Recordset = rs_datosE
            Set dg_datosE.DataSource = Ado_datosE.Recordset
            If Ado_datosE.Recordset.RecordCount > 0 Then
                'VAR_PAISC = "EUROPA"
                'txt_codigo1.Caption = Ado_datosE.Recordset!cotiza_codigo
            End If
    End Select
        
'    dtc_desc31.BoundText = dtc_codigo31.BoundText
'    dtc_desc32.BoundText = dtc_codigo31.BoundText
'    dtc_desc33.BoundText = dtc_codigo31.BoundText
'    dtc_desc34.BoundText = dtc_codigo31.BoundText
    
'    dtc_desc41.BoundText = dtc_codigo41.BoundText
'    dtc_desc42.BoundText = dtc_codigo41.BoundText
'    dtc_desc43.BoundText = dtc_codigo41.BoundText
'    dtc_desc44.BoundText = dtc_codigo41.BoundText
    
'    dtc_desc51.BoundText = dtc_codigo51.BoundText
'    dtc_desc52.BoundText = dtc_codigo51.BoundText
'    dtc_desc53.BoundText = dtc_codigo51.BoundText
'    dtc_desc54.BoundText = dtc_codigo51.BoundText
End Sub

'Private Sub OptFilGral1_Click()
'    'parametro = "estado_codigo" + " = " + "'REG'"
'    Call ABRIR_TABLA
'End Sub
'
'Private Sub OptFilGral2_Click()
'    'parametro = "estado_codigo" + " <> " + "'0'"
'    Call ABRIR_TABLA
'End Sub

'Private Sub Img_03_Click()
' If AdoPermiso.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo asociado al Registro, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'   If GlServidor = "SRVPRO" Then
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   Else
'      If AdoPermiso.Recordset!TipoPermiso = "VC" Then
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(AdoPermiso.Recordset!solicitud_codigo) & "\LICENCIAS\" & Trim(AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      End If
'   End If
' End If
'
'End Sub

'Private Sub Img_CTO_Click()
' If Ado_Memo.Recordset!ARCHIVO = "Cargar_Archivo" Then
'    MsgBox "No Existe el Archivo Asociado al Contrato, debe Cargarlo ...", vbExclamation, "Advertencia"
' Else
'    'If GlServidor <> GlMaquina Then      ' "-" Then
'    If GlServidor = "SRVPRO" Then
'        'e = ShellExecute(Img_CTO, "open", "\\" & Trim(GlServidor) & "\SIS_PROAGRO\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    Else
'        'e = ShellExecute(Img_CTO, "open", App.Path & "\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, SW_SHOWNORMAL)
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_Memo.Recordset!solicitud_codigo) & "\CONTRATOS\" & Trim(Ado_Memo.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'    End If
' End If
'End Sub

'Private Sub Img_CV_Click()
''    Dim e As Long
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_HOJAVIDA = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "C_V"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'         ' e = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario.AdoMovilidad.Recordset!solicitud_codigo) & "\FINIQUITO\" & Trim(Ado_Auxiliar.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      'MsgBox ""
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "C_V"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'  End If
'  If GlServidor = "SRVPRO" Then
'        imag2 = ShellExecute(0, vbNullString, "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  Else
'        imag2 = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\VACACIONES\" & Trim(Ado_datos.Recordset!ARCHIVO_VAC), vbNullString, vbNullString, vbNormalFocus)
'  End If
'End Sub
'
'Private Sub Img_Foto_Click()
'  If swnuevo <> "X" Then
'    If Ado_datos.Recordset!ARCHIVO_FOTO = "Cargar_Archivo" Then
'      NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Frmexporta.DirDestino.Path = NombreCarpeta
'      GlArch = "FOT"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'      Else
'         e = NombreCarpeta
'      End If
'      Frmexporta.DirDestino2.Path = e
'      Frmexporta.Show vbModal
'    Else
'      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'          NombreCarpeta = App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Frmexporta.DirDestino.Path = NombreCarpeta
'          GlArch = "FOT"
'          'If GlServidor <> GlMaquina Then      ' "-" Then
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!solicitud_codigo) & "\"
'          Else
'            e = NombreCarpeta
'          End If
'          Frmexporta.DirDestino2.Path = e
'          Frmexporta.Show vbModal
'      End If
'    End If
'
'    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\" & Trim(GLCarpeta) & "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
'        ARCH_FOTO = App.Path + "\" & Trim(GLCarpeta2) & "\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("solicitud_codigo")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    End If
'    If Guardar_Imagen(db, "Select Foto From Gc_beneficiario Where solicitud_codigo= '" & Ado_datos.Recordset("solicitud_codigo") & "' ", "Foto", ARCH_FOTO) Then
'        MsgBox "Se cargo la Imagen Correctamente !!"
'    Else
'        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
'    End If
'  End If
'End Sub

'Private Sub SSTab1_DblClick()
'    If SSTab1.Tab = 0 Then
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'  If glPersNew = "P" Then
'  End If
'  glPersNew = "N"
'
''   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    If VAR_PAISC = "AMERICA" Then
        'rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Val(GlCotiza) & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle1.Recordset = rs_det1
        If Ado_detalle1.Recordset.RecordCount > 0 Then
            Set dg_det1.DataSource = Ado_detalle1.Recordset
        Else
            Set dg_det1.DataSource = rsNada
        End If
    End If
    If VAR_PAISC = "ASIA" Then
        'rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Val(GlCotiza) & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle1.Recordset = rs_det1
        If Ado_detalle1.Recordset.RecordCount > 0 Then
            Set dg_det1.DataSource = Ado_detalle1.Recordset
        Else
            Set dg_det1.DataSource = rsNada
        End If
    End If
    If VAR_PAISC = "EUROPA" Then
        'rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
        rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & txt_campo1.Text & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = '" & VAR_PAISC & "' and cotiza_codigo = " & Val(GlCotiza) & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText           'txt_codigo1.Caption
        Set Ado_detalle1.Recordset = rs_det1
        If Ado_detalle1.Recordset.RecordCount > 0 Then
            Set dg_det1E.DataSource = Ado_detalle1.Recordset
        Else
            Set dg_det1E.DataSource = rsNada
        End If
        
    End If
    
    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    'rs_datos21.Open "Select * from ac_bienes where edif_codigo = '" & dtc_codigo3.Text & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    rs_datos21.Open "Select * from ac_bienes where edif_codigo = '" & GlEdificio & "' OR modelo_codigo= 'NA' ", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
'    dtc_desc21.BoundText = dtc_codigo21.BoundText

'    If VAR_PAISC = "AMERICA" Then
'        'Cotiza AMERICA (Brasil, EE.UU, y otros)
'        Set rs_datos1 = New ADODB.Recordset
'        If rs_datos1.State = 1 Then rs_datos1.Close
'        'rs_datos1.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'AMERICA' and cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & " ", db, adOpenKeyset, adLockOptimistic
'        rs_datos1.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'AMERICA' and cotiza_codigo = " & txt_codigo1.Caption & " ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos1.Recordset = rs_datos1
'        'Set dg_datosA.DataSource = Ado_datosA.Recordset
'    End If
    
'    If VAR_PAISC = "ASIA" Then
'        'Cotiza ASIA (China, Japon)
'        Set rs_datos1A = New ADODB.Recordset
'        If rs_datos1A.State = 1 Then rs_datos1A.Close
'        'rs_datos1A.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'ASIA' and cotiza_codigo = " & Ado_datosA.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'        rs_datos1A.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'ASIA' and cotiza_codigo = " & txt_codigo1.Caption & "  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datos1A.Recordset = rs_datos1A
'        Set dg_datos1A.DataSource = Ado_datos1A.Recordset
'    End If
    
'    If VAR_PAISC = "EUROPA" Then
'        'Cotiza Europa (España...)
'        Set rs_datos1E = New ADODB.Recordset
'        If rs_datos1E.State = 1 Then rs_datos1A.Close
'        'rs_datos1E.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'EUROPA' and cotiza_codigo = " & Ado_datosE.Recordset!cotiza_codigo & "  ", db, adOpenKeyset, adLockOptimistic
'        rs_datos1E.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " and pais_continente = 'EUROPA' and cotiza_codigo = " & txt_codigo1.Caption & "  ", db, adOpenKeyset, adLockOptimistic
'        Set Ado_datosE.Recordset = rs_datos1E
'        Set dg_datos1E.DataSource = Ado_datosE.Recordset
'    End If

End Sub

Private Sub sstab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0
            'lbl_titulo1.Caption = sstab1.Caption
            FraNavega.Caption = VAR_FRA + SSTab1.Caption
'            FraGrabarCancelar.Visible = False
            Call ABRIR_TABLA
        Case 1
            FraNavegaA.Caption = VAR_FRA + SSTab1.Caption
'            FraGrabarCancelarA.Visible = False
            Call ABRIR_TABLA
        Case 2
            FraNavegaE.Caption = VAR_FRA + SSTab1.Caption
'            FraGrabarCancelar5.Visible = False
            Call ABRIR_TABLA
    End Select
End Sub

Private Sub Txt_campo1_Click(Area As Integer)
    txt_campo12.BoundText = txt_campo1.BoundText
End Sub

Private Sub Txt_campo12_Click(Area As Integer)
    txt_campo1.BoundText = txt_campo12.BoundText
End Sub

Private Sub txt_codigo3_Click(Area As Integer)
    txt_desc3.BoundText = txt_codigo3.BoundText
    txt_aux3.BoundText = txt_codigo3.BoundText
End Sub


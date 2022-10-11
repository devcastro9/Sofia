VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_to_solicitud_cotiza_venta 
   BackColor       =   &H00000000&
   Caption         =   "Cotización de Servicios"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "frm_to_solicitud_cotiza_venta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1640
      Left            =   120
      Picture         =   "frm_to_solicitud_cotiza_venta.frx":0A02
      ScaleHeight     =   1575
      ScaleWidth      =   1875
      TabIndex        =   126
      Top             =   8120
      Width           =   1935
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Adiciona Detalle"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   945
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":6CE76
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   600
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":6D2B8
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_to_solicitud_cotiza_venta.frx":6D6FA
      ScaleHeight     =   960
      ScaleWidth      =   15240
      TabIndex        =   114
      Top             =   60
      Width           =   15300
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":D972C
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":D9D50
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DA330
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6840
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DAFFA
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "R-222"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DB204
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DB7C1
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DBD79
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DBF83
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DC3C5
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir2 
         BackColor       =   &H00808000&
         Caption         =   "R-224"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":DC5CF
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
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
         Left            =   10635
         TabIndex        =   125
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_to_solicitud_cotiza_venta.frx":DCB8C
      ScaleHeight     =   915
      ScaleWidth      =   15240
      TabIndex        =   110
      Top             =   60
      Width           =   15300
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":148BBE
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_to_solicitud_cotiza_venta.frx":148DC8
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITUD DE COTIZACIÓN"
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
         Left            =   8460
         TabIndex        =   113
         Top             =   300
         Width           =   4185
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE COSTOS"
      ForeColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   2160
      TabIndex        =   55
      Top             =   8040
      Width           =   12975
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   1335
         Left            =   195
         TabIndex        =   56
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   2355
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
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   5
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "costo_monto"
            Caption         =   "Modelo1 Bs."
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
            Caption         =   "Modelo1 ME"
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
            Caption         =   "Modelo2 Bs"
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
            Caption         =   "Modelo2 ME"
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
            Caption         =   "Modelo3 Bs"
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
            Caption         =   "Modelo3 ME"
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
            Caption         =   "Observaciones"
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
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   5580.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_datos2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000040&
      Height          =   1680
      Left            =   5880
      TabIndex        =   39
      Top             =   1080
      Width           =   9465
      Begin VB.TextBox txt_codigo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         DataField       =   "cotiza_codigo"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   6195
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   49
         Top             =   490
         Width           =   315
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         DataField       =   "trafico_codigo"
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
         ForeColor       =   &H80000015&
         Height          =   290
         Left            =   8920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   42
         Top             =   1155
         Width           =   330
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         DataField       =   "estado_codigo"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         TabIndex        =   40
         Top             =   1150
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   44
         Top             =   240
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
      Begin MSDataListLib.DataCombo dtc_desc1 
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1680
         TabIndex        =   45
         Top             =   480
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "unidad_descripcion"
         BoundColumn     =   "unidad_codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   1140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   58523651
         CurrentDate     =   41678
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   46
         Top             =   1140
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   47
         Top             =   1140
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "codigo1"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   $"frm_to_solicitud_cotiza_venta.frx":148FD2
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
         Left            =   240
         TabIndex        =   130
         Top             =   880
         Width           =   9000
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "unidad_codigo_ant"
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
         Left            =   7560
         TabIndex        =   108
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "solicitud_codigo"
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
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"frm_to_solicitud_cotiza_venta.frx":149063
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
         Index           =   12
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   8790
      End
   End
   Begin VB.PictureBox Fra_datos 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5235
      Left            =   5880
      ScaleHeight     =   5175
      ScaleWidth      =   9405
      TabIndex        =   22
      Top             =   2760
      Width           =   9465
      Begin VB.TextBox Text8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   9120
         TabIndex        =   105
         Top             =   4680
         Width           =   375
      End
      Begin VB.Frame FraModeloCosto 
         BackColor       =   &H80000017&
         Caption         =   "----- Modelo del Equipo --------------- Precio FOB Bs -- Precio FOB (ME) ----  Total Bs ---------- Total (ME)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   1695
         Left            =   0
         TabIndex        =   66
         Top             =   1605
         Width           =   9400
         Begin VB.CommandButton CmdMod3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Modelo3"
            Height          =   315
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Ver Características del Modelo3"
            Top             =   1240
            Width           =   765
         End
         Begin VB.CommandButton CmdMod2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Modelo2"
            Height          =   315
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Ver Características del Modelo2"
            Top             =   750
            Width           =   765
         End
         Begin VB.CommandButton CmdMod1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Modelo1"
            Height          =   315
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Ver Características del Modelo1"
            Top             =   280
            Width           =   765
         End
         Begin VB.TextBox txt_monto10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_dol_x"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4680
            TabIndex        =   10
            Text            =   "0"
            Top             =   1240
            Width           =   1365
         End
         Begin VB.TextBox txt_monto6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_dol_h"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4680
            TabIndex        =   7
            Text            =   "0"
            Top             =   750
            Width           =   1365
         End
         Begin VB.TextBox txt_monto2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_dol"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4680
            TabIndex        =   4
            Text            =   "0"
            Top             =   280
            Width           =   1365
         End
         Begin VB.TextBox txt_monto11 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_bs_x"
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
            Height          =   285
            Left            =   6240
            TabIndex        =   82
            Text            =   "0"
            Top             =   1240
            Width           =   1485
         End
         Begin VB.TextBox txt_monto7 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_bs_h"
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
            Height          =   285
            Left            =   6240
            TabIndex        =   81
            Text            =   "0"
            Top             =   750
            Width           =   1485
         End
         Begin VB.TextBox txt_monto3 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_bs"
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
            Height          =   285
            Left            =   6240
            TabIndex        =   80
            Text            =   "0"
            Top             =   280
            Width           =   1485
         End
         Begin VB.TextBox txt_monto12 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_dol_x"
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
            Height          =   285
            Left            =   7800
            TabIndex        =   79
            Text            =   "0"
            Top             =   1240
            Width           =   1485
         End
         Begin VB.TextBox txt_monto8 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_dol_h"
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
            Height          =   285
            Left            =   7800
            TabIndex        =   78
            Text            =   "0"
            Top             =   750
            Width           =   1485
         End
         Begin VB.TextBox txt_monto4 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            BorderStyle     =   0  'None
            DataField       =   "cotiza_precio_total_dol"
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
            Height          =   285
            Left            =   7800
            TabIndex        =   77
            Text            =   "0"
            Top             =   280
            Width           =   1485
         End
         Begin VB.TextBox txt_monto9 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_bs_x"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   1240
            Width           =   1365
         End
         Begin VB.TextBox txt_monto5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_bs_h"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0"
            Top             =   750
            Width           =   1365
         End
         Begin VB.TextBox txt_monto1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            DataField       =   "cotiza_precio_fob_bs"
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "0"
            Top             =   280
            Width           =   1365
         End
         Begin VB.Label Txt_campo6 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "modelo_codigo_x"
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
            Left            =   360
            TabIndex        =   85
            Top             =   1240
            Width           =   1935
         End
         Begin VB.Label Txt_campo5 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "modelo_codigo_h"
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
            Left            =   360
            TabIndex        =   84
            Top             =   750
            Width           =   1935
         End
         Begin VB.Label Txt_campo4 
            Alignment       =   2  'Center
            BackColor       =   &H80000013&
            Caption         =   "0"
            DataField       =   "modelo_codigo"
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
            Left            =   360
            TabIndex        =   83
            Top             =   280
            Width           =   1935
         End
         Begin VB.Label lbl_campo3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "1."
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
            Left            =   120
            TabIndex        =   69
            Top             =   285
            Width           =   150
         End
         Begin VB.Label lbl_campo4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "2."
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
            Left            =   120
            TabIndex        =   68
            Top             =   750
            Width           =   150
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "3."
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
            Left            =   120
            TabIndex        =   67
            Top             =   1245
            Width           =   150
         End
      End
      Begin VB.Frame FraModelo 
         BackColor       =   &H80000017&
         ForeColor       =   &H0080C0FF&
         Height          =   1695
         Left            =   0
         TabIndex        =   86
         Top             =   1605
         Width           =   9390
         Begin VB.CommandButton CmdVolver 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Volver"
            Height          =   720
            Left            =   8680
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Cerrar Ventana"
            Top             =   600
            Width           =   645
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   8640
            TabIndex        =   106
            Top             =   240
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dtc_desc54 
            DataField       =   "modelo_codigo_x"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   102
            Top             =   1245
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "vel_equipo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_desc44 
            DataField       =   "modelo_codigo_h"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   103
            Top             =   750
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "vel_equipo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_desc34 
            DataField       =   "modelo_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   7560
            TabIndex        =   104
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "vel_equipo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "0"
         End
         Begin MSDataListLib.DataCombo dtc_desc53 
            DataField       =   "modelo_codigo_x"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   99
            Top             =   1245
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pasajeros_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc43 
            DataField       =   "modelo_codigo_h"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   100
            Top             =   750
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pasajeros_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_desc33 
            DataField       =   "modelo_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   5880
            TabIndex        =   101
            Top             =   285
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pasajeros_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc51 
            DataField       =   "modelo_codigo_x"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   96
            Top             =   1245
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc41 
            DataField       =   "modelo_codigo_h"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   97
            Top             =   750
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtc_desc31 
            DataField       =   "modelo_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   4560
            TabIndex        =   98
            Top             =   285
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "pais_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc52 
            DataField       =   "modelo_codigo_x"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   93
            Top             =   1245
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "marca_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_desc42 
            DataField       =   "modelo_codigo_h"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   94
            Top             =   750
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "marca_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_desc32 
            DataField       =   "modelo_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   2880
            TabIndex        =   95
            Top             =   285
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "marca_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo51 
            DataField       =   "modelo_codigo_x"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   480
            TabIndex        =   90
            Top             =   1245
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "modelo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo41 
            DataField       =   "modelo_codigo_h"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   480
            TabIndex        =   91
            Top             =   750
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "modelo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_codigo31 
            DataField       =   "modelo_codigo"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   480
            TabIndex        =   92
            Top             =   285
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            ListField       =   "modelo_codigo"
            BoundColumn     =   "modelo_codigo"
            Text            =   ""
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "3."
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
            Left            =   120
            TabIndex        =   89
            Top             =   1245
            Width           =   150
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "2."
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
            Left            =   120
            TabIndex        =   88
            Top             =   750
            Width           =   150
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "1."
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
            Left            =   120
            TabIndex        =   87
            Top             =   285
            Width           =   150
         End
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   7200
         TabIndex        =   75
         Top             =   1060
         Width           =   375
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   1500
         TabIndex        =   74
         Top             =   1060
         Width           =   375
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   9030
         TabIndex        =   73
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   7110
         TabIndex        =   72
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   5310
         TabIndex        =   71
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   310
         Left            =   3390
         TabIndex        =   70
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Txt_campo7 
         DataField       =   "bien_cotiza_num_accesos"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   13
         Text            =   "0"
         Top             =   3480
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "cotiza_energia"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "0"
         Top             =   3525
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo3 
         DataField       =   "cotiza_luz"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4560
         TabIndex        =   12
         Text            =   "0"
         Top             =   3525
         Width           =   1365
      End
      Begin VB.TextBox txt_monto0 
         Alignment       =   2  'Center
         DataField       =   "cotiza_tdc_bol"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   7560
         TabIndex        =   1
         Text            =   "0"
         Top             =   1060
         Width           =   1485
      End
      Begin VB.TextBox Txt_campo10 
         DataField       =   "bien_cotiza_dimension_fosa_frente"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   4560
         TabIndex        =   15
         Text            =   "0"
         Top             =   4120
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo9 
         DataField       =   "bien_cotiza_dimension_fosa_fondo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Text            =   "0"
         Top             =   4120
         Width           =   1365
      End
      Begin VB.TextBox Txt_campo8 
         DataField       =   "bien_cotiza_dimension_fosa_m"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7800
         TabIndex        =   16
         Text            =   "0"
         Top             =   4120
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         DataField       =   "trafico_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "trafico_h_nro_total_equipos"
         BoundColumn     =   "trafico_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc12 
         DataField       =   "trafico_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3840
         TabIndex        =   24
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "trafico_h_partidas_por_hora"
         BoundColumn     =   "trafico_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc13 
         DataField       =   "trafico_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5760
         TabIndex        =   25
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "trafico_h_intervalo_trafico"
         BoundColumn     =   "trafico_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc14 
         DataField       =   "trafico_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7570
         TabIndex        =   26
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "trafico_h_capacidad_trafico"
         BoundColumn     =   "trafico_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc21 
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   1060
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc22 
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "par_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc23 
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6720
         TabIndex        =   29
         Top             =   795
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "subgrupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc24 
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   30
         Top             =   795
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "grupo_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "0"
      End
      Begin MSDataListLib.DataCombo dtc_desc61 
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   4680
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_descripcion"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc62 
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8160
         TabIndex        =   31
         Top             =   4680
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "tiene_sala_maq"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         DataField       =   "trafico_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1200
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "trafico_codigo"
         BoundColumn     =   "trafico_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo21 
         DataField       =   "bien_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   1060
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   -2147483629
         ForeColor       =   16777215
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "36NO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_codigo61 
         DataField       =   "cuadro_ctrl_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   4440
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "cuadro_ctrl_codigo"
         BoundColumn     =   "cuadro_ctrl_codigo"
         Text            =   ""
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensión Fosa Alto (mm)"
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
         Height          =   600
         Left            =   6480
         TabIndex        =   109
         Top             =   4005
         Width           =   1410
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tiene Sala Máquinas?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   6120
         TabIndex        =   65
         Top             =   4695
         Width           =   2025
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Grupo"
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Left            =   5040
         TabIndex        =   64
         Top             =   795
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Sub Grupo"
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Left            =   6360
         TabIndex        =   63
         Top             =   795
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cód. Equipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   120
         TabIndex        =   62
         Top             =   795
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Descripción del Equipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   61
         Top             =   795
         Width           =   2100
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Capacidad Tráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   7440
         TabIndex        =   60
         Top             =   105
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Intérvalo Tráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   5760
         TabIndex        =   59
         Top             =   100
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Partidas por Hora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   3840
         TabIndex        =   58
         Top             =   100
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro.Total Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   240
         Left            =   1920
         TabIndex        =   57
         Top             =   100
         Width           =   1635
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Accesos"
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
         Height          =   480
         Left            =   6480
         TabIndex        =   54
         Top             =   3410
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7560
         TabIndex        =   53
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cuarto de Control"
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
         Left            =   120
         TabIndex        =   52
         Top             =   4695
         Width           =   1545
      End
      Begin VB.Label lbl_campo13 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dim. Fosa Frente (mm)"
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
         Height          =   600
         Left            =   3240
         TabIndex        =   50
         Top             =   4005
         Width           =   1290
      End
      Begin VB.Label lbl_campo12 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Dim. Fosa Fondo (mm)"
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
         Height          =   480
         Left            =   120
         TabIndex        =   38
         Top             =   4000
         Width           =   1305
      End
      Begin VB.Label lbl_campo7 
         BackColor       =   &H00000000&
         Caption         =   "Iluminación / Luz (V)"
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
         Height          =   480
         Left            =   3240
         TabIndex        =   37
         Top             =   3400
         Width           =   1035
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo de Tráfico"
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
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lbl_campo6 
         BackColor       =   &H00000000&
         Caption         =   "Fuerza Motriz / Energía (V)"
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
         Height          =   480
         Left            =   120
         TabIndex        =   35
         Top             =   3405
         Width           =   1245
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFC0&
      Height          =   6915
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   5655
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3120
         TabIndex        =   21
         Top             =   6435
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1320
         TabIndex        =   20
         Top             =   6435
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   6360
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   6090
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   10742
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "solicitud_codigo"
            Caption         =   "Nro.Tramite"
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
            Caption         =   "U. Ejecutora"
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
            Caption         =   "Cod.Proyecto"
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
            Caption         =   "No.Cotiza"
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
            DataField       =   "bien_codigo"
            Caption         =   "Cod.Equipo"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   794.835
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Ado_datos21 
      Height          =   330
      Left            =   6480
      Top             =   9840
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
      Top             =   9840
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
      Top             =   9840
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
      Top             =   9840
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
      Top             =   10200
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
      Top             =   9840
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
      Left            =   4320
      Top             =   10200
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
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -1560
      Top             =   23640
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
      Caption         =   "Ado_datos23"
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
      Top             =   10200
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
End
Attribute VB_Name = "frm_to_solicitud_cotiza_venta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_datos21 As New ADODB.Recordset
Dim rs_datos31 As New ADODB.Recordset
Dim rs_datos41 As New ADODB.Recordset
Dim rs_datos51 As New ADODB.Recordset
Dim rs_datos61 As New ADODB.Recordset

Dim rstbeneficiario As New ADODB.Recordset
Dim rst_ben, rsNada As New ADODB.Recordset
Dim RsTmp As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_det1 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

'OTROS
Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
'Dim swnuevo As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim imag2 As Long
Dim parametro As String
Dim var_cod As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW As String
Dim VAR_AUX, VAR_CONT2 As Double

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
        If Ado_datos.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLA_DET
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    Fra_datos.Enabled = False
    Fra_datos2.Enabled = False
        
    aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.txt_codigo.Caption     ' Nro. Negociacion (Cod.solicitud)
    aw_p_ao_solicitud_cotiza_detalle.txt_campo1.Caption = Me.dtc_codigo1.Text       ' Codigo Unidad
    aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.dtc_desc1.Text    ' Descripcion Unidad
    aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.txt_codigo1.Text       ' Nro. Cotización
    aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = Me.dtc_codigo3.Text       ' Codigo Edificio
    If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
    Else
        aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Me.txt_monto1.Text   ' Monto Modelo1(ME)
    End If
    If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
    Else
        aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = Me.txt_monto5.Text   ' Monto Modelo2(ME)
    End If
    If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
    Else
        aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = Me.txt_monto9.Text   ' Monto Modelo3(ME)
    End If
    Ado_detalle1.Recordset.AddNew
    aw_p_ao_solicitud_cotiza_detalle.Show vbModal
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
  End If
  'Call ABRIR_TABLA_DET
  Ado_datos.Recordset.Move marca1 - 1
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
  On Error GoTo UpdateErr
   Set rs_aux2 = New ADODB.Recordset
   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux2.RecordCount > 0 Then
        VAR_CONT2 = rs_aux2.RecordCount
   End If
   'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
   If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
                SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datos.Recordset!cotiza_precio_total_dol
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
                    rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = VAR_AUX
                    rs_aux1!venta_monto_total_bs = Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!fecha_registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
'            Case "4"
'        End Select
        'GRABA VENTA DETALLE
        If var_cod = "" Then
            var_cod = rs_aux1!venta_codigo
        End If
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
        rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & Year(Date) & "'   ", db, adOpenKeyset, adLockOptimistic
        'If rs_aux3.RecordCount > 0 Then
            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
        'Else
            VAR_AUX = rs_aux3.RecordCount + 1
            rs_aux3.AddNew
            rs_aux3!ges_gestion = Year(Date)
            rs_aux3!venta_codigo = var_cod
            rs_aux3!venta_codigo_det = VAR_AUX
            rs_aux3!bien_codigo = Ado_datos.Recordset!bien_codigo
            rs_aux3!venta_det_cantidad = Ado_datos.Recordset!cotiza_cantidad
            rs_aux3!venta_precio_unitario_bs = 0
            rs_aux3!venta_descuento_bs = 0
            rs_aux3!venta_precio_total_bs = 0
            rs_aux3!venta_precio_unitario_dol = 0
            rs_aux3!venta_descuento_dol = 0
            rs_aux3!venta_precio_total_dol = 0
            rs_aux3!concepto_venta = dtc_desc21.Text + " - " + Ado_datos.Recordset!bien_codigo
            'ok
            rs_aux3!grupo_codigo = "40000"
            rs_aux3!subgrupo_codigo = "43000"
            rs_aux3!par_codigo = "43340"
            'ok
            rs_aux3!tipo_descuento = 0
            rs_aux3!almacen_codigo = 0
            rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
            rs_aux3!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h
            rs_aux3!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x
            rs_aux3!modelo_elegido = "N"
            rs_aux3!modelo_elegido_h = "N"
            rs_aux3!modelo_elegido_x = "N"
            'rs_aux3!estado_codigo = "REG"
            rs_aux3!fecha_registro = Date
            rs_aux3!usr_codigo = glusuario
            rs_aux3.Update
        'End If
        'INI GRABA ALMACEN DETALLE (EN LA ENTREGA EN OBRA)
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        rs_aux4.Open "Select * from ao_almacen_detalle where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux4.RecordCount = 0 Then
'            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'            rs_aux4.AddNew
'            rs_aux4!almacen_codigo = 0
'            rs_aux4!bien_codigo = Ado_datos.Recordset!bien_codigo
'            rs_aux4!grupo_codigo = "40000"
'            rs_aux4!subgrupo_codigo = "43000"
'            rs_aux4!par_codigo = "43340"
'            rs_aux4!stock_ingreso = 1
'            rs_aux4!stock_salida = 0
'            rs_aux4!stock_actual = 1
'            rs_aux4!estado_codigo = "REG"
'            rs_aux4!usr_codigo = GlUsuario
'            rs_aux4!fecha_registro = Date
'            rs_aux4.Update
'        End If
        'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
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
        VAR_ARCH = "COM_" + RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
        rs_datos!archivo_respaldo_cargado = "N"
        'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo2 & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            rs_datos!doc_numero2 = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos!doc_codigo2) + "-") + LTrim(Str(rs_datos!doc_numero2))
        rs_datos!archivo_respaldo2 = VAR_ARCH2 + ".PDF"
        rs_datos!archivo_respaldo_cargado2 = "N"
        
        rs_datos!estado_codigo = "APR"
        rs_datos!fecha_registro = Date
        rs_datos!usr_codigo = glusuario
        rs_datos.UpdateBatch adAffectAll
      End If
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
        Call ABRIR_TABLA
        rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Enabled = False
        Fra_datos2.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
    End If

End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     rs_datos!cotiza_fecha = DTPfecha1.Value
     'A.
     rs_datos!trafico_codigo = dtc_codigo11.Text
     rs_datos!bien_codigo = dtc_codigo21.Text
     
     rs_datos!modelo_codigo = Txt_campo4.Caption     'dtc_codigo31.Text
     rs_datos!modelo_codigo_h = Txt_campo5.Caption   'dtc_codigo41.Text
     rs_datos!modelo_codigo_x = Txt_campo6.Caption   'dtc_codigo51.Text
    
     If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     rs_datos!cotiza_tdc_bol = txt_monto0.Text
     rs_datos!cotiza_precio_fob_dol = IIf(txt_monto2 = "", "0", txt_monto2)
     rs_datos!cotiza_precio_fob_bs = CDbl(txt_monto2) * CDbl(txt_monto0.Text)  'Txt_campo6.Text
     rs_datos!cotiza_precio_fob_dol_h = IIf(txt_monto6 = "", "0", txt_monto6)
     rs_datos!cotiza_precio_fob_bs_h = CDbl(txt_monto6) * CDbl(txt_monto0.Text)
     rs_datos!cotiza_precio_fob_dol_x = IIf(txt_monto10 = "", "0", txt_monto10)
     rs_datos!cotiza_precio_fob_bs_x = CDbl(txt_monto10) * CDbl(txt_monto0.Text)
     
     'costo_monto
     rs_datos!cotiza_energia = Txt_campo2.Text
     rs_datos!cotiza_luz = Txt_campo3.Text
     rs_datos!bien_cotiza_num_accesos = Txt_campo7.Text
     rs_datos!bien_cotiza_dimension_fosa_m = Txt_campo9.Text      'Txt_campo8.Text
     rs_datos!bien_cotiza_dimension_fosa_fondo = Txt_campo9.Text
     rs_datos!bien_cotiza_dimension_fosa_frente = Txt_campo10.Text
     
     rs_datos!cuadro_ctrl_codigo = IIf((dtc_codigo61.Text = ""), 1, dtc_codigo61.Text)
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     'rs_datos!hora_aprueba = ""
     rs_datos!proceso_codigo = "COM"
     rs_datos!subproceso_codigo = "COM-01"
     rs_datos!etapa_codigo = "COM-01-04"
     rs_datos!clasif_codigo = "COM"
     rs_datos!doc_codigo = "R-222"
     rs_datos!doc_numero = "0"  'txt_campo1.Text
     rs_datos!clasif_codigo2 = "COM"
     rs_datos!doc_codigo2 = "R-224"
     rs_datos!doc_numero2 = "0"  'txt_campo1.Text
     rs_datos!poa_codigo = "3.1.1"
     
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll

     Call ABRIR_TABLA
     rs_datos.MoveLast
'     mbDataChanged = False

     Fra_datos.Enabled = False
     Fra_datos2.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
'     dtc_codigo9.Enabled = True

  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'A.
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo21.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo4 = "") Then
    MsgBox "Debe registrar el Modelo1 del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo5 = "") Then
    MsgBox "Debe registrar el Modelo2 del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo6 = "") Then
    MsgBox "Debe registrar el Modelo3 del Equipo ... ", vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_campo3.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  
End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim IResult As Integer
    'Dim co As New ADODB.Command
    'CR01.ReportFileName = App.Path & "\Reportes\comercial\ar_cotizacion_equipos.rpt"
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R-222_ar_cotiza_venta_cliente.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!edif_codigo
    CR01.StoredProcParam(4) = Me.Ado_datos.Recordset!cotiza_codigo
    IResult = CR01.PrintReport
    If IResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim IResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\comercial\R-224_ar_cotiza_venta_cliente.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
      'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
      'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
    'Call CREAVISTAF11          'JQA JUN-2008
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!edif_codigo
    CR01.StoredProcParam(4) = Me.Ado_datos.Recordset!cotiza_codigo
    IResult = CR01.PrintReport
    If IResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    Fra_datos.Enabled = False
    Fra_datos2.Enabled = False

'    Select Case dtc_codigo2.Text
'        Case "1"
'        Case "2"
'        Case "3"
        'Call ABRIR_TABLA_DET
            
        aw_p_ao_solicitud_cotiza_detalle.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo") ' Nro. Negociacion (Cod.solicitud)
        aw_p_ao_solicitud_cotiza_detalle.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")    ' Codigo Unidad
        aw_p_ao_solicitud_cotiza_detalle.Txt_descripcion.Caption = Me.dtc_desc1.Text                        ' Descripcion Unidad
        aw_p_ao_solicitud_cotiza_detalle.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("cotiza_codigo")    ' Nro. Cotización
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo2.Caption = Me.Ado_detalle1.Recordset("edif_codigo")      ' Codigo Edificio
        
        aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("codigo_costo")     ' Codigo Costo
        aw_p_ao_solicitud_cotiza_detalle.dtc_desc1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
        aw_p_ao_solicitud_cotiza_detalle.dtc_aux1.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
        aw_p_ao_solicitud_cotiza_detalle.Dtc_aux2.BoundText = aw_p_ao_solicitud_cotiza_detalle.dtc_codigo1.BoundText
        
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo3.Text = Me.Ado_detalle1.Recordset("costo_porcentaje")    ' % Costo
               
        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = "0"                  ' Monto Modelo1(ME)
        Else
            aw_p_ao_solicitud_cotiza_detalle.txt_monto01.Caption = Me.txt_monto1.Text   ' Monto Modelo1(ME)
        End If
        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
            aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = "0"                  ' Monto Modelo2(ME)
        Else
            aw_p_ao_solicitud_cotiza_detalle.txt_monto02.Caption = Me.txt_monto5.Text   ' Monto Modelo2(ME)
        End If
        If txt_monto1.Text = "0" Or txt_monto1.Text = "" Then
            aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = "0"                  ' Monto Modelo3(ME)
        Else
            aw_p_ao_solicitud_cotiza_detalle.txt_monto03.Caption = Me.txt_monto9.Text   ' Monto Modelo3(ME)
        End If
        
        aw_p_ao_solicitud_cotiza_detalle.Txt_campo4.Text = Me.Ado_detalle1.Recordset("costo_observaciones") ' Observaciones
        
        
        aw_p_ao_solicitud_cotiza_detalle.Show vbModal
'        Case "4"
'
'    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        Fra_datos2.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        txt_monto1.SetFocus
    '    BtnVer.Visible = True
        'dtc_codigo9.Enabled = False
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer_Click()
    'ARREGLO 1
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc11 = dtc_aux41.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc21 = dtc_aux51.Text
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoc31 = IIf(IsNull(Ado_datos.Recordset!trafico_c_time_entrada_salida), 0, Ado_datos.Recordset!trafico_c_time_entrada_salida)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campod11 = IIf(IsNull(Ado_datos.Recordset!trafico_d_num_paradas_probables), 0, Ado_datos.Recordset!trafico_d_num_paradas_probables)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe11 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_e_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe21 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_e_tiempo_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe31 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre), 0, Ado_datos.Recordset!trafico_e_tiempo_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campoe41 = IIf(IsNull(Ado_datos.Recordset!trafico_e_tiempo_entrada_salida), 0, Ado_datos.Recordset!trafico_e_tiempo_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof11 = IIf(IsNull(Ado_datos.Recordset!trafico_f_tiempo_recorrido), 0, Ado_datos.Recordset!trafico_f_tiempo_recorrido)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof21 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_asc_desaceleracion), 0, Ado_datos.Recordset!trafico_f_time_asc_desaceleracion)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof31 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_apertura_cierre), 0, Ado_datos.Recordset!trafico_f_time_apertura_cierre)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campof41 = IIf(IsNull(Ado_datos.Recordset!trafico_f_time_entrada_salida), 0, Ado_datos.Recordset!trafico_f_time_entrada_salida)
'
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog11 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti), 0, Ado_datos.Recordset!trafico_g_capacidad_tiempo_cti)
'    aw_p_ao_solicitud_calculo_trafico_det.lbl_campog21 = IIf(IsNull(Ado_datos.Recordset!trafico_g_capacidad_total_arreglo), 0, Ado_datos.Recordset!trafico_g_capacidad_total_arreglo)
    
End Sub


Private Sub CmdMod1_Click()
    'Modelo 1
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "Select * from av_solicitud_cotiza_modelo where modelo_codigo = '" & Txt_campo4 & "'", db, adOpenStatic
    If rs_aux1.RecordCount > 0 Then
        FraModelo.Visible = True
        FraModeloCosto.Visible = False
    Else
        MsgBox "No Existe el Modelo asociado al Registro, debe Registrarlo para ver sus Características ...", vbExclamation, "Advertencia"
    End If
End Sub

Private Sub CmdMod2_Click()
    'Modelo 2
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "Select * from av_solicitud_cotiza_modelo where modelo_codigo = '" & Txt_campo5 & "'", db, adOpenStatic
    If rs_aux2.RecordCount > 0 Then
        FraModelo.Visible = True
        FraModeloCosto.Visible = False
    Else
        MsgBox "No Existe el Modelo asociado al Registro, debe Registrarlo para ver sus Características ...", vbExclamation, "Advertencia"
    End If
End Sub

Private Sub CmdMod3_Click()
    'Modelo 3
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from av_solicitud_cotiza_modelo where modelo_codigo = '" & Txt_campo6 & "'", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        FraModelo.Visible = True
        FraModeloCosto.Visible = False
    Else
        MsgBox "No Existe el Modelo asociado al Registro, debe Registrarlo para ver sus Características ...", vbExclamation, "Advertencia"
    End If
End Sub

Private Sub CmdVolver_Click()
    FraModelo.Visible = False
    FraModeloCosto.Visible = True
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    dtc_desc12.BoundText = dtc_codigo11.BoundText
    dtc_desc13.BoundText = dtc_codigo11.BoundText
    dtc_desc14.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
    dtc_desc12.BoundText = dtc_desc11.BoundText
    dtc_desc13.BoundText = dtc_desc11.BoundText
    dtc_desc14.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc12_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_desc12.BoundText
    dtc_codigo11.BoundText = dtc_desc12.BoundText
    dtc_desc13.BoundText = dtc_desc12.BoundText
    dtc_desc14.BoundText = dtc_desc12.BoundText
End Sub

Private Sub dtc_desc13_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_desc13.BoundText
    dtc_desc12.BoundText = dtc_desc13.BoundText
    dtc_codigo11.BoundText = dtc_desc13.BoundText
    dtc_desc14.BoundText = dtc_desc13.BoundText
End Sub

Private Sub dtc_desc14_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_desc14.BoundText
    dtc_desc12.BoundText = dtc_desc14.BoundText
    dtc_desc13.BoundText = dtc_desc14.BoundText
    dtc_codigo11.BoundText = dtc_desc14.BoundText
End Sub

Private Sub dtc_desc24_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_desc24.BoundText
    dtc_desc22.BoundText = dtc_desc24.BoundText
    dtc_desc23.BoundText = dtc_desc24.BoundText
    dtc_codigo21.BoundText = dtc_desc24.BoundText
End Sub

Private Sub dtc_desc34_Click(Area As Integer)
    dtc_desc31.BoundText = dtc_desc34.BoundText
    dtc_desc32.BoundText = dtc_desc34.BoundText
    dtc_desc33.BoundText = dtc_desc34.BoundText
    dtc_codigo31.BoundText = dtc_desc34.BoundText
End Sub

Private Sub dtc_desc44_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_desc44.BoundText
    dtc_desc42.BoundText = dtc_desc44.BoundText
    dtc_desc43.BoundText = dtc_desc44.BoundText
    dtc_codigo41.BoundText = dtc_desc44.BoundText
End Sub

Private Sub dtc_desc54_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_desc54.BoundText
    dtc_desc52.BoundText = dtc_desc54.BoundText
    dtc_desc53.BoundText = dtc_desc54.BoundText
    dtc_codigo51.BoundText = dtc_desc54.BoundText
End Sub

Private Sub dtc_desc64_Click(Area As Integer)
'    dtc_codigo64.BoundText = dtc_codigo64.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    parametro = aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    
    Fra_datos.Enabled = False
    Fra_datos2.Enabled = False
    dg_datos.Enabled = True
    'lbl_aux1.Visible = False
'    FraNavega.Caption = lbl_titulo.Caption
'    lbl_titulo2.Caption = lbl_titulo.Caption
   
   'If Not Ado_datos.Recordset.EOF Then
            'SSTab1.Tab = 0
            'SSTab1.TabEnabled(0) = True
            ''SSTab1.TabEnabled(1) = False
            'SSTab1.TabVisible(1) = False
   'End If
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
        
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'Cálculo de Tráfico.
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    rs_datos11.Open "Select * from ao_solicitud_calculo_trafico where estado_codigo = 'APR' ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
    'Bien (Equipo)
    Set rs_datos21 = New ADODB.Recordset
    If rs_datos21.State = 1 Then rs_datos21.Close
    rs_datos21.Open "Select * from ac_bienes ", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos21.Recordset = rs_datos21
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    'Modelo 1
    Set rs_datos31 = New ADODB.Recordset
    If rs_datos31.State = 1 Then rs_datos31.Close
    rs_datos31.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo = 'BRA'", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos31.Recordset = rs_datos31
    dtc_desc31.BoundText = dtc_codigo31.BoundText
    'Modelo 2
    Set rs_datos41 = New ADODB.Recordset
    If rs_datos41.State = 1 Then rs_datos41.Close
    rs_datos41.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo <> 'BRA' AND pais_codigo <> 'CHN' ", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos41.Recordset = rs_datos41
    dtc_desc41.BoundText = dtc_codigo41.BoundText
    'Modelo 3
    Set rs_datos51 = New ADODB.Recordset
    If rs_datos51.State = 1 Then rs_datos51.Close
    rs_datos51.Open "Select * from av_solicitud_cotiza_modelo where pais_codigo = 'CHN' ", db, adOpenStatic
    'rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos51.Recordset = rs_datos51
    dtc_desc51.BoundText = dtc_codigo51.BoundText
    'Cuadro de Control
    Set rs_datos61 = New ADODB.Recordset
    If rs_datos61.State = 1 Then rs_datos61.Close
    rs_datos61.Open "Select * from ac_bienes_equipo_cuadro_ctrl ", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos61.Recordset = rs_datos61
    dtc_desc61.BoundText = dtc_codigo61.BoundText
    
End Sub

Private Sub Maximo_Numerador()
'  TxtCrr.Text = "1"
'  Set RsTmp = New ADODB.Recordset
''  Set rst_ben = New ADODB.Recordset
''  rst_ben.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", DB, adOpenStatic
''  Set AdoTip_ben.Recordset = rst_ben
'  RsTmp.Open "Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ", db, adOpenStatic
'  'Set RsTmp = DbConex.Execute("Select max(trafico_codigo) + 1 as Codigo from ao_solicitud_ctrl_trafico ;")
'  If Not RsTmp.EOF Then
'     TxtCrr.Text = RsTmp!Codigo
'  End If
End Sub

Private Sub Carga_Beneficiario()
'  Set rstbeneficiario = New ADODB.Recordset
'  If rstbeneficiario.State = 1 Then rstbeneficiario.Close
'  sql = "SELECT ges_gestion as gestion,unidad_codigo as Unid_Ejec,solicitud_codigo as Codigo,trafico_codigo,estado_codigo,edif_codigo,trafico_num_paradas,trafico_recorrido," _
'  & " trafico_nro_equipos,vel_equipo_codigo,tipo_puerta,trafico_ancho_puerta,cabina_codigo," _
'  & " tecnologia_codigo , sist_puerta, condicion_ventas " _
'  & " From ao_solicitud_ctrl_trafico WHERE estado_codigo = 'REG'"
''  SQL = "Select ges_gestion,unidad_codigo,solicitud_codigo,trafico_codigo from ao_solicitud_ctrl_trafico order by unidad_codigo,solicitud_codigo,trafico_codigo"
'  rstbeneficiario.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
'  Set Ado_datos.Recordset = rstbeneficiario
'  'Ado_datos.ConnectionString = sConex
'  'Ado_datos.RecordSource = SQL
'  'Ado_datos.Refresh
'
'  dg_datos.Columns(0).Width = 800 'maxWidth
'  dg_datos.Columns(1).Width = 1556
'  dg_datos.Columns(2).Width = 1556
'  dg_datos.Columns(4).Alignment = dbgRight
''  dg_datos.Columns(2).Alignment = dbgRight
''  dg_datos.Columns(3).Alignment = dbgRight
''  dg_datos.Columns(4).Alignment = dbgCenter
''  dg_datos.Columns(2).NumberFormat = ("###0.00")
''  dg_datos.Columns(3).NumberFormat = ("###0.00")
'
'  'LblReg.Caption = "Total Registros --> " & Ado_datos.Recordset.RecordCount
End Sub

Function Llena_Combos()
'  CmbReco.Clear
'  sql = " SELECT recorrido_descripcion From ac_bienes_equipo_recorrido; "
'  If RsTmp.State = 1 Then RsTmp.Close
'  RsTmp.Open sql, db, adOpenStatic
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbReco.AddItem RsTmp!recorrido_descripcion
'         RsTmp.MoveNext
'     Wend
'  End If
''---
'  CmbNroPasaj.Clear
'  sql = " SELECT pasajeros_descripcion From ac_bienes_equipo_nro_pasajeros; "
'  If RsTmp.State = 1 Then RsTmp.Close
'  RsTmp.Open sql, db, adOpenStatic
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbNroPasaj.AddItem RsTmp!pasajeros_descripcion
'         RsTmp.MoveNext
'     Wend
'  End If
''---
''  CmbVelEq.Clear
''  SQL = " SELECT vel_equipo_descripcion From ac_bienes_equipo_velocidad WHERE vel_equipo_codigo = " & nCod & "; "
''  If RsTmp.State = 1 Then RsTmp.Close
''  RsTmp.Open SQL, DB, adOpenStatic
''  If Not RsTmp.EOF Then
''     While Not (RsTmp.EOF)
''           CmbVelEq.AddItem RsTmp!pasajeros_descripcion
''         RsTmp.MoveNext
''     Wend
''  End If
End Function

Function Llena_Clientes1()
'  CmbCodCli1.Clear
'  CmbCliente.Clear
'  Call ABRE_CONECCION
'  Set RsTmp = DbConex.Execute("select * from CLIENTES order by nomBRECLI ;")
'  If Not RsTmp.EOF Then
'     While Not (RsTmp.EOF)
'           CmbCodCli1.AddItem RsTmp!CodCli
'           CmbCliente.AddItem RsTmp!nombrecli
'         RsTmp.MoveNext
'     Wend
'  End If
'  Call CERRAR_CONECCION
End Function

Private Sub CmbCliente_Click()
' If CmbCliente.ListIndex = -1 Then Exit Sub
' CmbCodCli1.ListIndex = CmbCliente.ListIndex
End Sub

Private Sub dg_datos_Click()
'  MsgBox "sss"
'   Call Llena_Varios
'  txtDescrip = dg_datos.Columns(1).Text
End Sub

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
Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
'    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo21_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_codigo21.BoundText
    dtc_desc22.BoundText = dtc_codigo21.BoundText
    dtc_desc23.BoundText = dtc_codigo21.BoundText
    dtc_desc24.BoundText = dtc_codigo21.BoundText
End Sub

Private Sub dtc_codigo22_Click(Area As Integer)
'    dtc_desc22.BoundText = dtc_codigo22.BoundText
End Sub

Private Sub dtc_codigo23_Click(Area As Integer)
    'dtc_desc23.BoundText = dtc_codigo23.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    'dtc_desc3.BoundText = dtc_codigo3.BoundText
    'dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_codigo31_Click(Area As Integer)
    dtc_desc31.BoundText = dtc_codigo31.BoundText
    dtc_desc32.BoundText = dtc_codigo31.BoundText
    dtc_desc33.BoundText = dtc_codigo31.BoundText
    dtc_desc34.BoundText = dtc_codigo31.BoundText
End Sub

Private Sub dtc_codigo41_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_codigo41.BoundText
    dtc_desc42.BoundText = dtc_codigo41.BoundText
    dtc_desc43.BoundText = dtc_codigo41.BoundText
    dtc_desc44.BoundText = dtc_codigo41.BoundText
End Sub

Private Sub dtc_codigo51_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_codigo51.BoundText
    dtc_desc52.BoundText = dtc_codigo51.BoundText
    dtc_desc53.BoundText = dtc_codigo51.BoundText
    dtc_desc54.BoundText = dtc_codigo51.BoundText
End Sub

Private Sub dtc_codigo61_Click(Area As Integer)
    dtc_desc61.BoundText = dtc_codigo61.BoundText
    dtc_desc62.BoundText = dtc_codigo61.BoundText
End Sub
'
Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    dtc_aux1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

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

Private Sub dtc_desc21_Click(Area As Integer)
    dtc_codigo21.BoundText = dtc_desc21.BoundText
    dtc_desc22.BoundText = dtc_desc21.BoundText
    dtc_desc23.BoundText = dtc_desc21.BoundText
    dtc_desc24.BoundText = dtc_desc21.BoundText
End Sub

Private Sub dtc_desc22_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_desc22.BoundText
    dtc_codigo21.BoundText = dtc_desc22.BoundText
    dtc_desc23.BoundText = dtc_desc22.BoundText
    dtc_desc24.BoundText = dtc_desc22.BoundText
End Sub

Private Sub dtc_desc23_Click(Area As Integer)
    dtc_desc21.BoundText = dtc_desc23.BoundText
    dtc_desc22.BoundText = dtc_desc23.BoundText
    dtc_codigo21.BoundText = dtc_desc23.BoundText
    dtc_desc24.BoundText = dtc_desc23.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub dtc_desc31_Click(Area As Integer)
    dtc_codigo31.BoundText = dtc_desc31.BoundText
    dtc_desc32.BoundText = dtc_desc31.BoundText
    dtc_desc33.BoundText = dtc_desc31.BoundText
    dtc_desc34.BoundText = dtc_desc31.BoundText
End Sub

Private Sub dtc_desc32_Click(Area As Integer)
    dtc_desc31.BoundText = dtc_desc32.BoundText
    dtc_codigo31.BoundText = dtc_desc32.BoundText
    dtc_desc33.BoundText = dtc_desc32.BoundText
    dtc_desc34.BoundText = dtc_desc32.BoundText
End Sub

Private Sub dtc_desc33_Click(Area As Integer)
    dtc_desc31.BoundText = dtc_desc33.BoundText
    dtc_desc32.BoundText = dtc_desc33.BoundText
    dtc_codigo31.BoundText = dtc_desc33.BoundText
    dtc_desc34.BoundText = dtc_desc33.BoundText
End Sub

Private Sub dtc_desc41_Click(Area As Integer)
    dtc_codigo41.BoundText = dtc_desc41.BoundText
    dtc_desc42.BoundText = dtc_desc41.BoundText
    dtc_desc43.BoundText = dtc_desc41.BoundText
    dtc_desc44.BoundText = dtc_desc41.BoundText
End Sub

Private Sub dtc_desc42_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_desc42.BoundText
    dtc_codigo41.BoundText = dtc_desc42.BoundText
    dtc_desc43.BoundText = dtc_desc42.BoundText
    dtc_desc44.BoundText = dtc_desc42.BoundText
End Sub

Private Sub dtc_desc43_Click(Area As Integer)
    dtc_desc41.BoundText = dtc_desc43.BoundText
    dtc_desc42.BoundText = dtc_desc43.BoundText
    dtc_codigo41.BoundText = dtc_desc43.BoundText
    dtc_desc44.BoundText = dtc_desc43.BoundText
End Sub

Private Sub dtc_desc51_Click(Area As Integer)
    dtc_codigo51.BoundText = dtc_desc51.BoundText
    dtc_desc52.BoundText = dtc_desc51.BoundText
    dtc_desc53.BoundText = dtc_desc51.BoundText
    dtc_desc54.BoundText = dtc_desc51.BoundText
End Sub

Private Sub dtc_desc52_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_desc52.BoundText
    dtc_codigo51.BoundText = dtc_desc52.BoundText
    dtc_desc53.BoundText = dtc_desc52.BoundText
    dtc_desc54.BoundText = dtc_desc52.BoundText
End Sub

Private Sub dtc_desc53_Click(Area As Integer)
    dtc_desc51.BoundText = dtc_desc53.BoundText
    dtc_desc52.BoundText = dtc_desc53.BoundText
    dtc_codigo51.BoundText = dtc_desc53.BoundText
    dtc_desc54.BoundText = dtc_desc53.BoundText
End Sub

Private Sub dtc_desc61_Click(Area As Integer)
    dtc_codigo61.BoundText = dtc_desc61.BoundText
    dtc_desc62.BoundText = dtc_desc61.BoundText
End Sub

Private Sub dtc_desc62_Click(Area As Integer)
    dtc_codigo61.BoundText = dtc_desc62.BoundText
    dtc_desc61.BoundText = dtc_desc62.BoundText
End Sub

Private Sub OptFilGral1_Click()
  '===== Proceso para filtrado general de datos(registros no aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From ao_solicitud_cotiza_venta WHERE estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
    'queryinicial = "Select * from ao_solicitud where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
  '===== Proceso para filtrado general de datos (todos los registros )
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From av_ventas_cabecera "
    queryinicial = "Select * from ao_solicitud_cotiza_venta where  unidad_codigo = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLA()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from ao_solicitud_cotiza_venta where " + parametro
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
        
    dtc_desc31.BoundText = dtc_codigo31.BoundText
    dtc_desc32.BoundText = dtc_codigo31.BoundText
    dtc_desc33.BoundText = dtc_codigo31.BoundText
    dtc_desc34.BoundText = dtc_codigo31.BoundText
    
    dtc_desc41.BoundText = dtc_codigo41.BoundText
    dtc_desc42.BoundText = dtc_codigo41.BoundText
    dtc_desc43.BoundText = dtc_codigo41.BoundText
    dtc_desc44.BoundText = dtc_codigo41.BoundText
    
    dtc_desc51.BoundText = dtc_codigo51.BoundText
    dtc_desc52.BoundText = dtc_codigo51.BoundText
    dtc_desc53.BoundText = dtc_codigo51.BoundText
    dtc_desc54.BoundText = dtc_codigo51.BoundText
End Sub

Private Sub FraModelo_Click()
    FraModelo.Visible = False
    FraModeloCosto.Visible = True
End Sub

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
  If glPersNew = "P" Then
  End If
  glPersNew = "N"
   
'   If (rstbeneficiario.State = adStateClosed) Then rstbeneficiario.Close
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
End Sub

Private Sub txt_monto1_LostFocus()
     If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto1 = "" Or txt_monto1 = "0" Then
        txt_monto2.Text = "0"
     Else
        txt_monto2.Text = CDbl(txt_monto1) / CDbl(txt_monto0.Text)
     End If
End Sub

Private Sub txt_monto10_LostFocus()
    If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto10 = "" Then
        txt_monto9.Text = "0"
     Else
        txt_monto9 = CDbl(txt_monto10) * CDbl(txt_monto0.Text)
     End If
     
End Sub

Private Sub txt_monto2_LostFocus()
    If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto2 = "" Then
        txt_monto1.Text = "0"
     Else
        txt_monto1.Text = CDbl(txt_monto2) * CDbl(txt_monto0.Text)
     End If
     
End Sub

Private Sub txt_monto5_LostFocus()
    If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto5 = "" Or txt_monto5 = "0" Then
        txt_monto6.Text = "0"
     Else
        txt_monto6.Text = CDbl(txt_monto5) / CDbl(txt_monto0.Text)
     End If
End Sub

Private Sub txt_monto6_LostFocus()
    If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto6 = "" Then
        txt_monto5.Text = "0"
     Else
        txt_monto5.Text = CDbl(txt_monto6) * CDbl(txt_monto0.Text)
     End If
     
End Sub

Private Sub txt_monto9_LostFocus()
     If txt_monto0.Text = "0" Or txt_monto0.Text = "" Then
        txt_monto0.Text = GlTipoCambioOficial
     End If
     If txt_monto9 = "" Or txt_monto9 = "0" Then
        txt_monto10.Text = "0"
     Else
        txt_monto10 = CDbl(txt_monto9) / CDbl(txt_monto0.Text)
     End If
End Sub

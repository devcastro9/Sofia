VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ao_solicitud_servicio_basico 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Financieros - Ejecucion dej gastos - Solicitud Pago x Servicio Basico y Otros"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frm_ao_solicitud_servicio_basico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   64
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":0E44
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnA�adir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":104E
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   71
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":180D
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   70
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":2122
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   69
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6960
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":286E
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   68
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":30A1
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   67
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":3856
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   66
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":4123
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   65
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   12855
         TabIndex        =   74
         Top             =   195
         Width           =   1815
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   676
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   20280
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":48E5
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   62
         Top             =   0
         Width           =   1335
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":50BB
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   61
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   13215
         TabIndex        =   63
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "frm_ao_solicitud_servicio_basico.frx":59A7
      ScaleHeight     =   1425
      ScaleWidth      =   1875
      TabIndex        =   53
      Top             =   5440
      Width           =   1935
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Solicitud"
         Height          =   640
         Left            =   960
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":719D9
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":7315B
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":7359D
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":739DF
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Adiciona Detalle"
         Top             =   60
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1490
      Left            =   120
      Picture         =   "frm_ao_solicitud_servicio_basico.frx":73E21
      ScaleHeight     =   1425
      ScaleWidth      =   1875
      TabIndex        =   48
      Top             =   7080
      Width           =   1935
      Begin VB.CommandButton BtnImprimir1 
         BackColor       =   &H80000018&
         Caption         =   "Bit�cora"
         Height          =   640
         Left            =   980
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":DFE53
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Imprime Nota de Venta"
         Top             =   740
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":E15D5
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Adiciona Detalle"
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   960
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":E1A17
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   120
         Picture         =   "frm_ao_solicitud_servicio_basico.frx":E1E59
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   740
         Width           =   765
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2140
      TabIndex        =   44
      Top             =   6980
      Width           =   12855
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E229B
         Height          =   1215
         Left            =   240
         TabIndex        =   75
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            Caption         =   "Unidad"
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Solicitud"
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
            DataField       =   "bitacora_codigo"
            Caption         =   "Bitacora"
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
            DataField       =   "beneficiario_codigo"
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
         BeginProperty Column05 
            DataField       =   "negocia_observaciones"
            Caption         =   "Observacion"
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
            DataField       =   "beneficiario_nombre_ref"
            Caption         =   "Ref Beneficiario"
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
            DataField       =   "beneficiario_codigo_cgi"
            Caption         =   "Beneficiario Cgi"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2849.953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2475.213
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "DETALLE DE LA SOLICITUD"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2140
      TabIndex        =   23
      Top             =   5355
      Width           =   12855
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E22B6
         Height          =   1215
         Left            =   195
         TabIndex        =   24
         Top             =   240
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo de Servicio"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Unidad Medida"
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
         BeginProperty Column02 
            DataField       =   "bien_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "bien_precio_compra"
            Caption         =   "Costo Unitario Bs."
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
            DataField       =   "bien_total_compra"
            Caption         =   "Monto Total Bs."
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Servicio"
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
               Locked          =   -1  'True
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   6105.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "LISTADO"
      ForeColor       =   &H00FFFFFF&
      Height          =   4080
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E22D1
         Height          =   3330
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   5874
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
            Weight          =   700
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
            DataField       =   "solicitud_codigo"
            Caption         =   "Tr�mite"
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
            DataField       =   "beneficiario_codigo"
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
         BeginProperty Column03 
            DataField       =   "solicitud_fecha_solicitud"
            Caption         =   "Fecha.Reg."
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
            DataField       =   "estado_cotiza"
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
         BeginProperty Column05 
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
               Alignment       =   2
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   39
         Top             =   3700
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00FFFFFF&
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
         Left            =   3600
         TabIndex        =   40
         Top             =   3700
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3640
         Width           =   5625
         _ExtentX        =   9922
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
         BackColor       =   16777215
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   4080
      Left            =   6120
      TabIndex        =   12
      Top             =   720
      Width           =   8895
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   8400
         TabIndex        =   47
         Top             =   1800
         Width           =   270
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6480
         TabIndex        =   43
         Top             =   525
         Width           =   270
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   8500
         TabIndex        =   38
         Top             =   1090
         Width           =   260
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2520
         Visible         =   0   'False
         Width           =   1605
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E22E9
         DataField       =   "beneficiario_codigo_resp2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E2303
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_sigla"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E231C
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E2335
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   2925
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E234F
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2400
         TabIndex        =   16
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
      Begin VB.TextBox Txt_descripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_justificacion"
         DataSource      =   "Ado_datos"
         Height          =   555
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2205
         Width           =   7185
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   795
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
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E2368
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5160
         TabIndex        =   18
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
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E2381
         DataField       =   "solicitud_tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   1800
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   ""
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
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E239A
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   510
         Width           =   4605
         _ExtentX        =   8123
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
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E23B3
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Top             =   2925
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E23CD
         DataField       =   "beneficiario_codigo_resp2"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_ao_solicitud_servicio_basico.frx":E23E7
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2280
         TabIndex        =   58
         Top             =   1080
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   4800
         TabIndex        =   59
         Top             =   3660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   42074113
         CurrentDate     =   42043
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin VB.Label dtc_codigo9 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "0"
         DataField       =   "doc_codigo"
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
         Left            =   180
         TabIndex        =   46
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo Tr�mite"
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
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite del Tr�mite"
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
         Index           =   6
         Left            =   6960
         TabIndex        =   42
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Txt_campo2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "36NO"
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
         Left            =   7020
         TabIndex        =   41
         Top             =   510
         Width           =   1695
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Concepto:"
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
         TabIndex        =   37
         Top             =   2220
         Width           =   915
      End
      Begin VB.Label lbl_campo10 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Actividad POA"
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
         TabIndex        =   36
         Top             =   2940
         Width           =   1305
      End
      Begin VB.Label lbl_campo9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "C�digo Registro"
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
         TabIndex        =   35
         Top             =   3375
         Width           =   1470
      End
      Begin VB.Label lbl_campo11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Responsable CGI"
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
         TabIndex        =   34
         Top             =   1470
         Width           =   1605
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Beneficiario"
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
         TabIndex        =   33
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Unidad Ejecutora"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   225
         Width           =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   8880
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   180
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   2460
         TabIndex        =   27
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Nro. Doc. Respaldo"
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
         Index           =   13
         Left            =   2520
         TabIndex        =   22
         Top             =   3375
         Width           =   1785
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha de Registro"
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
         Index           =   12
         Left            =   4800
         TabIndex        =   21
         Top             =   3375
         Width           =   1665
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   7080
         TabIndex        =   5
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Tr�mite"
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
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Estado"
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
         Index           =   2
         Left            =   7080
         TabIndex        =   13
         Top             =   3375
         Width           =   765
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
      ScaleWidth      =   20250
      TabIndex        =   6
      Top             =   10950
      Width           =   20250
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   11
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   240
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
   Begin Crystal.CrystalReport CR01 
      Left            =   7200
      Top             =   11040
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2640
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
   Begin MSAdodcLib.Adodc Ado_datos3 
      Height          =   330
      Left            =   4800
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   7080
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
   Begin MSAdodcLib.Adodc Ado_datos5 
      Height          =   330
      Left            =   9360
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
      Left            =   11640
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
      Left            =   13920
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
      Left            =   240
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
      Left            =   2640
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
      Left            =   4800
      Top             =   9600
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
   Begin MSAdodcLib.Adodc Ado_detalle1 
      Height          =   330
      Left            =   240
      Top             =   8760
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2520
      Top             =   8760
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
      Caption         =   "Ado_detalle2"
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
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   7080
      Top             =   9600
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
   Begin MSAdodcLib.Adodc Ado_datos12 
      Height          =   330
      Left            =   9360
      Top             =   9600
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
      Caption         =   "Ado_datos12"
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
Attribute VB_Name = "frm_ao_solicitud_servicio_basico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset
Dim rs_datos10 As New ADODB.Recordset
Dim rs_datos11 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String

Dim VAR_REG As Integer
Dim VAR_CMPBTE As Integer
Dim VAR_AUX, VAR_CONT2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean
Private Sub Pagos(Unidad, formulario, org_codigo, solicitud_codigo, justificacion, observaciones, beneficiario_codigo, concepto_pago, obs_fo_gastos_detalle, mon_pago As String)
  'PAGOS
    'WWWWWWWWWWWWWW
    Dim VAR_CMPBTE As Integer
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        'VAR_COD4 = parametro    'UNIDAD
        'VAR_SOL = Ado_datos.Recordset!beneficiario_codigo  '
        'tipo_formulario TERCER PARAMETRO
        'org_codigo CUARTO PARAMETRO
        'ini generaci�n de correlativo
        
        If Unidad = "DRRHH" Then
        Set rs_aux5 = New ADODB.Recordset
        If rs_aux5.State = 1 Then rs_aux5.Close
        rs_aux5.Open "select * from ro_rrhh_adjudica_personas where beneficiario_codigo = '" & beneficiario_codigo & "'", db, adOpenKeyset, adLockOptimistic
        solicitud_codigo = rs_aux5!solicitud_codigo
        Unidad = rs_aux5!unidad_codigo
        End If
        
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from fc_organismo_financiamiento where org_codigo = '" & org_codigo & "'", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
           rs_aux2!correlativo_gasto = rs_aux2!correlativo_gasto + 1
           VAR_CMPBTE = rs_aux2!correlativo_gasto
           rs_aux2.Update
        End If
        
        'WWWWWWWWWWWWWWW
        'correlv = Ado_datos.Recordset!venta_codigo
        'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        rs_aux3.Open "select * from  fo_gastos_cabecera where unidad_codigo = '" & Unidad & "' AND solicitud_codigo = " & solicitud_codigo & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux3.RecordCount = 0 Then
            rs_aux3.AddNew
            rs_aux3!ges_gestion = Year(Date) 'glGestion     'Year(Date)
            rs_aux3!org_codigo = org_codigo
            rs_aux3!gasto_codigo = VAR_CMPBTE
            rs_aux3!tipo_comp = "DEV"
            rs_aux3!gasto_codigo_anterior = VAR_CMPBTE
            rs_aux3!unidad_codigo = Unidad
            rs_aux3!solicitud_codigo = solicitud_codigo
            rs_aux3!tipo_formulario = formulario     'GC_TIPO_SOLICITUD
            rs_aux3!pago_codigo = rs_aux3.RecordCount + 1
          
            
            rs_aux3!proceso_codigo = "FIN"
            rs_aux3!subproceso_codigo = "FIN-03"
            rs_aux3!etapa_codigo = "FIN-03-03"
            rs_aux3!clasif_codigo = "ADM"
            rs_aux3!doc_codigo = "R-111"
            rs_aux3!doc_numero = 0
            rs_aux3!poa_codigo = "4.2.3"
            
            rs_aux3!fecha_egreso = Date
            rs_aux3!tipo_moneda = "BOB"
            rs_aux3!da_codigo = "1.1"
            
            rs_aux3!fte_codigo = "10"   'REVISAR DE LA TABLA fc_organismo_financiamiento
            rs_aux3!monto_Bolivianos = 0
            rs_aux3!monto_dolares = 0
            rs_aux3!liquido_pagar = 0
            rs_aux3!monto_Bolivianos_pag = 0
            rs_aux3!monto_dolares_pag = 0
            rs_aux3!Deducciones = 0
            rs_aux3!fecha_autorizacion = Date
            rs_aux3!justificacion = justificacion
            rs_aux3!es_base = "S"
            
            rs_aux3!CODIGO_GRUPO = VAR_CMPBTE   'rs_aux3!pago_codigo
            rs_aux3!NUMERO_PAGO = 1
            rs_aux3!observaciones = observaciones
            
            'rs_aux3!edif_codigo = VAR_PROY2
            'rs_aux3!beneficiario_codigo = VAR_BENEF
            'rs_aux3!solicitud_tipo = "10"
            'rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
            
            rs_aux3!estado_devengado = "APR"
            rs_aux3!estado_pagado = "REG"
            rs_aux3!estado_contabilidad = "REG"
            
            
            rs_aux3!estado_codigo = "REG"
            rs_aux3!usr_codigo = glusuario
            rs_aux3!fecha_registro = Date
            rs_aux3!usr_codigo_aprueba = glusuario
            rs_aux3!fecha_aprueba = Date
            rs_aux3.Update
            
            'DETALLE Carga fo_gastos_detalle
            
            Set rstdestino = New ADODB.Recordset
            If rstdestino.State = 1 Then rstdestino.Close
            rstdestino.Open "select * from fo_gastos_detalle where org_codigo = '" & org_codigo & "' AND gasto_codigo= " & VAR_CMPBTE & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rstdestino.RecordCount > 0 Then
            End If

            Set rs_aux4 = New ADODB.Recordset
            If rs_aux4.State = 1 Then rs_aux4.Close
            rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & Unidad & "' AND solicitud_codigo = " & solicitud_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux4.RecordCount > 0 Then
               VAR_REG = 1
               rs_aux4.MoveFirst
               Dim bien_total_compra As Double
               Dim par_codigo As String
               If Unidad = DRRHH Then
               bien_total_compra = mon_pago
               par_codigo = ""
               Else
               bien_total_compra = rs_aux4!bien_total_compra
               par_codigo = rs_aux4!par_codigo
               End If
               While Not rs_aux4.EOF
               '     db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
               '     "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
               '     rs_aux4.MoveNext
                    db.Execute "INSERT INTO fo_gastos_detalle (ges_gestion, org_codigo, gasto_codigo, gasto_codigo_detalle, par_codigo, pro_codigo, codigo_beneficiario, concepto_pago, monto_total, monto_dolares_dev, tipo_cambio_dev, monto_Bolivianos, monto_Dolares, saldo_bolivianos, tipo_cambio, Porcentaje, deducciones, fecha_pago, depto_codigo, estado_aprobacion, fecha_autorizacion, Observacion, codigo_dev, usr_usuario, fecha_registro, hora_registro,  estado_conciliacion, codigo_poa " & _
                    "VALUES ('" & glGestion & "','" & org_codigo & "', " & VAR_CMPBTE & ", '" & VAR_REG & "', " & par_codigo & ", '8', '" & beneficiario_codigo & "', '" & concepto_pago & "', " & bien_total_compra & ", " & bien_total_compra * GlTipoCambioOficial & ", " & GlTipoCambioOficial & ", " & bien_total_compra & ", " & bien_total_compra * GlTipoCambioOficial & " , '0', " & GlTipoCambioOficial & ", '100', '0', '2', 'REG', '" & Date & "', '" & obs_fo_gastos_detalle & "', " & VAR_CMPBTE & ", '" & glusuario & "', '" & Date & "', 'REG', '4.2.3')"
                   VAR_REG = VAR_REG + 1
               '     'cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino,
               '     'Fecha_Aprobacion_tesoreria, fecha_impresion_cheque, banco_destino,
               Wend
            End If
            'If rstdestino.State = 1 Then rstdestino.Close
 
        End If
        'db.Execute "update ao_compra_planilla_pagos set estado_codigo = 'APR' where compra_codigo = " & Ado_detalle2.Recordset!compra_codigo & " and adjudica_codigo = " & Ado_detalle2.Recordset!adjudica_codigo & " and pago_codigo=" & Ado_detalle3.Recordset!pago_codigo & "   "
        'Call ABRIR_TABLA_DET
    Else
        MsgBox "NO se puede APROBAR un registro Anulado o previamente Aprobado. ", vbExclamation, "Atenci�n!"
    End If
        'WWWWWWWWWW
End Sub
Private Sub ABRIR_TABLA_DET3()
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    'rs_det2.Open "SELECT ges_gestion, bitacora_codigo, estado_codigo, fecha_registro, hora_registro, usr_codigo, unidad_codigo, solicitud_codigo as codigo2, negocia_forma  as codigo3, beneficiario_codigo  as codigo4, beneficiario_codigo_cgi  as codigo5, negocia_tarea_realizada as descripcion, negocia_observaciones as campo1, negocia_fecha_prevista As fecha1, negocia_fecha_real As fecha2, negocia_hora_prevista As campo2, negocia_hora_real As campo3, negocia_gasto_estimado As monto1, bitacora_cite, beneficiario_nombre_ref From ao_negociacion_bitacora WHERE unidad_codigo = '" & parametro & "' and solicitud_codigo = " & GlSolicitud & " ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det2.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle2.Recordset = rs_det2
    Set dg_det1.DataSource = Ado_detalle2.Recordset
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo <> "ERR" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    Call ABRIR_TABLA_DET 'ABRIR_TABLA_DET3
    aw_p_ao_bitacora.txt_codigo.Caption = Me.txt_codigo.Caption
    aw_p_ao_bitacora.txt_campo1.Caption = Me.dtc_codigo1.Text
    aw_p_ao_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    aw_p_ao_bitacora.Txt_Correl.Caption = 0
    aw_p_ao_bitacora.Txt_estado.Caption = "REG"
    aw_p_ao_bitacora.txt_cliente.Text = txt_obs
    Ado_detalle1.Recordset.AddNew
    aw_p_ao_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET 'ABRIR_TABLA_DET3
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
    Ado_datos.Recordset.Move marca1 - 1
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este fue Anulado!! ", vbExclamation
  End If
   

'----------------
'  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos!estado_cotiza = "REG" Then
'    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'    swnuevo = 1
'    fraOpciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet1.Enabled = False
'    FrmABMDet.Enabled = False
'    FraDet2.Enabled = False
'    FrmABMDet2.Enabled = False
'    Fra_datos.Enabled = False
'    Call ABRIR_TABLA_DET
'    frm_ao_solicitud_bitacora.txt_codigo.Caption = Me.txt_codigo.Caption
'    frm_ao_solicitud_bitacora.txt_campo1.Caption = Me.dtc_codigo1.Text
'    frm_ao_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
'    frm_ao_solicitud_bitacora.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
'    frm_ao_solicitud_bitacora.Txt_estado.Caption = "REG"
'    frm_ao_solicitud_bitacora.lbl_bitacora.Caption = Me.FraDet1.Caption
'    Ado_detalle1.Recordset.AddNew
'    frm_ao_solicitud_bitacora.Show vbModal
'
'    Call ABRIR_TABLA_DET
'
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraNavega.Enabled = True
'    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Enabled = True
'    'Fra_datos.Enabled = True
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya est� Aprobado!! ", vbExclamation
'  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            VAR_DET = "21000"
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes7A.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes7A.txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes7A.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes7A.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes7A.lbl_det.Caption = VAR_DET     '"34110"
            frm_solicitud_bienes7A.Txt_estado.Caption = "REG"
            frm_solicitud_bienes7A.Show vbModal
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    'COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Call ABRIR_TABLA_DET
            Ado_detalle1.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            aw_p_ao_solicitud_edificacion.Txt_Correl.Caption = 0
'            aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
'            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
'            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        Case "5"    ' SERVICIO MODERNIZACION
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya est� Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle2.Recordset.RecordCount > 0 Then
   sino = MsgBox("Est� Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atenci�n")
   If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle2.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado ...", vbExclamation, "Validaci�n de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro no fue identificado correctamente ...", vbExclamation, "Validaci�n de Registro"
 End If

'--------------------
'  If Ado_detalle1.Recordset.RecordCount > 0 Then
'   sino = MsgBox("Est� Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atenci�n")
'   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
'      If sino = vbYes Then
'        Ado_detalle1.Recordset.Delete 'adAffectAll
''        Ado_detalle1.Recordset("estado_codigo") = "ERR"
''        Ado_detalle1.Recordset("fecha_registro") = Date
''        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
''        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
''        Ado_detalle1.Recordset.Update  'Batch adAffectAll
'      End If
'   Else
'        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado ...", vbExclamation, "Validaci�n de Registro"
'   End If
' Else
'     MsgBox "No se puede ANULAR, el registro no fue identificado correctamente ...", vbExclamation, "Validaci�n de Registro"
' End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   sino = MsgBox("Est� Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atenci�n")
   
   'If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
     If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle2.Recordset.Delete 'adAffectAll
'        Ado_detalle1.Recordset("estado_codigo") = "ERR"
'        Ado_detalle1.Recordset("fecha_registro") = Date
'        Ado_detalle1.Recordset("usr_codigo") = GlUsuario
'        Ado_detalle1.Recordset("campo1") = "REG. ANULADO"
'        Ado_detalle1.Recordset.Update  'Batch adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validaci�n de Registro"
   End If
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
        MsgBox "No se puede APROBAR, debe registrar: " + lbl_campo4.Caption, vbExclamation, "Validaci�n de Registro"
        Exit Sub
   End If
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
   If rs_datos!estado_cotiza = "REG" Then
      VAR_COD4 = Ado_datos.Recordset!unidad_codigo
      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
      VAR_PROY2 = Ado_datos.Recordset!edif_codigo
      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
      sino = MsgBox("Est� Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atenci�n")
      If sino = vbYes Then
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
                'PAGOS
                'WWWWWWWWWWWWWW
                    ' ini generaci�n de correlativo
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "select * from fc_organismo_financiamiento where org_codigo = '111'  ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux2.RecordCount > 0 Then
                       rs_aux2!correlativo_gasto = rs_aux2!correlativo_gasto + 1
                       VAR_CMPBTE = rs_aux2!correlativo_gasto
                       rs_aux2.Update
                    End If
                    'WWWWWWWWWWWWWWW
                    'correlv = Ado_datos.Recordset!venta_codigo
                    'VAR_TIPOV = Ado_datos.Recordset!venta_tipo
           
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    rs_aux3.Open "select * from  fo_gastos_cabecera where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & " ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux3.RecordCount = 0 Then
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion     'Year(Date)
                        rs_aux3!org_codigo = "111"
                        rs_aux3!gasto_codigo = VAR_CMPBTE
                        rs_aux3!tipo_comp = CYD
                        rs_aux3!gasto_codigo_anterior = VAR_CMPBTE
                        rs_aux3!unidad_codigo = VAR_COD4
                        rs_aux3!solicitud_codigo = VAR_SOL
                        rs_aux3!tipo_formulario = "F02"
                        rs_aux3!pago_codigo = rs_aux3.RecordCount + 1
                        
                        rs_aux3!proceso_codigo = "FIN"
                        rs_aux3!subproceso_codigo = "FIN-03"
                        rs_aux3!etapa_codigo = "FIN-03-03"
                        rs_aux3!clasif_codigo = "ADM"
                        rs_aux3!doc_codigo = "R-111"
                        rs_aux3!doc_numero = 0
                        rs_aux3!poa_codigo = "4.2.3"
                        
                        rs_aux3!fecha_egreso = Date
                        rs_aux3!tipo_moneda = "BOB"
                        rs_aux3!da_codigo = "1.1"
                        
                        rs_aux3!fte_codigo = "10"
                        rs_aux3!monto_Bolivianos = 0
                        rs_aux3!monto_dolares = 0
                        rs_aux3!liquido_pagar = 0
                        rs_aux3!monto_Bolivianos_pag = 0
                        rs_aux3!monto_dolares_pag = 0
                        rs_aux3!Deducciones = 0
                        rs_aux3!fecha_autorizacion = Date
                        rs_aux3!justificacion = Txt_descripcion.Text
                        rs_aux3!es_base = "S"
                        
                        rs_aux3!CODIGO_GRUPO = rs_aux3!pago_codigo
                        rs_aux3!NUMERO_PAGO = 1
                        rs_aux3!observaciones = txt_obs.Text
                        
                        'rs_aux3!edif_codigo = VAR_PROY2
                        'rs_aux3!beneficiario_codigo = VAR_BENEF
                        'rs_aux3!solicitud_tipo = "10"
                        'rs_aux3!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant   'VAR_CITE
                        
                        rs_aux3!estado_devengado = "APR"
                        rs_aux3!estado_pagado = "REG"
                        rs_aux3!estado_contabilidad = "REG"
                        
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3!fecha_registro = Date
                        rs_aux3!usr_codigo_aprueba = glusuario
                        rs_aux3!fecha_aprueba = Date
                        rs_aux3.Update
                        
                        'DETALLE Carga fo_gastos_detalle
                        Set rstdestino = New ADODB.Recordset
                        If rstdestino.State = 1 Then rstdestino.Close
                        rstdestino.Open "select * from fo_gastos_detalle where org_codigo = '111' AND gasto_codigo= " & VAR_CMPBTE & "  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rstdestino.RecordCount > 0 Then
                        End If
                        Set rs_aux4 = New ADODB.Recordset
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        rs_aux4.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & VAR_COD4 & "' AND solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux4.RecordCount > 0 Then
                           VAR_REG = 1
                           rs_aux4.MoveFirst
                           While Not rs_aux4.EOF
                           '     db.Execute "INSERT INTO ao_compra_detalle (ges_gestion, compra_codigo, compra_codigo_det, bien_codigo, compra_cantidad, compra_precio_unitario_bs, compra_descuento_bs, compra_precio_total_bs, compra_precio_unitario_dol, compra_descuento_dol, compra_precio_total_dol, compra_concepto, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo , usr_usuario, fecha_registro) " & _
                           '     "VALUES ('" & glGestion & "', " & rs_aux3!compra_codigo & ", " & VAR_REG & ", '" & rs_aux4!bien_codigo & "', " & rs_aux4!bien_cantidad & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", " & rs_aux4!bien_precio_venta_base & ", '0', " & rs_aux4!bien_total_venta & ", '" & rs_aux3!compra_descripcion & "', '" & rs_aux4!grupo_codigo & "', '" & rs_aux4!subgrupo_codigo & "', '" & rs_aux4!par_codigo & "', '1', '0', '" & glusuario & "', '" & Date & "')"
                           '     rs_aux4.MoveNext
                                
                                db.Execute "INSERT INTO fo_gastos_detalle (ges_gestion, org_codigo, gasto_codigo, gasto_codigo_detalle, par_codigo, pro_codigo, codigo_beneficiario, concepto_pago, monto_total, monto_dolares_dev, tipo_cambio_dev, monto_Bolivianos, monto_Dolares, saldo_bolivianos, tipo_cambio, Porcentaje, deducciones, fecha_pago, depto_codigo, estado_aprobacion, fecha_autorizacion, Observacion, codigo_dev, usr_usuario, fecha_registro, hora_registro,  estado_conciliacion, codigo_poa " & _
                                "VALUES ('" & glGestion & "', '111', " & VAR_CMPBTE & ", '" & VAR_REG & "', " & rs_aux4!par_codigo & ", '8', '" & VAR_BENEF & "', '" & Txt_descripcion.Text & "', " & rs_aux4!bien_total_compra & ", " & rs_aux4!bien_total_compra * GlTipoCambioOficial & ", " & GlTipoCambioOficial & ", " & rs_aux4!bien_total_compra & ", " & rs_aux4!bien_total_compra * GlTipoCambioOficial & " , '0', " & GlTipoCambioOficial & ", '100', '0', '2', 'REG', '" & Date & "', '" & txt_obs.Text & "', " & VAR_CMPBTE & ", '" & glusuario & "', '" & Date & "', 'REG', '4.2.3')"
                               VAR_REG = VAR_REG + 1
                           '     'cta_codigo, cheque_o_trf, numero_cheque_trf, cta_codigo_destino, cheque_o_trf_destino, numero_cheque_trf_destino,
                           '     'Fecha_Aprobacion_tesoreria, fecha_impresion_cheque, banco_destino,
                           Wend
                        End If
                        If rstdestino.State = 1 Then rstdestino.Close
                    End If
                    
                    'WWWWWWWWWW
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            Case "5"    ' SERVICIO MODERNIZACION
        End Select
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!estado_codigo = "APR"
'        rs_datos!fecha_registro = Date
'        rs_datos!usr_codigo = glusuario
'        rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validaci�n de Registro"
   End If
  Else
      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexi�n = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Est� Seguro de CANCELAR la operaci�n ? ", vbYesNo + vbQuestion, "Atenci�n")
   If sino = vbYes Then
        rs_datos.CancelUpdate
'        If mvBookMark > 0 Then
'          rs_datos.BookMark = mvBookMark
'        Else
'          rs_datos.MoveFirst
'        End If
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            Call OptFilGral1_Click
        Else
            Call OptFilGral2_Click
        End If
        rs_datos.MoveFirst
        mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        'txt_codigo.Enabled = True
        VAR_SW = ""
'        dtc_codigo9.Enabled = True
    End If
'    dtc_desc1.Visible = True
'    lbl_aux1.Visible = False
End Sub

Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    If ExisteReg(Ado_datos.Recordset!unidad_codigo, Ado_datos.Recordset!solicitud_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atenci�n": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Est� Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atenci�n")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ERR"
          rs_datos!fecha_registro = Date
          rs_datos!usr_codigo = glusuario
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validaci�n de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
  End If
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnDesAprobar_Click()
  On Error GoTo UpdateErr
   sino = MsgBox("Est� Seguro de DESAPROBAR el Registro ? ", vbYesNo + vbQuestion, "Atenci�n")
   If rs_datos!estado_codigo = "APR" Then
      If sino = vbYes Then
         rs_datos!estado_codigo = "REG"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
        MsgBox "No se puede DESAPROBAR un registro Elaborado o Errado ...", vbExclamation, "Validaci�n de Registro"
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
    VAR_UNI = dtc_codigo1.Text
    var_cod = IIf(txt_codigo.Caption = "", rs_datos!solicitud_codigo, txt_codigo.Caption)
    If VAR_SW = "ADD" Then
        'VAR_UNI = dtc_codigo1.Text
        'var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' and solicitud_codigo = " & var_cod & "  "
        SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1.RecordCount + 1
            'MsgBox "El c�digo ya existe, consulte con el administrador del Sistema..."
            'var_cod = 0
            'Exit Sub
        Else
            'var_cod = rs_datos.RecordCount '+ 1
            var_cod = 1
        End If
        'var_cod = RTrim(RTrim(dtc_codigo2.Text) + "-") + LTrim(Str(Val(dtc_aux2) + 1))
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' Year(Date)   'no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        'rs_datos!correl_edificacion = 0
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
     End If
     'var_cod = rs_datos!solicitud_codigo
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = IIf(dtc_codigo3.Text = "", "20101-0", dtc_codigo3.Text)
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     
     Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "FIN"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "FIN-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "FIN-03-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "ADM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-XXX"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
  
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            If VAR_UNI = "DNINS" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNAJS" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNMAN" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNREP" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNEME" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
            If VAR_UNI = "DNMOD" Then
                rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            End If
        Case "5"    ' SERVICIO MODERNIZACION
        Case Else
            Select Case parametro
                Case "UALMI"    'INSUMOS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                Case "UALMR"    'REPUESTOS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
                Case "UALMH"    'HERRAMIENTAS
                    rs_datos!etapa_codigo = "TEC-06-01"
                    rs_datos!doc_codigo = "R-126"
            End Select
     End Select
     rs_datos!poa_codigo = dtc_codigo10.Text
     rs_datos!solicitud_observaciones = txt_obs.Text
     rs_datos!solicitud_fecha_recepci�n = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text
     
     rs_datos!ges_gestion_ant = glGestion       'Year(Date)
     If var_cod < 10 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00000" + Trim(txt_codigo)
     End If
     If var_cod > 9 And var_cod < 100 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0000" + Trim(txt_codigo)
     End If
     If var_cod > 99 And var_cod < 1000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-000" + Trim(txt_codigo)
     End If
     If var_cod > 999 And var_cod < 10000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-00" + Trim(txt_codigo)
     End If
     If var_cod > 9999 And var_cod < 100000 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-0" + Trim(txt_codigo)
     End If
     If var_cod > 99999 Then
        rs_datos!unidad_codigo_ant = VAR_UNI + "-" + Trim(txt_codigo)
     End If
     rs_datos!solicitud_codigo_ant = 0
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
     Else
        Call OptFilGral2_Click
     End If
     rs_datos.MoveLast
     mbDataChanged = False
      
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
'     dtc_desc1.BackColor = &HFFFFC0
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
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo8.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo10.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validaci�n de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
'    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\contabilidad\fr_listado_solicitud_servicios_basicos.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR01.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR01.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!depto_codigo
        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!subgrupo_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresi�n"
        CR01.WindowState = crptMaximized
'    Else
'''        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atenci�n"
'    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atenci�n"
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR01.ReportFileName = App.Path & "\Reportes\Contabilidad\ar_bitacora_negociaciones.rpt"
        CR01.WindowShowPrintSetupBtn = True
        CR01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR01.PrintReport
        If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresi�n"
        CR01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atenci�n"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atenci�n"
  End If
End Sub

Private Sub BtnModDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And Ado_detalle1.Recordset!estado_codigo = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
        
    aw_p_ao_bitacora.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'solicitud_codigo
    aw_p_ao_bitacora.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
    aw_p_ao_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
    aw_p_ao_bitacora.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
    'aw_p_ao_negociacion_bitacora.Txt_estado.Caption = "REG"
    'Ado_detalle1.Recordset.AddNew
     
   ' aw_p_ao_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
    ' aw_p_ao_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset!negocia_forma
     aw_p_ao_bitacora.dtc_desc1.BoundText = Me.Ado_detalle1.Recordset!negocia_forma
     aw_p_ao_bitacora.dtc_desc3.BoundText = IIf(IsNull(Me.Ado_detalle1.Recordset!beneficiario_codigo_cgi), "0", Me.Ado_detalle1.Recordset!beneficiario_codigo_cgi)
     
    aw_p_ao_bitacora.DTPfecha1.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_fecha_real), Date, Me.Ado_detalle1.Recordset!negocia_fecha_real)        'Fecha
    aw_p_ao_bitacora.Txt_campo2.Value = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_real) Or (Me.Ado_detalle1.Recordset!negocia_hora_real = "0"), Str(Time), Me.Ado_detalle1.Recordset!negocia_hora_real) 'Hora
    aw_p_ao_bitacora.Txt_monto1.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_gasto_estimado), 0, Me.Ado_detalle1.Recordset!negocia_gasto_estimado)
    aw_p_ao_bitacora.dtc_codigo2.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!beneficiario_codigo), "0", Me.Ado_detalle1.Recordset!beneficiario_codigo)
    aw_p_ao_bitacora.dtc_codigo3.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!beneficiario_codigo_cgi), "0", Me.Ado_detalle1.Recordset!beneficiario_codigo_cgi)
    aw_p_ao_bitacora.Txt_campo3.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_tarea_realizada), "NINGUNA", Me.Ado_detalle1.Recordset!negocia_tarea_realizada)
    aw_p_ao_bitacora.Txt_campo4.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_observaciones), "", Me.Ado_detalle1.Recordset!negocia_observaciones)
    aw_p_ao_bitacora.Txt_campo5.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!bitacora_cite), "-", Me.Ado_detalle1.Recordset!bitacora_cite)
    If swnuevo = 2 Then
        'aw_p_ao_bitacora.dtc_desc1.BoundText = aw_p_ao_bitacora.dtc_codigo1.BoundText
        aw_p_ao_bitacora.dtc_desc2.BoundText = aw_p_ao_bitacora.dtc_codigo2.BoundText
       ' aw_p_ao_bitacora.dtc_desc3.BoundText = aw_p_ao_bitacora.dtc_codigo3.BoundText
    End If
    
    aw_p_ao_bitacora.Show vbModal
    
    Call ABRIR_TABLA_DET 'ABRIR_TABLA_DET3
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar un registro Aprobado o verifique si fue correctamente identificado !! ", vbExclamation
  End If

'-------------
'  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos.RecordCount > 0 And rs_datos!estado_cotiza = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
'    swnuevo = 2
'    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
'    fraOpciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet1.Enabled = False
'    FrmABMDet.Enabled = False
'    FraDet2.Enabled = False
'    FrmABMDet2.Enabled = False
'    Fra_datos.Enabled = False
'
'    frm_ao_solicitud_bitacora.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'    frm_ao_solicitud_bitacora.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'    frm_ao_solicitud_bitacora.Txt_descripcion.Caption = Me.dtc_desc1.Text
'    frm_ao_solicitud_bitacora.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'    'frm_ao_solicitud_bitacora.Txt_estado.Caption = "REG"
'    'Ado_detalle1.Recordset.AddNew
'
'    frm_ao_solicitud_bitacora.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
'    frm_ao_solicitud_bitacora.dtpFecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
'    frm_ao_solicitud_bitacora.Txt_campo2.Value = Me.Ado_detalle1.Recordset("negocia_hora_real")
'    frm_ao_solicitud_bitacora.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
'    frm_ao_solicitud_bitacora.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
'    frm_ao_solicitud_bitacora.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
'    frm_ao_solicitud_bitacora.Txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
'    frm_ao_solicitud_bitacora.Txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
'    frm_ao_solicitud_bitacora.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
'    If swnuevo = 2 Then
'        frm_ao_solicitud_bitacora.dtc_desc1.BoundText = frm_ao_solicitud_bitacora.dtc_codigo1.BoundText
'        frm_ao_solicitud_bitacora.dtc_desc2.BoundText = frm_ao_solicitud_bitacora.dtc_codigo2.BoundText
'        frm_ao_solicitud_bitacora.dtc_desc3.BoundText = frm_ao_solicitud_bitacora.dtc_codigo3.BoundText
''        frm_ao_solicitud_bitacora.HH = Left(frm_ao_solicitud_bitacora.Txt_campo2.Value, 2)
''        frm_ao_solicitud_bitacora.MM = Right(frm_ao_solicitud_bitacora.Txt_campo2.Text, 2)
'    End If
'
'    frm_ao_solicitud_bitacora.Show vbModal
'
'    Call ABRIR_TABLA_DET
'
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraNavega.Enabled = True
'    FraDet1.Enabled = True
'    FrmABMDet.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Enabled = True
'    'Fra_datos.Enabled = True
'    Call OptFilGral1_Click
'  Else
'    MsgBox "No se puede Modificar el registro, verifique si est� Aprobado o fue correctamente identificado !! ", vbExclamation
'  End If

End Sub

Private Sub BtnModDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    VAR_DET = "21000"
    Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            If VAR_DET = "21000" Then
                frm_solicitud_bienes7A.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes7A.txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes7A.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                frm_solicitud_bienes7A.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes7A.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes7A.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes7A.dtc_desc1.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.dtc_aux1.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.Dtc_aux2.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.dtc_aux3.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.Txt_campo2.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.Txt_campo3.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.Txt_campo4.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                frm_solicitud_bienes7A.Txt_campo18.BoundText = frm_solicitud_bienes7A.dtc_codigo1.BoundText
                
                frm_solicitud_bienes7A.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion), "-", Me.Ado_detalle2.Recordset!bien_descripcion)
                frm_solicitud_bienes7A.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle2.Recordset!bien_descripcion_anterior)
                frm_solicitud_bienes7A.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                frm_solicitud_bienes7A.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                frm_solicitud_bienes7A.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                frm_solicitud_bienes7A.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_compra")
                frm_solicitud_bienes7A.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_compra")
                
                frm_solicitud_bienes7A.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                frm_solicitud_bienes7A.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                frm_solicitud_bienes7A.dtc_desc2.BoundText = frm_solicitud_bienes7A.dtc_codigo2.BoundText
                
                frm_solicitud_bienes7A.Txt_campo15.Text = Me.Ado_detalle2.Recordset("fosa_dimension_frente")
                
                
                
                frm_solicitud_bienes7A.lbl_det.Caption = VAR_DET
                frm_solicitud_bienes7A.Show vbModal
            End If
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            Call ABRIR_TABLA_DET
            aw_p_ao_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
            aw_p_ao_solicitud_edificacion.txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
            aw_p_ao_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
            'aw_p_ao_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
            'aw_p_ao_solicitud_edificacion.Txt_estado.Caption = "REG"
            aw_p_ao_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
            aw_p_ao_solicitud_edificacion.dtc_desc1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux1.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.Dtc_aux2.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            aw_p_ao_solicitud_edificacion.dtc_aux3.BoundText = aw_p_ao_solicitud_edificacion.dtc_codigo1.BoundText
            
            aw_p_ao_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
            aw_p_ao_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
            aw_p_ao_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
            aw_p_ao_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
            aw_p_ao_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
            aw_p_ao_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
            aw_p_ao_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
            aw_p_ao_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
            aw_p_ao_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
            aw_p_ao_solicitud_edificacion.Txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
            aw_p_ao_solicitud_edificacion.Txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
            aw_p_ao_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
        
            aw_p_ao_solicitud_edificacion.Show vbModal
        Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
        Case "5"    ' SERVICIO MODERNIZACION
           
    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
'    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Modificar el registro, porque este ya est� Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_cotiza = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
        DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc4.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validaci�n de Registro"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atenci�n!"
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

Private Sub BtnVer_Click()
  On Error GoTo QError
  If rs_datos!estado_codigo = "APR" Then
    Dim ARCH_FOTO As String
    Dim SW0 As String
    Select Case Left(Trim(Ado_datos.Recordset("edif_codigo")), 1)
        Case "1"    'CHQ
            VAR_DPTO = "CHQ"
        Case "2"    'LPZ
            VAR_DPTO = "LPZ"
        Case "3"    'CBB
            VAR_DPTO = "CBB"
        Case "4"    'SCZ
            VAR_DPTO = "SCZ"
        Case "5"    'PTS
            VAR_DPTO = "PTS"
        Case "6"    'ORU
            VAR_DPTO = "ORU"
        Case "7"    'TJA
            VAR_DPTO = "TJA"
        Case "8"    'BEN
            VAR_DPTO = "BEN"
        Case "9"    'PDO
            VAR_DPTO = "PDO"
    End Select
    If Ado_datos.Recordset!archivo_respaldo_cargado = "N" Then
      'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
      NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "DED2"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      'MsgBox ""
      'negocia_codigo, unidad_codigo, negocia_fecha_inicio as fecha1, negocia_descripcion, estado_codigo, fecha_registro, usr_codigo, solicitud_tipo as codigo2, edif_codigo as codigo3, beneficiario_codigo as codigo4, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, hora_registro, ges_gestion, archivo_respaldo, archivo_respaldo_cargado
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atenci�n")
      If sino = vbYes Then
          'NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(Ado_datos.Recordset!edif_tipo) & "\" & Trim(Ado_datos.Recordset!negocia_codigo) & "\"
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        'e = ShellExecute(0, vbNullString, App.Path & "\" & Trim(GLCarpeta2) & "\" & Trim(TxtInicial.Text) & "-" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!codigo_beneficiario) & "\LICENCIAS\" & Trim(frmBeneficiario_Control.AdoPermiso.Recordset!ARCHIVO), vbNullString, vbNullString, vbNormalFocus)
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    '    If SW0 = 1 Then
    '    '    If GlServidor = "SRVPRO" Then
    '    '        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
    '    '    Else
    '            'ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo)
    '            ARCH_FOTO = App.Path + "\BIENES\EDIFICIOS\" + Trim(Ado_datos.Recordset!edif_tipo) + "\" + Trim(Ado_datos.Recordset!edif_codigo) + ".JPG"
    '    '    End If
    '        'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    '        CodBien = Ado_datos.Recordset!edif_codigo
    '        If Guardar_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
    '            MsgBox "Se cargo la Imagen Correctamente !!"
    '        Else
    '            MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    '        End If
    '    Else
    '        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    '        Image2 = Img_Foto
    '    End If
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validaci�n de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atenci�n"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText

End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

'Private Sub dtc_codigo5_Click(Area As Integer)
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'End Sub

'Private Sub dtc_codigo6_Click(Area As Integer)
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'End Sub

'Private Sub dtc_codigo7_Click(Area As Integer)
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'End Sub

'Private Sub dtc_codigo8_Click(Area As Integer)
'    dtc_desc8.BoundText = dtc_codigo8.BoundText
'End Sub

'Private Sub dtc_codigo9_Click(Area As Integer)
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
'End Sub

'Private Sub dtc_codigo9_LostFocus()
''  If VAR_SW = "ADD" Then
''    Set rs_aux2 = New ADODB.Recordset
''    SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9.Text & "'  "
''    rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
''    If rs_aux2.RecordCount > 0 Then
''        rs_aux2!correl_doc = rs_aux2!correl_doc + 1
''        txt_campo1.Caption = rs_aux2!correl_doc
''        rs_aux2.Update
''    End If
''  End If
'  txt_aux9.Text = dtc_desc9.Text
'End Sub

'Private Sub dtc_desc5_Click(Area As Integer)
'    dtc_codigo5.BoundText = dtc_desc5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub
   
'Private Sub pnivel5(codigo5 As String)
'   'Dim strConsultaF As String
'   'strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo5 & "'"
'
'   Set dtc_codigo6.RowSource = Nothing
'   'Set dtc_codigo6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_codigo6.ReFill
'   dtc_codigo6.BoundText = Empty
'
'   Set dtc_desc6.RowSource = Nothing
'   'Set dtc_desc6.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_desc6.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_proceso_nivel2 '" & codigo5 & "' ")
'   dtc_desc6.ReFill
'   dtc_desc6.BoundText = Empty
'End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    dtc_aux1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub
   
Private Sub pnivel1(codigo1 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from pc_poa_actividad where unidad_codigo = '" & codigo1 & "'"
   
   Set dtc_codigo10.RowSource = Nothing
'   Set dtc_codigo10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo10.ReFill
   dtc_codigo10.BoundText = Empty
   
   Set dtc_desc10.RowSource = Nothing
   'Set dtc_desc10.RowSource = db.Execute(strConsultaF, , adCmdText)
   Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc10.ReFill
   dtc_desc10.BoundText = Empty
End Sub
  
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
'    Call pnivel5(dtc_codigo5.BoundText)
'    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText

End Sub
 
Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc4_LostFocus()
    dtc_codigo4.BoundText = dtc_desc4.BoundText
    If Txt_descripcion.Text = "" Then
        Txt_descripcion.Text = lbl_titulo.Caption + " - " + dtc_desc4
    End If
    'Call pnivel1(dtc_codigo4.BoundText)
    Call pnivel1(parametro)
    dtc_desc10.Enabled = True
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    parametro = Aux
    'parametro = "DVTA"
    '
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    'txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    'JQA 2014-JUL-14
    'db.Execute (" EXEC gp_actualiza_beneficiario_edif ")
'    lbl_aux1.Visible = False
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    'rs_datos2.Open "Select * from gc_tipo_solicitud order by solicitud_tipo", db, adOpenStatic
    rs_datos2.Open "gp_listar_apr_gc_tipo_solicitud", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from fo_proyectos_ejecucion order by pro_codigo_det_descripcion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from gc_beneficiario where tipoben_codigo = '22' order by beneficiario_denominacion", db, adOpenStatic
    'rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'pc_poa_actividad
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    'rs_datos10.Open "Select * from pc_poa_actividad order by poa_codigo", db, adOpenStatic
    rs_datos10.Open "pp_listar_apr_pc_poa_actividad", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    'gc_beneficiario (Personal CGI)
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

'Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepci�n as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
'    queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub ABRIR_TABLA_DET()
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    Dim sqlBita As String
    sqlBita = "select * from ao_solicitud_bitacora where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
    rs_det1.Open sqlBita, db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    
    Set rs_aux2 = New ADODB.Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    'rs_aux2.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Select Case parametro
        Case "UALMI"    'INSUMOS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39800' and par_codigo <> '34800'))  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "UALMR"    'REPUESTOS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '39800' )  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '39810' or par_codigo = '39820')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "UALMH"    'HERRAMIENTAS
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Case "DCONT"    'Servicios Basicos
            rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    ", db, adOpenKeyset, adLockOptimistic, adCmdText
            'rs_aux2.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    ", db, adOpenKeyset, adLockOptimistic, adCmdText
            'and (subgrupo_codigo = '21000' or subgrupo_codigo = '22000')
            
    End Select
    Set Ado_detalle2.Recordset = rs_aux2
    Set dg_det2.DataSource = Ado_detalle2.Recordset
    
End Sub

Private Sub ABRIR_TABLA_AUX2()
    Set rs_datos11 = New ADODB.Recordset
    If rs_datos11.State = 1 Then rs_datos11.Close
    'rs_datos11.Open "Select * from gv_personal_contratado where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' order by beneficiario_denominacion", db, adOpenKeyset, adLockOptimistic, adCmdText   ', adOpenStatic
    rs_datos11.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos11.Recordset = rs_datos11
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  lblStatus.Width = Me.Width - 1500
  cmdNext.Left = lblStatus.Width + 700
  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrar� la posici�n de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificaci�n del Cliente                Fin -->   'esto es de Caption
    'Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
    'Image2 = Img_Foto
'    If Ado_datos.Recordset!archivo_foto_cargado = "S" Then
'        'BtnVer.Visible = True
'        Set Img_Foto = Leer_Imagen(db, "Select Foto From gc_edificaciones Where edif_codigo = '" & Ado_datos.Recordset("edif_codigo") & "' ", "Foto")
'        Image2 = Img_Foto
'    Else
'        'BtnVer.Visible = False
'        'chkEstado.Value = vbUnchecked
'    End If
    If VAR_SW <> "ADD" Then
'        Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'                Call ABRIR_TABLA_DET
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'                Call ABRIR_TABLA_DET
'            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'                Call ABRIR_TABLA_DET
'            Case "5"    ' SERVICIO MODERNIZACION
'
'            Case Else
'                Call ABRIR_TABLA_DET
'        End Select
        DTPfecha1.Value = Ado_datos.Recordset!solicitud_fecha_solicitud
        Call ABRIR_TABLA_DET
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    FraDet1.Caption = "BIT�CORA DE: " + dtc_desc1.Text
'    txt_aux9.Text = dtc_desc9.Text
    If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet2.Visible = False
    Else
            FrmABMDet2.Visible = True
    End If
  End If
End Sub

Private Sub Ado_datos_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Aqu� se coloca el c�digo de validaci�n
  'Se llama a este evento cuando ocurre la siguiente acci�n
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

Private Sub BtnA�adir_Click()
  On Error GoTo AddErr
    VAR_SW = "ADD"
    'lblStatus.Caption = "Agregar registro"
    Fra_datos.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    'txt_codigo.Enabled = False
'    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
'    rs_datos.AddNew
    Ado_datos.Recordset.AddNew
    dtc_desc11.SetFocus
    'dtc_desc1.BackColor = &H80000005
    dtc_codigo1.Text = parametro
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
    dtc_desc2.Locked = True
    Select Case parametro
        Case "DVTA"        'INI COMERCIAL
            dtc_codigo2.Text = 3
        Case "COMEX"        'INI COMEX
            dtc_codigo2.Text = 3
        Case "DNINS"                        'INI GRABA INSTALACIONES
            '
            dtc_codigo2.Text = 4
        Case "DNAJS"
            '
            dtc_codigo2.Text = 4
        Case "DNMAN"
            'DCONT
            dtc_codigo2.Text = 4
        Case "DCONT"
            dtc_codigo2.Text = 1
        Case Else
            dtc_codigo2.Text = 5
    End Select
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    DTPfecha1.Value = Date
'    dtc_codigo5.Text = "COM"
'    dtc_desc5.BoundText = dtc_codigo5.BoundText
'    dtc_codigo6.Text = "COM-01"
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'    dtc_codigo7.Text = "COM-01-02"
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'    BtnVer.Visible = False
'    dtc_codigo9.Enabled = False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Esto s�lo es necesario en aplicaciones multiusuario
  On Error GoTo RefreshErr
  rs_datos.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Function ExisteReg(Unidad As String, solicitud As Integer) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_compra_cabecera WHERE unidad_codigo = '" & Unidad & "' and solicitud_codigo=" & solicitud & "  "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from av_solicitud_insumos where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
        Case "UALMR"    'REPUESTOS
            queryinicial = "Select * from av_solicitud_repuestos where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas where estado_cotiza = 'REG' "     'AND unidad_codigo = '" & parametro & "'
        Case "DCONT"    'Servicios Basicos
            queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'"
            'queryinicial = "Select * from av_solicitud_servicio_basico where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "'"
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case parametro
        Case "UALMI"    'INSUMOS
            queryinicial = "Select * from av_solicitud_insumos "
        Case "UALMR"    'REPUESTOS
            queryinicial = "Select * from av_solicitud_repuestos "
        Case "UALMH"    'HERRAMIENTAS
            queryinicial = "Select * from av_solicitud_herramientas "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

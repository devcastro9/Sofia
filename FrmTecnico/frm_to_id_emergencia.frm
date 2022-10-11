VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_to_id_emergencia 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Administrativos - Area Técnica - Identificación de Emergencias"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frm_to_id_emergencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.CommandButton BtnImprimir7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Servicios"
      Height          =   640
      Left            =   990
      Picture         =   "frm_to_id_emergencia.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "Cotizacion y Costos del Servicio"
      Top             =   9000
      Width           =   765
   End
   Begin VB.Frame FraDet7 
      BackColor       =   &H00000000&
      Caption         =   "SERVICIO TECNICO EXTERNO"
      ForeColor       =   &H00FFFFC0&
      Height          =   1520
      Left            =   1880
      TabIndex        =   103
      Top             =   8160
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det7 
         Bindings        =   "frm_to_id_emergencia.frx":2184
         Height          =   1215
         Left            =   60
         TabIndex        =   104
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12632319
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
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "bien_total_venta"
            Caption         =   "Precio X Servicio"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion Servicio"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Equipo"
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
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   120
      Picture         =   "frm_to_id_emergencia.frx":219F
      ScaleHeight     =   1380
      ScaleWidth      =   1635
      TabIndex        =   99
      Top             =   8240
      Width           =   1695
      Begin VB.CommandButton BtnAnlDetalle7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":6E1D1
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":6E613
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":6EA55
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Adiciona Detalle"
         Top             =   40
         Width           =   765
      End
   End
   Begin VB.CommandButton BtnImprimir1 
      BackColor       =   &H80000018&
      Caption         =   "Bitácora"
      Height          =   640
      Left            =   14790
      Picture         =   "frm_to_id_emergencia.frx":6EE97
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "Bitacora de Eventos"
      Top             =   5945
      Width           =   765
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   13920
      Picture         =   "frm_to_id_emergencia.frx":70619
      ScaleHeight     =   1380
      ScaleWidth      =   1635
      TabIndex        =   91
      Top             =   8240
      Width           =   1695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Herram."
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":DAA37
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Cotizacion y Costos del Servicio"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":DC1B9
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":DC5FB
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":DCA3D
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Adiciona Producto"
         Top             =   40
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1430
      Left            =   13920
      Picture         =   "frm_to_id_emergencia.frx":DCE7F
      ScaleHeight     =   1365
      ScaleWidth      =   1635
      TabIndex        =   87
      Top             =   6710
      Width           =   1695
      Begin VB.CommandButton BtnImprimir4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Repuest."
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":148EB1
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Cotizacion y Costos del Servicio"
         Top             =   710
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":14A633
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Adiciona Detalle"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":14AA75
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":14AEB7
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   720
         Width           =   765
      End
   End
   Begin VB.Frame FraDet6 
      BackColor       =   &H00000000&
      Caption         =   "SOLICITUD DE HERRAMIENTAS (COSTOS)"
      ForeColor       =   &H00FFFFC0&
      Height          =   1510
      Left            =   7890
      TabIndex        =   84
      Top             =   8160
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det6 
         Bindings        =   "frm_to_id_emergencia.frx":14B2F9
         Height          =   1215
         Left            =   60
         TabIndex        =   85
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "bien_total_venta"
            Caption         =   "Precio SubTotal"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Equipo"
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet5 
      BackColor       =   &H00000000&
      Caption         =   "SOLICITUD DE REPUESTOS (COSTOS)"
      ForeColor       =   &H00FFFFC0&
      Height          =   1520
      Left            =   7890
      TabIndex        =   83
      Top             =   6630
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det5 
         Bindings        =   "frm_to_id_emergencia.frx":14B314
         Height          =   1215
         Left            =   60
         TabIndex        =   86
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
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
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "bien_total_venta"
            Caption         =   "Precio SubTotal"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Equipo"
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton BtnImprimir3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Insumos"
      Height          =   640
      Left            =   990
      Picture         =   "frm_to_id_emergencia.frx":14B32F
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Cotizacion y Costos del Servicio"
      Top             =   7445
      Width           =   765
   End
   Begin VB.CommandButton BtnImprimir2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cotiza"
      Height          =   640
      Left            =   990
      Picture         =   "frm_to_id_emergencia.frx":14CAB1
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Cotizacion del Servicio"
      Top             =   5945
      Width           =   765
   End
   Begin VB.PictureBox FrmABMDet 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   13920
      Picture         =   "frm_to_id_emergencia.frx":14E233
      ScaleHeight     =   1380
      ScaleWidth      =   1635
      TabIndex        =   72
      Top             =   5195
      Width           =   1695
      Begin VB.CommandButton BtnAddDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":1BA265
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Adiciona Detalle"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":1BA6A7
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle 
         BackColor       =   &H80000018&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":1BAAE9
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   720
         Width           =   765
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "EMERGENCIA"
      ForeColor       =   &H00FFFFC0&
      Height          =   1500
      Left            =   7890
      TabIndex        =   70
      Top             =   5105
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "frm_to_id_emergencia.frx":1BAF2B
         Height          =   1215
         Left            =   75
         TabIndex        =   71
         Top             =   225
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bitacora_codigo"
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
            DataField       =   "negocia_fecha_real"
            Caption         =   "Fecha Evento"
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
            DataField       =   "negocia_tarea_realizada"
            Caption         =   "Tema Tratado"
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
            DataField       =   "negocia_observaciones"
            Caption         =   "Conclusiones u Observaciones"
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
            DataField       =   "negocia_hora_real"
            Caption         =   "Hora Evento"
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
            DataField       =   "negocia_gasto_estimado"
            Caption         =   "Gasto Estimado"
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
         BeginProperty Column06 
            DataField       =   "negocia_forma"
            Caption         =   "Tipo.Evento"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cliente Contactado"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Personal CGI"
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
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3030.236
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   3149.858
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox FrmABMDet3 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1440
      Left            =   120
      Picture         =   "frm_to_id_emergencia.frx":1BAF46
      ScaleHeight     =   1380
      ScaleWidth      =   1635
      TabIndex        =   66
      Top             =   6710
      Width           =   1695
      Begin VB.CommandButton BtnAddDetalle3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":225364
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Adiciona Producto"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":2257A6
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Modifica Producto Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAnlDetalle3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":225BE8
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Anula Producto Elegido"
         Top             =   720
         Width           =   765
      End
   End
   Begin VB.PictureBox FrmABMDet2 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   1430
      Left            =   120
      Picture         =   "frm_to_id_emergencia.frx":22602A
      ScaleHeight     =   1365
      ScaleWidth      =   1635
      TabIndex        =   62
      Top             =   5195
      Width           =   1695
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Borrar"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":29205C
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Elimina Detalle Elegido"
         Top             =   720
         Width           =   765
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modificar"
         Height          =   640
         Left            =   840
         Picture         =   "frm_to_id_emergencia.frx":29249E
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Modifica Detalle Elegido"
         Top             =   40
         Width           =   765
      End
      Begin VB.CommandButton BtnAddDetalle2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo"
         Height          =   640
         Left            =   50
         Picture         =   "frm_to_id_emergencia.frx":2928E0
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Adiciona Detalle"
         Top             =   40
         Width           =   765
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00400000&
      Height          =   1020
      Left            =   120
      Picture         =   "frm_to_id_emergencia.frx":292D22
      ScaleHeight     =   960
      ScaleWidth      =   15405
      TabIndex        =   51
      Top             =   60
      Width           =   15460
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "frm_to_id_emergencia.frx":2FED54
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "frm_to_id_emergencia.frx":2FEF5E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "frm_to_id_emergencia.frx":2FF3A0
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "frm_to_id_emergencia.frx":2FF5AA
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Listado"
         Height          =   720
         Left            =   4320
         Picture         =   "frm_to_id_emergencia.frx":2FFB62
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Lista de Tramites"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6000
         Picture         =   "frm_to_id_emergencia.frx":30011F
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "frm_to_id_emergencia.frx":300329
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "frm_to_id_emergencia.frx":300FF3
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "frm_to_id_emergencia.frx":3015D3
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TECNICO"
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
         Left            =   10515
         TabIndex        =   61
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00400000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "frm_to_id_emergencia.frx":301BF7
      ScaleHeight     =   915
      ScaleWidth      =   15405
      TabIndex        =   47
      Top             =   60
      Width           =   15460
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "frm_to_id_emergencia.frx":36DC29
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "frm_to_id_emergencia.frx":36DE33
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO2"
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
         Left            =   10020
         TabIndex        =   50
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00000000&
      Caption         =   "SOLICITUD DE INSUMOS (COSTOS)"
      ForeColor       =   &H00FFFFC0&
      Height          =   1510
      Left            =   1880
      TabIndex        =   30
      Top             =   6630
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det3 
         Bindings        =   "frm_to_id_emergencia.frx":36E03D
         Height          =   1215
         Left            =   60
         TabIndex        =   79
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "bien_total_venta"
            Caption         =   "Precio SubTotal"
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas / Observaciones"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo.Equipo"
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   4694.74
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3990.047
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_det4 
         Bindings        =   "frm_to_id_emergencia.frx":36E058
         Height          =   855
         Left            =   240
         TabIndex        =   95
         Top             =   240
         Visible         =   0   'False
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   17
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
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "COTIZACION DEL SERVICIO POR EQUIPO"
      ForeColor       =   &H00FFFFC0&
      Height          =   1520
      Left            =   1880
      TabIndex        =   24
      Top             =   5105
      Width           =   6000
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "frm_to_id_emergencia.frx":36E073
         Height          =   1215
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Codigo "
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
            DataField       =   "bien_codigo_anterior"
            Caption         =   "Nro.Eqp."
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
            DataField       =   "bien_total_venta"
            Caption         =   "Precio.X Servicio"
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
         BeginProperty Column03 
            DataField       =   "bien_cantidad_por_empaque"
            Caption         =   "Hrs.X Día"
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
            DataField       =   "marca_codigo"
            Caption         =   "Marca"
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
            DataField       =   "modelo_codigo"
            Caption         =   "Modelo"
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
         BeginProperty Column07 
            DataField       =   "bien_descripcion"
            Caption         =   "Descripcion del Bien"
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
            DataField       =   "bien_descripcion_anterior"
            Caption         =   "Caracteristicas/Identificacion.Ubicacion"
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
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   4470.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   2954.835
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00000000&
      Caption         =   "GERENCIA GENERAL"
      ForeColor       =   &H00FFFFC0&
      Height          =   3960
      Left            =   120
      TabIndex        =   14
      Top             =   1100
      Width           =   5895
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   3250
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   5741
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
            Caption         =   "Trámite"
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
         TabIndex        =   42
         Top             =   3585
         Value           =   -1  'True
         Width           =   1455
      End
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
         Left            =   3600
         TabIndex        =   43
         Top             =   3585
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3525
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00000000&
      Height          =   3960
      Left            =   6105
      TabIndex        =   11
      Top             =   1100
      Width           =   9495
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   9065
         TabIndex        =   78
         Top             =   3555
         Width           =   290
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6000
         TabIndex        =   46
         Top             =   525
         Width           =   290
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   9075
         TabIndex        =   41
         Top             =   1175
         Width           =   270
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "solicitud_observaciones"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   1605
      End
      Begin MSDataListLib.DataCombo dtc_codigo11 
         Bindings        =   "frm_to_id_emergencia.frx":36E08E
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3360
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux1 
         Bindings        =   "frm_to_id_emergencia.frx":36E0A8
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3600
         TabIndex        =   31
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
         Bindings        =   "frm_to_id_emergencia.frx":36E0C1
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7680
         TabIndex        =   21
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "subproceso_codigo"
         BoundColumn     =   "subproceso_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_to_id_emergencia.frx":36E0DA
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3060
         TabIndex        =   3
         Top             =   2760
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_aux3 
         Bindings        =   "frm_to_id_emergencia.frx":36E0F4
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "codigo5"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_to_id_emergencia.frx":36E10D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7440
         TabIndex        =   17
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
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_to_id_emergencia.frx":36E126
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
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
         Height          =   315
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2280
         Width           =   8145
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "frm_to_id_emergencia.frx":36E13F
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4740
         TabIndex        =   15
         Top             =   1155
         Width           =   4620
         _ExtentX        =   8149
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
         Bindings        =   "frm_to_id_emergencia.frx":36E158
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4560
         TabIndex        =   19
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
         Bindings        =   "frm_to_id_emergencia.frx":36E171
         DataField       =   "subproceso_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4635
         TabIndex        =   20
         Top             =   3540
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "subproceso_descripcion"
         BoundColumn     =   "subproceso_codigo"
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
         Bindings        =   "frm_to_id_emergencia.frx":36E18A
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1575
         TabIndex        =   0
         Top             =   510
         Width           =   4725
         _ExtentX        =   8334
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
         Bindings        =   "frm_to_id_emergencia.frx":36E1A3
         DataField       =   "poa_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   2760
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   4210752
         ForeColor       =   16777215
         ListField       =   "codigo"
         BoundColumn     =   "codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc11 
         Bindings        =   "frm_to_id_emergencia.frx":36E1BD
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   1800
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "solicitud_fecha_solicitud"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7875
         TabIndex        =   82
         Top             =   1800
         Width           =   1480
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85786625
         CurrentDate     =   41678
         MaxDate         =   55153
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_to_id_emergencia.frx":36E1D7
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   180
         TabIndex        =   106
         Top             =   1155
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         ListField       =   "descripcion"
         BoundColumn     =   "codigo"
         Text            =   ""
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
         TabIndex        =   77
         Top             =   3540
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Tipo de Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   1
         Left            =   4680
         TabIndex        =   76
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cite del Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   6
         Left            =   6630
         TabIndex        =   45
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
         Left            =   6495
         TabIndex        =   44
         Top             =   510
         Width           =   1815
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Edificio"
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
         TabIndex        =   40
         Top             =   885
         Width           =   660
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
         TabIndex        =   39
         Top             =   2300
         Width           =   915
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
         TabIndex        =   38
         Top             =   2775
         Width           =   1635
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   180
         TabIndex        =   37
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label lbl_campo11 
         BackColor       =   &H00000000&
         Caption         =   "Responsable CGI:"
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
         Top             =   1815
         Width           =   1815
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cliente"
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
         Left            =   4740
         TabIndex        =   35
         Top             =   885
         Width           =   615
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Left            =   1605
         TabIndex        =   34
         Top             =   225
         Width           =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9495
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   9495
         Y1              =   1620
         Y2              =   1620
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
         TabIndex        =   29
         Top             =   480
         Width           =   1215
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
         TabIndex        =   28
         Top             =   3540
         Width           =   1605
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   13
         Left            =   2400
         TabIndex        =   23
         Top             =   3240
         Width           =   1785
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Fecha Registro"
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
         Left            =   6465
         TabIndex        =   22
         Top             =   1820
         Width           =   1380
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
         Left            =   8505
         TabIndex        =   4
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Cod.Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   225
         Width           =   1110
      End
      Begin VB.Label lblLabels 
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
         ForeColor       =   &H00FFFF80&
         Height          =   240
         Index           =   2
         Left            =   8595
         TabIndex        =   12
         Top             =   225
         Width           =   645
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
      ScaleWidth      =   21360
      TabIndex        =   5
      Top             =   12930
      Width           =   21360
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
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9840
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
      Left            =   9840
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   2280
      Top             =   9840
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
      Left            =   4440
      Top             =   9840
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
      Left            =   6720
      Top             =   9840
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
      Left            =   9000
      Top             =   9840
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
      Left            =   11280
      Top             =   9840
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
      Left            =   13560
      Top             =   9840
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
      Left            =   120
      Top             =   10200
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
      Left            =   2280
      Top             =   10200
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
      Left            =   4440
      Top             =   10200
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
      Left            =   11280
      Top             =   10200
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
      Left            =   13560
      Top             =   10200
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
      Left            =   6720
      Top             =   10200
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
   Begin MSAdodcLib.Adodc Ado_detalle7 
      Height          =   330
      Left            =   9000
      Top             =   10200
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
      Caption         =   "Ado_detalle7"
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   120
      Top             =   10560
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
      Caption         =   "Ado_detalle3"
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
   Begin MSAdodcLib.Adodc Ado_detalle4 
      Height          =   330
      Left            =   2400
      Top             =   10560
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
      Caption         =   "Ado_detalle4"
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
   Begin MSAdodcLib.Adodc Ado_detalle5 
      Height          =   330
      Left            =   4680
      Top             =   10560
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
      Caption         =   "Ado_detalle5"
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
   Begin MSAdodcLib.Adodc Ado_detalle6 
      Height          =   330
      Left            =   6960
      Top             =   10560
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
      Caption         =   "Ado_detalle6"
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
   Begin Crystal.CrystalReport CR02 
      Left            =   10320
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   10800
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR00 
      Left            =   9360
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "frm_to_id_emergencia"
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

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset
Dim rs_det3 As New ADODB.Recordset
Dim rs_det4 As New ADODB.Recordset
Dim rs_det5 As New ADODB.Recordset
Dim rs_det6 As New ADODB.Recordset
Dim rs_det7 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod, VAR_DET As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim parametro As String
Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_TIPO, VAR_SOL As Integer

Dim iResult As Integer

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet1.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    Fra_datos.Enabled = False
    VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    Call ABRIR_TABLA_DET
    frm_ao_bitacora_emergencia.txt_codigo.Caption = Me.txt_codigo.Caption
    frm_ao_bitacora_emergencia.Txt_campo1.Caption = Me.dtc_codigo1.Text
    frm_ao_bitacora_emergencia.Txt_descripcion.Caption = Me.dtc_desc1.Text
    frm_ao_bitacora_emergencia.Txt_Correl.Caption = 0    'rs_datos!correl_bitacora + 1
    frm_ao_bitacora_emergencia.Txt_estado.Caption = "REG"
    frm_ao_bitacora_emergencia.lbl_bitacora.Caption = Me.FraDet1.Caption
    Ado_detalle1.Recordset.AddNew
    Txt_campo2.Visible = False
'    Txt_campo6.Visible = False
'    Txt_campo7.Visible = False
'    Txt_campo8.Visible = False
'    Txt_campo9.Visible = False
'    Txt_campo10.Visible = False
    frm_ao_bitacora_emergencia.Show vbModal
    
    Call ABRIR_TABLA_DET
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet1.Enabled = True
    FrmABMDet.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    'Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAddDetalle2_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
    Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
        Case "2"    'SOLO VENTA DE BIENES
        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
            

        Case "COM-01"    '4. COMPRA-VENTA DE EQUIPOS
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
'            mw_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
'            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Show vbModal
        Case "COM-02"    '3. VENTA DE SERVICIOS (PROVISION)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Show vbModal
        Case "COM-03"    '4. VENTA DE SERVICIOS (INSTALACIONES)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Show vbModal
        Case "COM-04"    '5. VENTA DE SERVICIOS (AJUSTE)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Show vbModal
        Case "TEC-01"    '6. VENTA DE SERVICIOS (MANTENIMIENTO GRATUITO)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal

        Case "TEC-02"    '10. VENTA DE SERVICIOS MANTENIMIENTO PREVENTIVO
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        Case "TEC-03"    '7. VENTA DE SERVICIOS REPARACION
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        Case "TEC-04"    '8. VENTA DE SERVICIOS (EMERGENCIAS)
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        Case "TEC-05"    '9. SERVICIO MODERNIZACION    End Select
            Call ABRIR_TABLA_DET
            Ado_detalle2.Recordset.AddNew
            frm_solicitud_bienes.txt_codigo.Caption = Me.txt_codigo.Caption
            frm_solicitud_bienes.Txt_campo1.Caption = Me.dtc_codigo1.Text
            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
            frm_solicitud_bienes.lbl_det.Caption = "43340"
            frm_solicitud_bienes.Txt_estado.Caption = "REG"
            frm_solicitud_bienes.Show vbModal
        End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnAddDetalle3_Click()
    VAR_DET = "30000"
    Call NuevoDetalle
    ''grupo_codigo = '30000' and (par_codigo <> '39800' and par_codigo <> '34800')
End Sub

Private Sub NuevoDetalle()
  marca1 = Ado_datos.Recordset.Bookmark
  If rs_datos!estado_codigo = "REG" Then
    swnuevo = 1
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False
'    Select Case dtc_codigo2.Text
'        Case "1"    'SOLO COMPRAS BB y SS
'        Case "2"    'SOLO VENTA DE BIENES
'        Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'
'
'        Case "TEC-01"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
''            mw_solicitud_edificacion.dtc_codigo1.Text = Me.dtc_codigo3.Text
''            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
''            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
''            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
''            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'        Case "COM-02"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'        Case "COM-03"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'        Case "COM-04"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'
'        Case "TEC-01"    '6. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'
'            Set rs_det1 = New ADODB.Recordset
'            If rs_det1.State = 1 Then rs_det1.Close
'            rs_det1.Open "select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and cotiza_codigo = " & Ado_datos.Recordset!cotiza_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
'            Set Ado_detalle1.Recordset = rs_det1
'            Set dg_det1.DataSource = Ado_detalle1.Recordset
'
'        Case "TEC-02"    '10. VENTA DE SERVICIOS (MANT)
'            Call ABRIR_TABLA_DET
'            If VAR_DET = "30000" Then
'                Ado_detalle3.Recordset.AddNew
'            End If
'            If VAR_DET = "39800" Then
'                Ado_detalle5.Recordset.AddNew
'            End If
'            If VAR_DET = "34800" Then
'                Ado_detalle6.Recordset.AddNew
'            End If
'            If VAR_DET = "24300" Then
'                Ado_detalle7.Recordset.AddNew
'            End If
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.lbl_det.Caption = VAR_DET     '"34110"
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'
'        Case "TEC-03"    '7. VENTA DE SERVICIOS (REP)
            Call ABRIR_TABLA_DET
            If VAR_DET = "30000" Then
                Ado_detalle3.Recordset.AddNew
                frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes3.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes3.Txt_estado.Caption = "REG"
                frm_solicitud_bienes3.Show vbModal
            End If
            If VAR_DET = "39800" Then
                Ado_detalle5.Recordset.AddNew
                frm_solicitud_bienes5.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes5.Txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes5.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes5.Txt_estado.Caption = "REG"
                frm_solicitud_bienes5.Show vbModal
            End If
            If VAR_DET = "34800" Then
                Ado_detalle6.Recordset.AddNew
                frm_solicitud_bienes6.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes6.Txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes6.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes6.Txt_estado.Caption = "REG"
                frm_solicitud_bienes6.Show vbModal
            End If
            If VAR_DET = "24300" Then
                Ado_detalle7.Recordset.AddNew
                frm_solicitud_bienes7.txt_codigo.Caption = Me.txt_codigo.Caption
                frm_solicitud_bienes7.Txt_campo1.Caption = Me.dtc_codigo1.Text
                frm_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
                frm_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes7.lbl_det.Caption = VAR_DET     '"34110"
                frm_solicitud_bienes7.Txt_estado.Caption = "REG"
                frm_solicitud_bienes7.Show vbModal
            End If
            
'        Case "TEC-04"    '8. VENTA DE SERVICIOS (EME)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.lbl_det.Caption = "34110"
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'        Case "TEC-05"    '9. VENTA DE SERVICIOS (MOD)
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.AddNew
'            frm_solicitud_bienes3.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes3.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes3.lbl_det.Caption = "34110"
'            frm_solicitud_bienes3.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes3.Show vbModal
'
'    End Select
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnAddDetalle5_Click()
    VAR_DET = "39800"
    Call NuevoDetalle
End Sub

Private Sub BtnAddDetalle6_Click()
    VAR_DET = "34800"
    Call NuevoDetalle
End Sub

Private Sub BtnAddDetalle7_Click()
    VAR_DET = "24300"
    Call NuevoDetalle
End Sub

Private Sub BtnAnlDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo --> " + Str(Ado_detalle1.Recordset!bitacora_codigo), vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, un registro Aprobado o Anulado, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
 Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
 End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   If Ado_detalle2.Recordset.RecordCount > 0 Then
       If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle2.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede BORRAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle3_Click()
   If Ado_detalle3.Recordset.RecordCount > 0 Then
       If Ado_detalle3.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle3.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle3.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnAnlDetalle5_Click()
   If Ado_detalle5.Recordset.RecordCount > 0 Then
       If Ado_detalle5.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle5.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle5.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle6_Click()
   If Ado_detalle6.Recordset.RecordCount > 0 Then
       If Ado_detalle6.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle6.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle6.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle7_Click()
   If Ado_detalle7.Recordset.RecordCount > 0 Then
       If Ado_detalle7.Recordset("estado_codigo") = "REG" Then
          sino = MsgBox("Está Seguro de BORRAR el Registro Activo --> " + Ado_detalle7.Recordset!bien_codigo, vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
            db.Execute "delete ao_solicitud_bienes Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and bien_codigo = '" & Ado_detalle7.Recordset!bien_codigo & "' "
            Call ABRIR_TABLA_DET
          End If
       Else
            MsgBox "No se puede ANULAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede ANULAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If

End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
'  If Ado_datos.Recordset.RecordCount > 0 Then
'   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
'        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
'        Exit Sub
'   End If
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
   If rs_datos!estado_codigo = "REG" Then       'And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "TEC-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
                Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_detalle1.Recordset!edif_codigo & "'  "
                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
                'If rs_aux1.RecordCount > 0 Then
                '    MsgBox "El código ya existe, consulte con el administrador del Sistema..."
                '    var_cod = 0
                '    Exit Sub
                'Else
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        VAR_AUX = rs_aux2!Codigo
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_detalle1.Recordset!edif_codigo
                    rs_aux1!trafico_codigo = var_cod
                    rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!Fecha_Registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
                'End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
            
            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
            'Case "TEC-05"    '5. SERVICIO MODERNIZACION
            Case "TEC-02", "TEC-03", "TEC-04", "TEC-05"     '10. SERVICIO MANTENIMIENTO
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               If dtc_codigo2.Text = "TEC-02" Then
                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               Else
                    rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, SUM(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               End If
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!depto_codigo = Left(Ado_datos.Recordset!edif_codigo, 1)
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!Fecha_Registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
                    Set rs_aux3 = New ADODB.Recordset
                    If rs_aux3.State = 1 Then rs_aux3.Close
                    'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & Year(Date) & "'   ", db, adOpenKeyset, adLockOptimistic
                    'If rs_aux3.RecordCount > 0 Then
                        'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    'Else
                        VAR_AUX = rs_aux3.RecordCount + 1
                        rs_aux3.AddNew
                        rs_aux3!ges_gestion = glGestion         'Year(Date)
                        rs_aux3!venta_codigo = var_cod
                        rs_aux3!venta_codigo_det = VAR_AUX
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = 0
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = 0
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
                        rs_aux3!modelo_elegido_h = "N"
                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!Fecha_Registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
    
                     rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique la Venta de Productos..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        Case "COM-03"    '3. SERVICIO INSTALACION
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!Fecha_Registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
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
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = 0
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = 0
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
                        rs_aux3!modelo_elegido_h = "N"
                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!Fecha_Registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                     rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique la Venta de Productos..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        Case "COM-04"    '4. SERVICIO AJUSTE
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
               Set rs_aux4 = New ADODB.Recordset
               If rs_aux4.State = 1 Then rs_aux4.Close
               'rs_aux4.Open "select sum(bien_precio_compra) as totbs2, sum(bien_total_compra) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               rs_aux4.Open "select sum(bien_precio_venta_base) as totbs2, sum(bien_total_venta) as totdl2, avg(bien_cantidad) as cant2  from ao_solicitud_bienes where unidad_codigo ='" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo =" & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic
               If IsNull(rs_aux4!totbs2) Then
                    'If CDbl(TxtMonto) > Ado_datos.Recordset!venta_monto_total_bs Then
                        MsgBox "No puede Aprobar, debe registrar <" + FraDet2.Caption + "> !! Vuelva a Intentar ...", vbExclamation, "Atención"
                        If rs_aux4.State = 1 Then rs_aux4.Close
                        Exit Sub
                    'End If
               Else

               Set rs_aux1 = New ADODB.Recordset
                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
               SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
               rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
               If rs_aux1.RecordCount > 0 Then
                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
                    '    var_cod = 0
                    '    Exit Sub
                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + rs_aux4!totdl2
                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + rs_aux4!totdl2 / GlTipoCambioOficial
               Else
                    'CREA VENTA CABECERA
                    Set rs_aux2 = New ADODB.Recordset
                    If rs_aux2.State = 1 Then rs_aux2.Close
                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
                    If Not rs_aux2.EOF Then
                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
                    End If
                    rs_aux1.AddNew
                    'var_cod = rs_aux1.RecordCount + 1
                    rs_aux1!ges_gestion = Year(Date)
                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
                    rs_aux1!venta_codigo = var_cod
                    rs_aux1!beneficiario_codigo = Ado_datos.Recordset!beneficiario_codigo
                    rs_aux1!venta_monto_total_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_monto_total_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_monto_cobrado_bs = 0
                    rs_aux1!venta_monto_cobrado_dol = 0
                    rs_aux1!venta_saldo_p_cobrar_bs = rs_aux4!totdl2                        'Ado_datos.Recordset!cotiza_precio_total_bs
                    rs_aux1!venta_saldo_p_cobrar_dol = rs_aux4!totdl2 / GlTipoCambioOficial 'Ado_datos.Recordset!cotiza_precio_total_dol
                    rs_aux1!venta_cantidad_total = rs_aux4!cant2
                    rs_aux1!venta_fecha = Ado_datos.Recordset!solicitud_fecha_solicitud
                    rs_aux1!venta_fecha_inicio = Ado_datos.Recordset!solicitud_fecha_solicitud
                    'VAR_CONT2 = 365 / 30 * rs_aux4!cant2
                    rs_aux1!venta_plazo_dias_calendario = 0 'VAR_CONT2
                    rs_aux1!correl_cobro_prog = 0
                    rs_aux1!venta_fecha_fin = FormatDateTime(Ado_datos.Recordset!solicitud_fecha_solicitud + VAR_CONT2, vbGeneralDate)
                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
                    rs_aux1!estado_codigo = "REG"
                    rs_aux1!Fecha_Registro = Date
                    rs_aux1!usr_codigo = glusuario
                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               End If
                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
               If var_cod = "" Then
                    var_cod = rs_aux1!venta_codigo
               End If
                'GRABA VENTA DETALLE
                'wwwwwwwwwwwwwwwwwww
               Set rs_aux5 = New ADODB.Recordset
               If rs_aux5.State = 1 Then rs_aux5.Close
               rs_aux5.Open "select * from ao_solicitud_bienes where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockBatchOptimistic   'and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'
               'Set AdoAux.Recordset = rsAuxDetalle
               If rs_aux5.RecordCount > 0 Then
                   'AdoAux.Recordset.MoveFirst
                  rs_aux5.MoveFirst
                  While Not rs_aux5.EOF   ' AdoAux.Recordset.EOF
    
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
                        rs_aux3!bien_codigo = rs_aux5!bien_codigo
                        rs_aux3!venta_det_cantidad = rs_aux5!bien_cantidad
                        rs_aux3!venta_precio_unitario_bs = rs_aux5!bien_precio_venta_base
                        rs_aux3!venta_descuento_bs = 0
                        rs_aux3!venta_precio_total_bs = rs_aux5!bien_total_venta
                        rs_aux3!venta_precio_unitario_dol = rs_aux5!bien_precio_venta_base / GlTipoCambioOficial
                        rs_aux3!venta_descuento_dol = 0
                        rs_aux3!venta_precio_total_dol = rs_aux5!bien_total_venta / GlTipoCambioOficial
                        'rs_aux3!concepto_venta = dtc_desc2.Text + " - " + Trim(dtc_desc3.Text)
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        rs_aux6.Open "Select * from ac_bienes where bien_codigo = '" & rs_aux3!bien_codigo & "'  ", db, adOpenKeyset, adLockOptimistic
                        If rs_aux6.RecordCount > 0 Then
                            rs_aux3!concepto_venta = rs_aux6!bien_descripcion '+ " - " + Trim(dtc_desc3.Text)
                        Else
                            rs_aux3!concepto_venta = "NA1"
                        End If
                        rs_aux3!modelo_codigo = rs_aux5!modelo_codigo
                        rs_aux3!grupo_codigo = rs_aux5!grupo_codigo
                        rs_aux3!subgrupo_codigo = rs_aux5!subgrupo_codigo
                        rs_aux3!par_codigo = rs_aux5!par_codigo
                        'ok
                        rs_aux3!bien_cantidad_por_empaque = rs_aux5!bien_cantidad_por_empaque
                        'If rs_aux5!par_codigo = "43340" Or rs_aux5!par_codigo = "99990" Then
                        If rs_aux5!par_codigo = "43340" Then
                            db.Execute "update ao_ventas_cabecera set unimed_codigo = '" & rs_aux5!unimed_codigo & "' WHERE venta_codigo = " & var_cod & ""
                        End If
                        rs_aux3!tipo_descuento = 0
                        rs_aux3!almacen_codigo = 0
                        rs_aux3!modelo_codigo1 = rs_aux5!modelo_codigo 'do_datos.Recordset!modelo_codigo
                        rs_aux3!modelo_codigo_h = "S/M" 'Ado_datos.Recordset!modelo_codigo_h
                        rs_aux3!modelo_codigo_x = "S/M" 'Ado_datos.Recordset!modelo_codigo_x
                        rs_aux3!modelo_elegido = "N"
                        rs_aux3!modelo_elegido_h = "N"
                        rs_aux3!modelo_elegido_x = "N"
                        rs_aux3!estado_codigo = "REG"
                        rs_aux3!Fecha_Registro = Date
                        rs_aux3!usr_codigo = glusuario
                        rs_aux3.Update
                     rs_aux5.MoveNext
                  Wend
               Else
                    MsgBox "Error Verifique la Venta de Productos..."
               End If
              End If
            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        End Select
        Set rs_aux2 = New ADODB.Recordset
        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
            Txt_campo1.Caption = rs_aux2!correl_doc
            rs_aux2.Update
        End If
        rs_datos!doc_numero = Txt_campo1.Caption
        'REVISAR !!! JQA 2014_07_08
        'VAR_ARCH = RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
        VAR_ARCH = "TEC_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(Txt_campo1.Caption)))
        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!estado_codigo = "APR"
        rs_datos!Fecha_Registro = Date
        rs_datos!usr_codigo = glusuario
        rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
'  Else
'      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
'  End If
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        buscados = 1
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
        If OptFilGral1.Value = True Then
            MsgBox "Esta Buscando los Registros... " + OptFilGral1.Caption, vbInformation, "Atención!"
        Else
            MsgBox "Esta Buscando... " + OptFilGral2.Caption + " los Registros.", vbInformation, "Atención!"
        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
'        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
      OptFilGral1.Visible = True
      OptFilGral2.Visible = True
    End If
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
    If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
    If rs_datos!estado_codigo = "APR" Then
       sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If sino = vbYes Then
          rs_datos!estado_codigo = "ERR"
          rs_datos!Fecha_Registro = Date
          rs_datos!usr_codigo = glusuario
          rs_datos.UpdateBatch adAffectAll
       End If
    Else
       MsgBox "No se puede ANULAR un registro Elaborado o Errado ...", vbExclamation, "Validación de Registro"
    End If
  Else
      MsgBox "NO se puede ANULAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
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
         rs_datos!Fecha_Registro = Date
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
    If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "Select max(solicitud_codigo) as Codigo from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        'SQL_FOR = "select * from ao_solicitud where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        'If rs_aux1.RecordCount > 0 Then
        If Not rs_aux1.EOF Then
            'var_cod = rs_aux1.RecordCount + 1
            var_cod = IIf(IsNull(rs_aux1!Codigo), 1, rs_aux1!Codigo + 1)
            'MsgBox "El código ya existe, consulte con el administrador del Sistema..."
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
     If VAR_SW = "MOD" Then
        VAR_UNI = rs_datos!unidad_codigo
        var_cod = rs_datos!solicitud_codigo
     End If
     rs_datos!solicitud_fecha_solicitud = DTPfecha1.Value
     'rs_datos!solicitud_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     rs_datos!solicitud_justificacion = Txt_descripcion.Text
     Select Case dtc_codigo2.Text
        Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL - Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "CMX-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
            rs_datos!proceso_codigo = "CMX"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "CMX-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "CMX-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-XXX"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-02"    '3. COMPRA-VENTA BB Y SS - COMERCIAL -         Case "2"    'SOLO VENTA DE BIENES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-01-02"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "COM"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-234"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-03"    'VENTA DE SERVICIOS INSTTALACIONES
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "COM-04" '5       'VENTA DE SERVICIOS AJUSTE
            rs_datos!proceso_codigo = "COM"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "COM-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "COM-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "TEC-01"    '6. SERVICIO MANTENIMIENTO GRATUITO
            rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = "TEC-01"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
            rs_datos!etapa_codigo = "TEC-01-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "TEC-02"    '10. SERVICIO MANTENIMIENTO PREVENTIVO
            'If VAR_UNI = "DNMAN" Then
            rs_datos!solicitud_tipo = "10"
            rs_datos!proceso_codigo = Left(dtc_codigo2.Text, 3) ' "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
            rs_datos!subproceso_codigo = IIf(dtc_codigo2.Text = "", "TEC-02", dtc_codigo2.Text)
            rs_datos!etapa_codigo = Trim(dtc_codigo2.Text) + "-01"  '"TEC-02-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
            rs_datos!clasif_codigo = Left(dtc_codigo2.Text, 3)  '"TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-355"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            'End If
        Case "TEC-03" '10 REPARACION    If VAR_UNI = "DNIREP" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-03"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TEC-03-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
        Case "TEC-04" '10 EMERGENCIAS   If VAR_UNI = "DNEME" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-04"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TEC-04-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
            
        Case "TEC-05"    '5. SERVICIO MODERNIZACION -If VAR_UNI = "DNMOD" Then
                rs_datos!proceso_codigo = "TEC"         'IIf(dtc_codigo5.Text = "", "COM", dtc_codigo5.Text)
                rs_datos!subproceso_codigo = "TEC-05"   'IIf(dtc_codigo6.Text = "", "COM-03", dtc_codigo6.Text)
                rs_datos!etapa_codigo = "TES-05-01"     'IIf(dtc_codigo7.Text = "", "COM-03-02", dtc_codigo7.Text)
                rs_datos!clasif_codigo = "TEC"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
                rs_datos!doc_codigo = "R-362"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
     End Select
     rs_datos!poa_codigo = IIf(dtc_codigo10.Text = "", "3.2.6", dtc_codigo10.Text)
     rs_datos!solicitud_observaciones = txt_obs.Text
     rs_datos!solicitud_fecha_recepción = DTPfecha1.Value
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
'     rs_datos!solicitud_codigo_ant = 0
     rs_datos!usr_codigo_aprueba = ""
     rs_datos!fecha_aprueba = Date
     rs_datos!hora_aprueba = ""
     'rs_datos!Foto = Date
     'rs_datos!ARCHIVO_Foto = var_cod + ".JPG"
     'rs_datos!archivo_foto_cargado = "N"
     'hora_registro
     rs_datos!Fecha_Registro = Date     'no cambia
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
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo11.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo11.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (dtc_codigo8.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo9.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo9.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If (dtc_codigo10.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo10.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar ... " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CR00.ReportFileName = App.Path & "\Reportes\comercial\ar_solicitud_cotizacion.rpt"
        CR00.ReportFileName = App.Path & "\Reportes\tecnico\tr_lista_solicitud_tecnico.rpt"
        
        CR00.WindowShowPrintSetupBtn = True
        CR00.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          'CR00.Formulas(1) = "cod_unidad = '" & adosolicitud.Recordset!codigo_unidad & "' "
          'CR00.Formulas(6) = "tc = " & GlTipoCambioOficial & " "
        'Call CREAVISTAF11          'JQA JUN-2008
        'CR00.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        'CR00.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        'CR00.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        CR00.StoredProcParam(0) = Me.Ado_datos.Recordset!unidad_codigo
        Select Case Me.Ado_datos.Recordset!unidad_codigo
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Mantenimiento"
          Case "DNREP"
              var_titulo = "Módulo Reparaciones"
          Case "DNEME"
              var_titulo = "Módulo Emergencias"
          Case "DNMOD"
              var_titulo = "Módulo Modernización"
      End Select
      CR00.Formulas(3) = "titulo = '" & var_titulo & "' "
      CR00.Formulas(4) = "subtitulo = '" & lbl_titulo.Caption & "' "

        iResult = CR00.PrintReport
        If iResult <> 0 Then MsgBox CR00.LastErrorNumber & " : " & CR00.LastErrorString, vbCritical, "Error de impresión"
        CR00.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        cr01.ReportFileName = App.Path & "\Reportes\tecnico\tr_emergencia_bitacora.rpt"
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          cr01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          cr01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        cr01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        cr01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        cr01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos... " & FraDet1.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnImprimir2_Click()
    Select Case parametro
        Case "DVTA"             'INI COMERCIAL
            'dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            'dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            'dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            'dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            If (Ado_datos.Recordset.RecordCount > 0) Then
              If Ado_detalle2.Recordset.RecordCount > 0 Then
                  'Dim iResult As Integer
                  'Dim co As New ADODB.Command
                  CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_propuesta.rpt"
                  CR02.WindowShowPrintSetupBtn = True
                  CR02.WindowShowRefreshBtn = True
                  'MsgBox rs.RecordCount
                    CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                    CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                    CR02.Formulas(2) = "Subtitulo2 = '" & dtc_desc2.Text & "' "
                  'Call CREAVISTAF11          'JQA JUN-2008
                  CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                  CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                  CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                  iResult = CR02.PrintReport
                  If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                  CR02.WindowState = crptMaximized
              Else
                  MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
              End If
            Else
              MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
            End If
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
              If (Ado_datos.Recordset.RecordCount > 0) Then
                  If Ado_detalle2.Recordset.RecordCount > 0 Then
                      
                      'Dim co As New ADODB.Command
                      'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_solicitud_cotizacion.rpt"
                      CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_reparacion.rpt"
                      CR02.WindowShowPrintSetupBtn = True
                      CR02.WindowShowRefreshBtn = True
                      'MsgBox rs.RecordCount
                        CR02.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
                        CR02.Formulas(1) = "Subtitulo = '" & FraDet2.Caption & "' "
                      'Call CREAVISTAF11          'JQA JUN-2008
                      CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
                      CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
                      CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
                      iResult = CR02.PrintReport
                      If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
                      CR02.WindowState = crptMaximized
                  Else
                      MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet2.Caption, , "Atención"
                  End If
                Else
                  MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
                End If

        Case "DNEME"            'EMERGENCIAS
            'dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            'dtc_codigo2.Text = "TEC-05" '10
        Case Else
            'dtc_codigo2.Text = "TEC-01"   '3
    End Select
  
End Sub

Private Sub BtnImprimir3_Click()
 
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle3.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet3.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet3.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir4_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle5.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet5.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet5.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnImprimir7_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle7.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cotizacion_costos.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        'MsgBox rs.RecordCount
          CR03.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          CR03.Formulas(1) = "Subtitulo = '" & FraDet7.Caption & "' "
        'Call CREAVISTAF11          'JQA JUN-2008
        CR03.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        CR03.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        CR03.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos de... " & FraDet7.Caption, , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub

Private Sub BtnModDetalle_Click()
  If Ado_detalle1.Recordset.RecordCount > 0 Then
    If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
      marca1 = Ado_datos.Recordset.Bookmark
      swnuevo = 2
      fraOpciones.Enabled = False
      FraNavega.Enabled = False
      FraDet1.Enabled = False
      FrmABMDet.Enabled = False
      FraDet2.Enabled = False
      FrmABMDet2.Enabled = False
      Fra_datos.Enabled = False
      VAR_SOL = Ado_datos.Recordset!solicitud_codigo
      frm_ao_bitacora_emergencia.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
      frm_ao_bitacora_emergencia.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
      frm_ao_bitacora_emergencia.Txt_descripcion.Caption = Me.dtc_desc1.Text
      frm_ao_bitacora_emergencia.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
      'frm_ao_bitacora_emergencia.Txt_estado.Caption = "REG"
      'Ado_detalle1.Recordset.AddNew
       
      frm_ao_bitacora_emergencia.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("negocia_forma")
      frm_ao_bitacora_emergencia.DTPfecha1.Value = Me.Ado_detalle1.Recordset("negocia_fecha_real")
      frm_ao_bitacora_emergencia.Txt_campo2.Text = Me.Ado_detalle1.Recordset("negocia_hora_real")
      frm_ao_bitacora_emergencia.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_envio), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_envio)
      frm_ao_bitacora_emergencia.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_llegada), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_llegada)
      frm_ao_bitacora_emergencia.Txt_campo8.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_mora), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_mora)
      frm_ao_bitacora_emergencia.Txt_campo9.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_salida), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_salida)
      frm_ao_bitacora_emergencia.Txt_campo10.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!negocia_hora_trabajo), "00:00", Me.Ado_detalle1.Recordset!negocia_hora_trabajo)
      
      frm_ao_bitacora_emergencia.dtc_codigo4.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!tipo_falla), "", Me.Ado_detalle1.Recordset!tipo_falla)
      frm_ao_bitacora_emergencia.dtc_desc4.BoundText = frm_ao_bitacora_emergencia.dtc_codigo4.BoundText
      
      frm_ao_bitacora_emergencia.dtc_codigo5.Text = IIf(IsNull(Me.Ado_detalle1.Recordset!falla_codigo), "", Me.Ado_detalle1.Recordset!falla_codigo)
      frm_ao_bitacora_emergencia.dtc_desc5.BoundText = frm_ao_bitacora_emergencia.dtc_codigo5.BoundText
      
      frm_ao_bitacora_emergencia.Txt_monto1.Text = Me.Ado_detalle1.Recordset("negocia_gasto_estimado")
      frm_ao_bitacora_emergencia.dtc_codigo2.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo")
      frm_ao_bitacora_emergencia.dtc_codigo3.Text = Me.Ado_detalle1.Recordset("beneficiario_codigo_resp")
      frm_ao_bitacora_emergencia.Txt_campo3.Text = Me.Ado_detalle1.Recordset("negocia_tarea_realizada")
      frm_ao_bitacora_emergencia.Txt_campo4.Text = Me.Ado_detalle1.Recordset("negocia_observaciones")
      frm_ao_bitacora_emergencia.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bitacora_cite")
      If swnuevo = 2 Then
          frm_ao_bitacora_emergencia.dtc_desc1.BoundText = frm_ao_bitacora_emergencia.dtc_codigo1.BoundText
          frm_ao_bitacora_emergencia.dtc_desc2.BoundText = frm_ao_bitacora_emergencia.dtc_codigo2.BoundText
          frm_ao_bitacora_emergencia.dtc_desc3.BoundText = frm_ao_bitacora_emergencia.dtc_codigo3.BoundText
'          If frm_ao_bitacora_emergencia.Txt_campo2 = ":" Then
'            frm_ao_bitacora_emergencia.HH = "00"    'Left(frm_ao_bitacora_emergencia.Txt_campo2, 2)
'            frm_ao_bitacora_emergencia.MM = "00"    'Right(frm_ao_bitacora_emergencia.Txt_campo2, 2)
'          Else
'            frm_ao_bitacora_emergencia.HH = Left(frm_ao_bitacora_emergencia.Txt_campo2.Text, 2)
'            frm_ao_bitacora_emergencia.MM = Right(frm_ao_bitacora_emergencia.Txt_campo2.Text, 2)
'          End If
          
      End If
      Txt_campo2.Visible = False
'      Txt_campo6.Visible = False
'      Txt_campo7.Visible = False
'      Txt_campo8.Visible = False
'      Txt_campo9.Visible = False
'      Txt_campo10.Visible = False
      frm_ao_bitacora_emergencia.Show vbModal
      
      Call ABRIR_TABLA_DET
      
      swnuevo = 0
      fraOpciones.Enabled = True
      FraNavega.Enabled = True
      FraDet1.Enabled = True
      FrmABMDet.Enabled = True
      FraDet2.Enabled = True
      FrmABMDet2.Enabled = True
      'Fra_datos.Enabled = True
      Call ABRIR_TABLA_DET
      Call OptFilGral1_Click
      Ado_datos.Recordset.Move marca1 - 1
    Else
      MsgBox "No se puede Modificar un registro APROBADO o ANULADO, Verifique por favor ...!! ", vbExclamation
    End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
  If Ado_detalle2.Recordset.RecordCount > 0 Then
      If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
        marca1 = Ado_detalle2.Recordset.Bookmark
        swnuevo = 2
        fraOpciones.Enabled = False
        FraNavega.Enabled = False
        FraDet2.Enabled = False
        FrmABMDet2.Enabled = False
        FraDet3.Enabled = False
        FrmABMDet3.Enabled = False
        Fra_datos.Enabled = False
    
        Select Case dtc_codigo2.Text
            Case "1"    'SOLO COMPRAS BB y SS
            Case "2"    'SOLO VENTA DE BIENES
            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'                Call ABRIR_TABLA_DET
'                mw_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'                mw_solicitud_edificacion.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
'                'mw_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'                'mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'                mw_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
'                mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'                mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'
'                mw_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
'                mw_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
'                mw_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
'                mw_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
'                mw_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
'                mw_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
'                mw_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
'                mw_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
'                mw_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
'                mw_solicitud_edificacion.Txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
'                mw_solicitud_edificacion.Txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
'                mw_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
'
'                mw_solicitud_edificacion.Show vbModal
            Case "COM-03"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
                Call ABRIR_TABLA_DET
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                'frm_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle2.Recordset("bitacora_codigo")
                'frm_solicitud_bienes.Txt_estado.Caption = "REG"
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes.dtc_desc1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.dtc_aux1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Dtc_aux2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.dtc_aux3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo4.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                
                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
                
            Case "TEC-05"    '5. SERVICIO MODERNIZACION
            Case "TEC-01"    '6. SERVICIO DE MANTENIMIENTO GRATUITO
                Call ABRIR_TABLA_DET
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                'frm_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle2.Recordset("bitacora_codigo")
                'frm_solicitud_bienes.Txt_estado.Caption = "REG"
                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                frm_solicitud_bienes.dtc_desc1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.dtc_aux1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Dtc_aux2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.dtc_aux3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                frm_solicitud_bienes.Txt_campo4.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
                
                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
            Case "TEC-02"    '10. VENTA DE SERVICIO DE MANTENIMIENTO PREVENTIVO
                'Call ABRIR_TABLA_DET
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                
                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                frm_solicitud_bienes.Txt_campo19.Text = Me.Ado_detalle2.Recordset("bien_cantidad_por_empaque")
                
                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
                frm_solicitud_bienes.Txt_campo15.Text = "10" 'Me.Ado_detalle2.Recordset("fosa_dimension_frente")
                
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
            Case "TEC-03"    '7. VENTA DE SERVICIOS REPARACION
                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
                
                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle2.Recordset("bien_codigo")
                    
                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle2.Recordset("bien_descripcion")
                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle2.Recordset("bien_descripcion_anterior")
                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle2.Recordset("marca_codigo")
                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle2.Recordset("modelo_codigo")
                
                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle2.Recordset("bien_cantidad")
                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle2.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle2.Recordset("bien_total_venta")
                
                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
    '            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle2.Recordset("unimed_codigo")
    '            frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
                frm_solicitud_bienes.lbl_det.Caption = "43340"
                frm_solicitud_bienes.Show vbModal
            
        End Select
        swnuevo = 0
        fraOpciones.Enabled = True
        FraNavega.Enabled = True
        FraDet2.Enabled = True
        FrmABMDet2.Enabled = True
        FraDet3.Enabled = True
        FrmABMDet3.Enabled = True
    '    Fra_datos.Enabled = True
        Call ABRIR_TABLA_DET
        Ado_detalle2.Recordset.Move marca1 - 1
      Else
        MsgBox "No se puede MODIFICAR, porque ya está APROBADO o ANULADO, Verifique por favor!! ", vbExclamation
      End If
  Else
     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub ModifDetalle()
  
  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
    swnuevo = 2
    fraOpciones.Enabled = False
    FraNavega.Enabled = False
    FraDet2.Enabled = False
    FrmABMDet2.Enabled = False
    FraDet3.Enabled = False
    FrmABMDet3.Enabled = False
    Fra_datos.Enabled = False

'    Select Case dtc_codigo2.Text
'        Case "COM-01"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'            Call ABRIR_TABLA_DET
'            mw_solicitud_edificacion.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'            mw_solicitud_edificacion.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'            mw_solicitud_edificacion.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            'mw_solicitud_edificacion.Txt_Correl.Caption = Me.Ado_detalle1.Recordset("bitacora_codigo")
'            'mw_solicitud_edificacion.Txt_estado.Caption = "REG"
'            mw_solicitud_edificacion.dtc_codigo1.Text = Me.Ado_detalle1.Recordset("edif_codigo")
'            mw_solicitud_edificacion.dtc_desc1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux1.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux2.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'            mw_solicitud_edificacion.dtc_aux3.BoundText = mw_solicitud_edificacion.dtc_codigo1.BoundText
'
'            mw_solicitud_edificacion.Txt_campo2.Text = Me.Ado_detalle1.Recordset("edif_area_total_m2")
'            mw_solicitud_edificacion.Txt_campo3.Text = Me.Ado_detalle1.Recordset("edif_area_util_m2")
'            mw_solicitud_edificacion.Txt_campo4.Text = Me.Ado_detalle1.Recordset("edif_num_pisos")
'            mw_solicitud_edificacion.Txt_campo5.Text = Me.Ado_detalle1.Recordset("edif_num_salas_may_200m")
'            mw_solicitud_edificacion.Txt_campo6.Text = Me.Ado_detalle1.Recordset("edif_num_salas_men_200m")
'            mw_solicitud_edificacion.Txt_campo7.Text = Me.Ado_detalle1.Recordset("edif_num_habit_libres")
'            mw_solicitud_edificacion.Txt_campo8.Text = Me.Ado_detalle1.Recordset("edif_num_habit_ocupadas")
'            mw_solicitud_edificacion.Txt_campo9.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_2")
'            mw_solicitud_edificacion.Txt_campo10.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_3")
'            mw_solicitud_edificacion.Txt_campo11.Text = Me.Ado_detalle1.Recordset("edif_num_habit_dorm_4")
'            mw_solicitud_edificacion.Txt_campo12.Caption = Me.Ado_detalle1.Recordset("edif_indicador_min_trafico")
'            mw_solicitud_edificacion.Txt_campo13.Caption = Me.Ado_detalle1.Recordset("edif_capacidad_min_trafico")
'
'            mw_solicitud_edificacion.Show vbModal
'        Case "COM-03"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Call ABRIR_TABLA_DET
'            frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
'            frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
'            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            'frm_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle3.Recordset("bitacora_codigo")
'            'frm_solicitud_bienes.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'            frm_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'            frm_solicitud_bienes.dtc_desc1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo4.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'
'            frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle3.Recordset("bien_descripcion")
'            frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle3.Recordset("bien_descripcion_anterior")
'            frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle3.Recordset("marca_codigo")
'            frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle3.Recordset("modelo_codigo")
'
'            frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle3.Recordset("bien_cantidad")
'            frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle3.Recordset("bien_precio_venta_base")
'            frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle3.Recordset("bien_total_venta")
'            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
'            frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
'            frm_solicitud_bienes.lbl_det.Caption = "34110"
'            frm_solicitud_bienes.Show vbModal
'
'        Case "TEC-05"    '5. SERVICIO MODERNIZACION
'        Case "TEC-01"    '6. SERVICIO DE MANTENIMIENTO GRATUITO
''            Call ABRIR_TABLA_DET
'            frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
'            frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
'            frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            'frm_solicitud_bienes.Txt_Correl.Caption = Me.Ado_detalle3.Recordset("bitacora_codigo")
'            'frm_solicitud_bienes.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'            frm_solicitud_bienes.dtc_codigo1.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'            frm_solicitud_bienes.dtc_desc1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.dtc_aux3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'            frm_solicitud_bienes.Txt_campo4.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'
'            frm_solicitud_bienes.Txt_campo6.Text = Me.Ado_detalle3.Recordset("bien_descripcion")
'            frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle3.Recordset("bien_descripcion_anterior")
'            frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle3.Recordset("marca_codigo")
'            frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle3.Recordset("modelo_codigo")
'
'            frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle3.Recordset("bien_cantidad")
'            frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle3.Recordset("bien_precio_venta_base")
'            frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle3.Recordset("bien_total_venta")
'            frm_solicitud_bienes.dtc_codigo2.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
'            frm_solicitud_bienes.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
'            frm_solicitud_bienes.lbl_det.Caption = "34110"
'            frm_solicitud_bienes.Show vbModal
'        Case "TEC-02"    '10. VENTA DE SERVICIO DE MANTENIMIENTO PREVENTIVO
'            If VAR_DET = "30000" Then
'                marca1 = Ado_detalle3.Recordset.Bookmark
'                Ado_detalle3.Recordset.AddNew
'                'Call ABRIR_TABLA_DET
'                frm_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes3.Txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'    '            frm_solicitud_bienes3.dtc_codigo1.Text = Me.Ado_detalle3.Recordset("bien_codigo")
'    '            frm_solicitud_bienes3.dtc_desc1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.dtc_aux1.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.dtc_aux2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.dtc_aux3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.Txt_campo2.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.txt_campo3.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'    '            frm_solicitud_bienes3.txt_campo4.BoundText = frm_solicitud_bienes.dtc_codigo1.BoundText
'
'                frm_solicitud_bienes3.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
'                frm_solicitud_bienes3.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
'                frm_solicitud_bienes3.Txt_campo8.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!marca_codigo), "S/M", Me.Ado_detalle3.Recordset!marca_codigo)
'                frm_solicitud_bienes3.Txt_campo9.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!modelo_codigo), "S/M", Me.Ado_detalle3.Recordset!modelo_codigo)
'
'                frm_solicitud_bienes3.Txt_campo16.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_cantidad), "1", Me.Ado_detalle3.Recordset!bien_cantidad)
'                frm_solicitud_bienes3.Txt_campo10.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_precio_venta_base), "0", Me.Ado_detalle3.Recordset!bien_precio_venta_base) '
'                frm_solicitud_bienes3.Txt_campo11.Caption = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_total_venta), "0", Me.Ado_detalle3.Recordset!bien_total_venta)
'    '            frm_solicitud_bienes3.dtc_codigo2.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
'    '            frm_solicitud_bienes3.dtc_desc2.BoundText = frm_solicitud_bienes.dtc_codigo2.BoundText
'                frm_solicitud_bienes3.lbl_det.Caption = VAR_DET
'                frm_solicitud_bienes3.Show vbModal
'                Ado_detalle3.Recordset.Move marca1 - 1
'            End If
'            If VAR_DET = "39800" Then
'
'
'                frm_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle5.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle5.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes3.Txt_campo5.Text = Me.Ado_detalle5.Recordset("bien_codigo")
'
'                frm_solicitud_bienes3.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
'                frm_solicitud_bienes3.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
'                frm_solicitud_bienes3.Txt_campo8.Text = Me.Ado_detalle5.Recordset("marca_codigo")
'                frm_solicitud_bienes3.Txt_campo9.Text = Me.Ado_detalle5.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes3.Txt_campo16.Text = Me.Ado_detalle5.Recordset("bien_cantidad")
'                frm_solicitud_bienes3.Txt_campo10.Text = Me.Ado_detalle5.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes3.Txt_campo11.Caption = Me.Ado_detalle5.Recordset("bien_total_venta")
'                frm_solicitud_bienes3.lbl_det.Caption = VAR_DET
'                frm_solicitud_bienes3.Show vbModal
'            End If
'            If VAR_DET = "34800" Then
'                frm_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle6.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle6.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes3.Txt_campo5.Text = Me.Ado_detalle6.Recordset("bien_codigo")
'
'                frm_solicitud_bienes3.Txt_campo6.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
'                frm_solicitud_bienes3.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
'                frm_solicitud_bienes3.Txt_campo8.Text = Me.Ado_detalle6.Recordset("marca_codigo")
'                frm_solicitud_bienes3.Txt_campo9.Text = Me.Ado_detalle6.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes3.Txt_campo16.Text = Me.Ado_detalle6.Recordset("bien_cantidad")
'                frm_solicitud_bienes3.Txt_campo10.Text = Me.Ado_detalle6.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes3.Txt_campo11.Caption = Me.Ado_detalle6.Recordset("bien_total_venta")
'                frm_solicitud_bienes3.lbl_det.Caption = VAR_DET
'                frm_solicitud_bienes3.Show vbModal
'
'            End If
'        Case "TEC-03"    '11. VENTA DE SERVICIO DE MANTENIMIENTO CORRECTIVO
            If VAR_DET = "30000" Then
                'marca1 = Ado_detalle3.Recordset.Bookmark
                frm_solicitud_bienes3.txt_codigo.Caption = Me.Ado_detalle3.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes3.Txt_campo1.Caption = Me.Ado_detalle3.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes3.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                frm_solicitud_bienes3.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes3.Txt_campo5.Text = Me.Ado_detalle3.Recordset("bien_codigo")
                
                frm_solicitud_bienes3.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
                frm_solicitud_bienes3.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                frm_solicitud_bienes3.Txt_campo8.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!marca_codigo), "S/M", Me.Ado_detalle3.Recordset!marca_codigo)
                frm_solicitud_bienes3.Txt_campo9.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!modelo_codigo), "S/M", Me.Ado_detalle3.Recordset!modelo_codigo)
                
                frm_solicitud_bienes3.Txt_campo16.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_cantidad), "1", Me.Ado_detalle3.Recordset!bien_cantidad)
                frm_solicitud_bienes3.Txt_campo10.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_precio_venta_base), "0", Me.Ado_detalle3.Recordset!bien_precio_venta_base)
                frm_solicitud_bienes3.Txt_campo11.Caption = IIf(IsNull(Me.Ado_detalle3.Recordset!bien_total_venta), "0", Me.Ado_detalle3.Recordset!bien_total_venta)
    
                frm_solicitud_bienes3.Txt_campo14.Text = Me.Ado_detalle3.Recordset("unimed_codigo")
                frm_solicitud_bienes3.Txt_campo15.Text = Me.Ado_detalle3.Recordset("fosa_dimension_frente")

                frm_solicitud_bienes3.lbl_det.Caption = VAR_DET
                frm_solicitud_bienes3.Show vbModal
                'Ado_detalle3.Recordset.Move marca1 - 1
            End If
            If VAR_DET = "39800" Then
                frm_solicitud_bienes5.txt_codigo.Caption = Me.Ado_detalle5.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes5.Txt_campo1.Caption = Me.Ado_detalle5.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes5.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                frm_solicitud_bienes5.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes5.Txt_campo5.Text = Me.Ado_detalle5.Recordset("bien_codigo")
                
                frm_solicitud_bienes5.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion), "-", Me.Ado_detalle5.Recordset!bien_descripcion)
                frm_solicitud_bienes5.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle5.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle5.Recordset!bien_descripcion_anterior)
                frm_solicitud_bienes5.Txt_campo8.Text = Me.Ado_detalle5.Recordset("marca_codigo")
                frm_solicitud_bienes5.Txt_campo9.Text = Me.Ado_detalle5.Recordset("modelo_codigo")
                
                frm_solicitud_bienes5.Txt_campo16.Text = Me.Ado_detalle5.Recordset("bien_cantidad")
                frm_solicitud_bienes5.Txt_campo10.Text = Me.Ado_detalle5.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes5.Txt_campo11.Caption = Me.Ado_detalle5.Recordset("bien_total_venta")
                
                frm_solicitud_bienes5.Txt_campo14.Text = Me.Ado_detalle5.Recordset("unimed_codigo")
                frm_solicitud_bienes5.Txt_campo15.Text = Me.Ado_detalle5.Recordset("fosa_dimension_frente")
                
                frm_solicitud_bienes5.lbl_det.Caption = VAR_DET
                frm_solicitud_bienes5.Show vbModal
            End If
            If VAR_DET = "34800" Then
                frm_solicitud_bienes6.txt_codigo.Caption = Me.Ado_detalle6.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes6.Txt_campo1.Caption = Me.Ado_detalle6.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes6.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                frm_solicitud_bienes6.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes6.Txt_campo5.Text = Me.Ado_detalle6.Recordset("bien_codigo")
                
                frm_solicitud_bienes6.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion), "-", Me.Ado_detalle3.Recordset!bien_descripcion)
                frm_solicitud_bienes6.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle6.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle3.Recordset!bien_descripcion_anterior)
                frm_solicitud_bienes6.Txt_campo8.Text = Me.Ado_detalle6.Recordset("marca_codigo")
                frm_solicitud_bienes6.Txt_campo9.Text = Me.Ado_detalle6.Recordset("modelo_codigo")
                
                frm_solicitud_bienes6.Txt_campo16.Text = Me.Ado_detalle6.Recordset("bien_cantidad")
                frm_solicitud_bienes6.Txt_campo10.Text = Me.Ado_detalle6.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes6.Txt_campo11.Caption = Me.Ado_detalle6.Recordset("bien_total_venta")
                
                frm_solicitud_bienes6.Txt_campo14.Text = Me.Ado_detalle6.Recordset("unimed_codigo")
                frm_solicitud_bienes6.Txt_campo15.Text = Me.Ado_detalle6.Recordset("fosa_dimension_frente")
                
                frm_solicitud_bienes6.lbl_det.Caption = VAR_DET
                frm_solicitud_bienes6.Show vbModal
            End If

            If VAR_DET = "24300" Then
                frm_solicitud_bienes7.txt_codigo.Caption = Me.Ado_detalle7.Recordset("solicitud_codigo")  'cod_cabecera
                frm_solicitud_bienes7.Txt_campo1.Caption = Me.Ado_detalle7.Recordset("unidad_codigo")  'Unidad
                frm_solicitud_bienes7.Txt_descripcion.Caption = Me.dtc_desc1.Text
            
                frm_solicitud_bienes7.lbl_edif.Caption = dtc_codigo3.Text
                frm_solicitud_bienes7.Txt_campo5.Text = Me.Ado_detalle7.Recordset("bien_codigo")
                
                frm_solicitud_bienes7.txt_campo6.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion), "-", Me.Ado_detalle7.Recordset!bien_descripcion)
                frm_solicitud_bienes7.Txt_campo7.Text = IIf(IsNull(Me.Ado_detalle7.Recordset!bien_descripcion_anterior), "-", Me.Ado_detalle7.Recordset!bien_descripcion_anterior)
                frm_solicitud_bienes7.Txt_campo8.Text = Me.Ado_detalle7.Recordset("marca_codigo")
                frm_solicitud_bienes7.Txt_campo9.Text = Me.Ado_detalle7.Recordset("modelo_codigo")
                
                frm_solicitud_bienes7.Txt_campo16.Text = Me.Ado_detalle7.Recordset("bien_cantidad")
                frm_solicitud_bienes7.Txt_campo10.Text = Me.Ado_detalle7.Recordset("bien_precio_venta_base")
                frm_solicitud_bienes7.Txt_campo11.Caption = Me.Ado_detalle7.Recordset("bien_total_venta")
                
                frm_solicitud_bienes7.Txt_campo14.Text = Me.Ado_detalle7.Recordset("unimed_codigo")
                frm_solicitud_bienes7.Txt_campo15.Text = Me.Ado_detalle7.Recordset("fosa_dimension_frente")
                
                frm_solicitud_bienes7.lbl_det.Caption = VAR_DET
                frm_solicitud_bienes7.Show vbModal
            End If
'    End Select
    
    swnuevo = 0
    fraOpciones.Enabled = True
    FraNavega.Enabled = True
    FraDet2.Enabled = True
    FrmABMDet2.Enabled = True
    FraDet3.Enabled = True
    FrmABMDet3.Enabled = True
'    Fra_datos.Enabled = True
    Call ABRIR_TABLA_DET
'    Ado_detalle3.Recordset.Move marca1 - 1
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If

End Sub

Private Sub BtnModDetalle3_Click()
   If Ado_detalle3.Recordset.RecordCount > 0 Then
       If Ado_detalle3.Recordset("estado_codigo") = "REG" Then
          VAR_DET = "30000"
          Call ModifDetalle
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle5_Click()
   If Ado_detalle5.Recordset.RecordCount > 0 Then
       If Ado_detalle5.Recordset("estado_codigo") = "REG" Then
          VAR_DET = "39800"
          Call ModifDetalle
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle6_Click()
   If Ado_detalle6.Recordset.RecordCount > 0 Then
       If Ado_detalle6.Recordset("estado_codigo") = "REG" Then
          VAR_DET = "34800"
          Call ModifDetalle
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle7_Click()
   If Ado_detalle7.Recordset.RecordCount > 0 Then
       If Ado_detalle7.Recordset("estado_codigo") = "REG" Then
          VAR_DET = "24300"
          Call ModifDetalle
       Else
            MsgBox "No se puede MODIFICAR, el registro ya está APROBADO o ANULADO, Verifique por favor ...", vbExclamation, "Validación de Registro"
       End If
   Else
     MsgBox "No se puede MODIFICAR, el registro No Existe o No fue identificado correctamente, Verifique por favor ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        marca1 = Ado_datos.Recordset.Bookmark
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "MOD"
    '    dtc_desc1.Visible = False
    '    lbl_aux1.Visible = True
    '    lbl_aux1.Caption = dtc_desc1.Text
        dtc_desc4.SetFocus
    '    BtnVer.Visible = True
'        dtc_codigo9.Enabled = False
        Call OptFilGral1_Click
        Ado_datos.Recordset.Move marca1 - 1
    Select Case parametro
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"            'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Else
        MsgBox "NO se puede MODIFICAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
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
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
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
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    '    db.RollbackTrans
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
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
    dtc_aux3.BoundText = dtc_codigo3.BoundText
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
  
Private Sub pnivel11(codigo2 As String)
    Select Case codigo2
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"            'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
'    Dim strConsultaF As String
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
End Sub

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
    dtc_aux3.BoundText = dtc_desc3.BoundText
End Sub
 
Private Sub dtc_desc3_LostFocus()
    dtc_codigo4.Text = dtc_aux3.Text
    Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc10.Enabled = True
'    Call pnivel11(dtc_codigo1.BoundText)
'    dtc_desc11.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

'Private Sub dtc_desc6_Click(Area As Integer)
'    dtc_codigo6.BoundText = dtc_desc6.BoundText
''    Call pnivel6(dtc_codigo6.BoundText)
''    dtc_desc7.Enabled = True
'End Sub
  

'Private Sub dtc_desc7_Click(Area As Integer)
'    dtc_codigo7.BoundText = dtc_desc7.BoundText
'End Sub

'Private Sub dtc_desc8_Click(Area As Integer)
'    dtc_codigo8.BoundText = dtc_desc8.BoundText
'    Call pnivel8(dtc_codigo8.BoundText)
'    'dtc_desc9.Enabled = True
'    dtc_codigo9.Enabled = True
'End Sub
   
'Private Sub pnivel8(codigo8 As String)
'   Dim strConsultaF As String
'
'   strConsultaF = "select * from gc_documentos_respaldo where clasif_codigo = '" & codigo8 & "'"
'
'   Set dtc_codigo9.RowSource = Nothing
'   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo9.ReFill
'   dtc_codigo9.BoundText = Empty
'
'   Set dtc_desc9.RowSource = Nothing
'   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc9.ReFill
'   dtc_desc9.BoundText = Empty
'End Sub

'Private Sub dtc_desc9_Click(Area As Integer)
'    dtc_codigo9.BoundText = dtc_codigo9.BoundText
'End Sub

Private Sub Form_Load()
    buscados = 0
    swnuevo = 0
    VAR_SW = ""
    parametro = Aux
    'parametro = "estado_codigo" + " = " + "'REG'"
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
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    'gc_tipo_solicitud
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    If parametro = "DNINS" Or parametro = "DNAJS" Then
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'COM' order by subproceso_descripcion", db, adOpenStatic
    Else
        rs_datos2.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'TEC' order by subproceso_descripcion", db, adOpenStatic
    End If
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    'gc_edificaciones
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    'rs_datos3.Open "Select * from gc_edificaciones order by edif_denominacion", db, adOpenStatic
    rs_datos3.Open "gp_listar_apr_gc_edificaciones", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'gc_beneficiario (Personas Nat. y Juridicas / Clientes, Proveedores, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "gp_listar_gc_beneficiario_personas", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
'    Set rs_datos5 = New ADODB.Recordset
'    If rs_datos5.State = 1 Then rs_datos5.Close
'    'rs_datos5.Open "Select * from gc_proceso_nivel1 order by proceso_descripcion", db, adOpenStatic
'    rs_datos5.Open "gp_listar_apr_gc_proceso_nivel1", db, adOpenStatic
'    Set Ado_datos5.Recordset = rs_datos5
''    dtc_desc5.BoundText = dtc_codigo5.BoundText

'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from gc_proceso_nivel2 WHERE proceso_codigo = 'TEC' order by subproceso_descripcion", db, adOpenStatic
'    'rs_datos6.Open "gp_listar_apr_gc_proceso_nivel2", db, adOpenStatic
'    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
'
'    Set rs_datos7 = New ADODB.Recordset
'    If rs_datos7.State = 1 Then rs_datos7.Close
'    'rs_datos7.Open "Select * from gc_proceso_nivel3 order by etapa_descripcion", db, adOpenStatic
'    rs_datos7.Open "gp_listar_apr_gc_proceso_nivel3", db, adOpenStatic
'    Set Ado_datos7.Recordset = rs_datos7
'    dtc_desc7.BoundText = dtc_codigo7.BoundText
'
'    Set rs_datos8 = New ADODB.Recordset
'    If rs_datos8.State = 1 Then rs_datos8.Close
'    'rs_datos8.Open "Select * from gc_documentos_clasificacion order by clasif_codigo", db, adOpenStatic
'    rs_datos8.Open "gp_listar_apr_gc_documentos_clasificacion", db, adOpenStatic
'    Set Ado_datos8.Recordset = rs_datos8
''    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
'    'gc_documentos_respaldo
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos9.Close
'    'rs_datos9.Open "Select * from gc_documentos_respaldo order by doc_codigo", db, adOpenStatic
'    rs_datos9.Open "gp_listar_apr_gc_documentos_respaldo", db, adOpenStatic
'    Set Ado_datos9.Recordset = rs_datos9
'    dtc_desc9.BoundText = dtc_codigo9.BoundText
    
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
'    'queryinicial = "select solicitud_codigo, unidad_codigo, solicitud_justificacion, solicitud_observaciones, estado_codigo, fecha_registro, usr_codigo, hora_registro, ges_gestion, solicitud_fecha_solicitud as fecha1,  solicitud_fecha_recepción as fecha2, solicitud_tipo as codigo2, beneficiario_codigo as codigo4, beneficiario_codigo_resp as codigo11, edif_codigo as codigo3, proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero As campo1, poa_codigo As codigo10, archivo_respaldo, archivo_respaldo_cargado, ges_gestion_ant, unidad_codigo_ant, solicitud_codigo_ant, usr_codigo_aprueba, fecha_aprueba, hora_aprueba From ao_solicitud WHERE estado_codigo = 'REG' "
'    queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub ABRIR_TABLA_DET()
    'BITACORA
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    'rs_det1.Open "select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det1.Open "select * from ao_solicitud_bitacora where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    If rs_det1.RecordCount > 0 Then
        dg_det1.Visible = True
        Set dg_det1.DataSource = Ado_detalle1.Recordset
    Else
        dg_det1.Visible = False
        'Set Ado_detalle1.Recordset = rsNada
        Set dg_det1.DataSource = rsNada
    End If
    
    'EQUIPOS par_codigo = '43340'
    Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    'rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det2.Open "select * from av_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and (par_codigo = '43340' ) ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and estado_codigo = 'APR'
    Set Ado_detalle2.Recordset = rs_det2
    If rs_det2.RecordCount > 0 Then
        dg_det2.Visible = True
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det2.Visible = False
        'Set Ado_detalle2.Recordset = rsNada
        Set dg_det2.DataSource = rsNada
    End If
    
    'INSUMOS y materiales par_codigo = '43340'
    Set rs_det3 = New Recordset
    If rs_det3.State = 1 Then rs_det3.Close
    rs_det3.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (grupo_codigo = '30000' and (par_codigo <> '39810' and par_codigo <> '39820' and par_codigo <> '34800'))   ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    Set Ado_detalle3.Recordset = rs_det3.DataSource
    If rs_det3.RecordCount > 0 Then
        dg_det3.Visible = True
        Set dg_det3.DataSource = Ado_detalle3.Recordset
    Else
        dg_det3.Visible = False
        Set dg_det3.DataSource = rsNada
    End If

    'REPUESTOS par_codigo = '39800'
    Set rs_det5 = New Recordset
    If rs_det5.State = 1 Then rs_det5.Close
    rs_det5.Open "select * from av_solicitud_bienes3 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '39810' or par_codigo = '39820')  ", db, adOpenKeyset, adLockOptimistic, adCmdText        'and estado_codigo = 'APR'
    Set Ado_detalle5.Recordset = rs_det5.DataSource
    If rs_det5.RecordCount > 0 Then
        dg_det5.Visible = True
        Set dg_det5.DataSource = Ado_detalle5.Recordset
    Else
        dg_det5.Visible = False
        Set dg_det5.DataSource = rsNada
    End If

    'HERRAMIENTAS par_codigo = '43700' - par_codigo = '34800'
    Set rs_det6 = New Recordset
    If rs_det6.State = 1 Then rs_det6.Close
    rs_det6.Open "select * from av_solicitud_bienes2 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  and (par_codigo = '43700' or par_codigo = '34800')  ", db, adOpenKeyset, adLockOptimistic, adCmdText     'and estado_codigo = 'APR'
    Set Ado_detalle6.Recordset = rs_det6.DataSource
    If rs_det6.RecordCount > 0 Then
        dg_det6.Visible = True
        Set dg_det6.DataSource = Ado_detalle6.Recordset
    Else
        dg_det6.Visible = False
        Set dg_det6.DataSource = rsNada
    End If
    
    Set rs_det4 = New Recordset
    If rs_det4.State = 1 Then rs_det4.Close
    'rs_det3.Open "select * from av_solicitud_bienes where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenKeyset, adLockOptimistic, adCmdText
    rs_det4.Open "select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle4.Recordset = rs_det4.DataSource
    Set dg_det4.DataSource = Ado_detalle4.Recordset
    
    'REPUESTOS par_codigo = '24000'
    Set rs_det7 = New Recordset
    If rs_det7.State = 1 Then rs_det7.Close
    rs_det7.Open "select * from av_solicitud_bienes7 where unidad_codigo = '" & parametro & "' and solicitud_codigo = " & VAR_SOL & " and par_codigo = '24300'   ", db, adOpenKeyset, adLockOptimistic, adCmdText      '
    Set Ado_detalle7.Recordset = rs_det7.DataSource
    If rs_det7.RecordCount > 0 Then
        dg_det7.Visible = True
        Set dg_det7.DataSource = Ado_detalle7.Recordset
    Else
        dg_det7.Visible = False
        Set dg_det7.DataSource = rsNada
    End If
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
  'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
    'Ado_datos.Caption = Ado_datos.Recordset.AbsolutePosition & " / " & Ado_datos.Recordset.RecordCount
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
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
        
        'Select Case rs_datos!solicitud_tipo     'dtc_codigo2.Text
'        If VAR_SOL = 0 Then
        'If VAR_SW <> "" Then
        If Not (Ado_datos.Recordset.EOF) Then   'And Not (Ado_datos.Recordset.BOF)
            VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        End If
        Call ABRIR_TABLA_DET
'        Select Case Ado_datos.Recordset!subproceso_codigo     'dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL - PROVISION E IMPORTACION DE EQUIPOS
'                Call ABRIR_TABLA_DET
'            Case "COM-03"    '4. VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'                Call ABRIR_TABLA_DET
'            Case "TEC-05"    '5. SERVICIO MODERNIZACION
'
'            Case "TEC-01"    '6. SERVICIO DE MANTENIMIENTO GRATUITO
'                Call ABRIR_TABLA_DET
'            Case "TEC-02"    '10. VENTA DE SERVICIOS (MANTENIMIENTO PREVENTIVO)
'                Call ABRIR_TABLA_DET
'
'            Case Else
'
'        End Select
        'VAR_SOL = Ado_datos.Recordset!solicitud_codigo
        Call ABRIR_TABLA_AUX2
    Else
        'Set rs_det1 = New ADODB.Recordset
        'Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
    FraDet1.Caption = "BITÁCORA DE " + dtc_desc2.Text
    'FraDet1.Caption = "BITÁCORA DE " + lbl_titulo
'    txt_aux9.Text = dtc_desc9.Text
    If Not (Ado_datos.Recordset.EOF) Then
        If Ado_datos.Recordset!estado_codigo = "APR" Then
            FrmABMDet2.Visible = False
            FrmABMDet3.Visible = False
        Else
            FrmABMDet2.Visible = True
            FrmABMDet3.Visible = True
        End If
    End If
  Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        Set dg_det2.DataSource = rsNada
        Set dg_det3.DataSource = rsNada
        Set dg_det5.DataSource = rsNada
        Set dg_det6.DataSource = rsNada
        Set dg_det7.DataSource = rsNada
     If buscados = 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
     Else
        OptFilGral1.Visible = False
        OptFilGral2.Visible = False
     End If
  End If
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
        Case "DVTA"             'INI COMERCIAL
            dtc_codigo2.Text = "COM-01"   '3
        Case "COMEX"            'INI COMEX
            dtc_codigo2.Text = "CMX-01"   '3
        Case "DNINS"            'INI GRABA INSTALACIONES
            dtc_codigo2.Text = "COM-03" '4
        Case "DNAJS"            'AJUSTE
            dtc_codigo2.Text = "COM-04" '5
        Case "DNMAN"            'MANTENIMIENTO PREVENTIVO
            dtc_codigo2.Text = "TEC-02" '10
        Case "DNREP"            'MANTENIMIENTO CORRECTIVO / REPARACIONES
            dtc_codigo2.Text = "TEC-03" '10
        Case "DNEME"            'EMERGENCIAS
            dtc_codigo2.Text = "TEC-04" '10
        Case "DNMOD"            'MODERNIZACION
            dtc_codigo2.Text = "TEC-05" '10
        Case Else
            dtc_codigo2.Text = "TEC-01"   '3
    End Select
    dtc_desc2.BoundText = dtc_codigo2.BoundText
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from ao_solicitud where estado_codigo = 'REG' AND unidad_codigo = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from ao_solicitud where unidad_codigo = '" & parametro & "' "
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

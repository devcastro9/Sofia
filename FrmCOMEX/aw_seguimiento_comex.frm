VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_seguimiento_comex 
   BackColor       =   &H00000000&
   Caption         =   "Procesos Administrativos - COMEX - Compra Servicios"
   ClientHeight    =   10260
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   16545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   16545
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   24
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   18840
         Picture         =   "aw_seguimiento_comex.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   34
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5520
         Picture         =   "aw_seguimiento_comex.frx":07C2
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   32
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4200
         Picture         =   "aw_seguimiento_comex.frx":108F
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   31
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
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   33
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
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   22
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5160
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   21
         Top             =   0
         Width           =   1335
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
         TabIndex        =   23
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.TextBox txtcodigopago 
      Height          =   285
      Left            =   4440
      TabIndex        =   18
      Top             =   9720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtcodigoadj 
      Height          =   285
      Left            =   5400
      TabIndex        =   17
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtprocesocomex 
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00000000&
      Caption         =   "EQUIPOS A IMPORTAR"
      ForeColor       =   &H00FFFFFF&
      Height          =   1605
      Left            =   105
      TabIndex        =   13
      Top             =   4860
      Width           =   18735
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   1305
         Left            =   75
         TabIndex        =   14
         Top             =   225
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   2302
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "grupo_codigo"
            Caption         =   "Grupo"
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
            DataField       =   "subgrupo_codigo"
            Caption         =   "Sub-Grupo"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Bien/Serv"
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
            DataField       =   "compra_concepto"
            Caption         =   "Descripcion Bien o Servicio"
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
            DataField       =   "compra_cantidad"
            Caption         =   "Cantidad"
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
            DataField       =   "compra_precio_unitario_bs"
            Caption         =   "Precio.BOB"
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
            DataField       =   "compra_precio_total_bs"
            Caption         =   "Precio.Total"
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
            DataField       =   "compra_precio_unitario_dol"
            Caption         =   "Precio.USD"
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
            DataField       =   "compra_precio_total_dol"
            Caption         =   "Total USD"
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
            DataField       =   "tipo_eqp_descripcion"
            Caption         =   "Tipo.Bien/Serv."
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
            DataField       =   "marca_descripcion"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00000000&
      Caption         =   "CRONOGRAMA - ORDEN DE PAGO"
      ForeColor       =   &H00FFFFFF&
      Height          =   1480
      Left            =   120
      TabIndex        =   11
      Top             =   8115
      Width           =   18735
      Begin MSDataGridLib.DataGrid dg_det3 
         Height          =   1200
         Left            =   75
         TabIndex        =   12
         Top             =   225
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   2117
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
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "pago_codigo"
            Caption         =   "Nro.O.P."
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
            DataField       =   "pago_descripcion"
            Caption         =   "Concepto.de.la. Orden de Pago"
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
            DataField       =   "pago_fecha_prog"
            Caption         =   "Fecha.O.P."
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
            DataField       =   "fecha_pago"
            Caption         =   "Fecha Pago"
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
            DataField       =   "tipo_transaccion"
            Caption         =   "Tipo Transaccion"
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
            DataField       =   "cta_codigo"
            Caption         =   "Cta. Bancaria"
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
            DataField       =   "concepto_pago"
            Caption         =   "Concepto Pago"
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
            DataField       =   "pago_total_bs"
            Caption         =   "Monto.a.Pagar.Bs"
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
         BeginProperty Column08 
            DataField       =   "pago_total_dol"
            Caption         =   "Monto.a.Pagar.USD"
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
            DataField       =   "numero_cheque_trf"
            Caption         =   "Nro. Documento"
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
            DataField       =   "cta_codigo_destino"
            Caption         =   "Cta. Bancaria Destino"
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
            DataField       =   "beneficiario_destino"
            Caption         =   "Beneficiario Destino"
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
            DataField       =   "tipo_cambio"
            Caption         =   "Tipo Cambio"
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
         BeginProperty Column14 
            DataField       =   "compra_codigo"
            Caption         =   "Nro.Compra"
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
            DataField       =   "adjudica_codigo"
            Caption         =   "Nro.Adjudica"
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
            DataField       =   "pago_emite_factura"
            Caption         =   "Emite.Factura"
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
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00000000&
      Caption         =   "ADJUDICACION"
      ForeColor       =   &H00FFFFFF&
      Height          =   1480
      Left            =   120
      TabIndex        =   7
      Top             =   6495
      Width           =   18735
      Begin MSDataGridLib.DataGrid dg_det2 
         Height          =   1200
         Left            =   75
         TabIndex        =   8
         Top             =   225
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   2117
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "tipo"
            Caption         =   "Tipo"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "Cod.Proveedor"
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
            DataField       =   "adjudica_descripcion"
            Caption         =   "Denominación.Proveedor"
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
            DataField       =   "adjudica_monto_dol"
            Caption         =   "Precio.Total.USD"
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
         BeginProperty Column04 
            DataField       =   "adjudica_monto_bs"
            Caption         =   "Precio.Total_BOB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4105
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "adjudica_cantidad_total"
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
         BeginProperty Column06 
            DataField       =   "fecha_inicio_contrato"
            Caption         =   "Fecha.Inicio"
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
            DataField       =   "fecha_fin_contrato"
            Caption         =   "Fecha.Finalizacion"
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
            DataField       =   "Fecha_envio_proveedor"
            Caption         =   "Fecha.Entrega/Salida"
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
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column08 
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
      TabIndex        =   6
      Top             =   720
      Width           =   18735
      Begin MSDataGridLib.DataGrid dg_datos 
         Height          =   3330
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14835
         _ExtentX        =   26167
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
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "unidad_codigo_ant"
            Caption         =   "Cite.Trámite"
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
            DataField       =   "compra_fecha"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "nombre_edificio"
            Caption         =   "Nombre Edificio"
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
            DataField       =   "compra_codigo"
            Caption         =   "Correlativo"
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
            DataField       =   "proveedor"
            Caption         =   "Fabricacion"
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
            DataField       =   "transporte"
            Caption         =   "Transporte"
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
            DataField       =   "aduanisacion"
            Caption         =   "Aduanizacion"
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
            DataField       =   "descarguio"
            Caption         =   "Descarguio"
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
            DataField       =   "tecnico"
            Caption         =   "Tecnico"
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
            DataField       =   "nombre_cliente"
            Caption         =   "Cliente"
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
            DataField       =   "estado_cabecera"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
               Alignment       =   2
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
         Left            =   5880
         TabIndex        =   10
         Top             =   3720
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3600
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "                                  Registros Principales"
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
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   16545
      TabIndex        =   0
      Top             =   10260
      Width           =   16545
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   0
         Width           =   3360
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
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
      Left            =   2280
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
      Top             =   9960
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
      Top             =   9960
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
      Top             =   9960
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
      Top             =   9960
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
      Top             =   9960
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
      Top             =   10320
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
      Left            =   120
      Top             =   10680
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
      Left            =   2400
      Top             =   10680
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
      Top             =   10320
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
      Left            =   9000
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Ado_detalle3 
      Height          =   330
      Left            =   4800
      Top             =   10680
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "36NO"
      DataField       =   "compra_codigo"
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
      Left            =   12000
      TabIndex        =   19
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblcodigopagosig 
      Caption         =   "1"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   9720
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "aw_seguimiento_comex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Inicio Variables Globales
Dim rs_datos As New ADODB.Recordset
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

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rsNada As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod As String
Dim VAR_VAL As String
Dim VAR_SW As String
Dim NombreCarpeta, e As String
Dim CodBien As String
Dim VAR_UNI As String
Dim sino As String
Dim VAR_PAIS As String

Dim VAR_CMPBTE As Integer

Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_FOBSEG, VAR_FOBSEG2 As Double

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean
Dim numAdjudicacion As Integer
Dim codProcesoComex As String
Dim numAdjudicacionCargado As Integer

' Fin Variables Globales

'Private Sub Ado_datos2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
' If numAdjudicacion > 0 Then
'    Call CargarGridDetalles
' End If
'End Sub

Private Sub Ado_detalle2_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'If numAdjudicacion > 0 Then
    If Ado_detalle2.Recordset.BOF = False Then
      numAdjudicacion = Ado_detalle2.Recordset!adjudica_codigo
       If numAdjudicacion <> numAdjudicacionCargado And numAdjudicacion > 0 Then
         'Call CargarGridDetalles
         Call CargarGridDetallesOrden
       End If
    End If
 'End If
End Sub


' Boton anular orden pago.
Private Sub btnAnular_Click()
   If Ado_detalle3.Recordset.BOF = False Then
          If Ado_detalle3.Recordset!estado_codigo = "REG" Then
                Dim sqlA As String
                
                sqlA = " UPDATE ao_compra_planilla_pagos SET estado_codigo = 'ANL' WHERE ges_gestion = '" & CStr(Ado_detalle3.Recordset!ges_gestion) & "' AND compra_codigo = " & CStr(Ado_detalle3.Recordset!compra_codigo) & " AND adjudica_codigo = '" & CStr(Ado_detalle3.Recordset!adjudica_codigo) & "' AND pago_codigo = '" & CStr(Ado_detalle3.Recordset!pago_codigo) & "' "
                db.Execute sqlA
                
                MsgBox "El registro se anulo."
                Call CargarGridDetallesOrden
          Else
              MsgBox "El registro no se puede anular por que el estado es diferente de REG."
          End If
        Else
          MsgBox "Seleccione una orden de pago"
        End If
End Sub

' ============ 1 COMANDOS =================================
' Boton Nuevo
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
            '
            dtc_codigo2.Text = 4
        Case Else
            dtc_codigo2.Text = 5
    End Select
    dtc_desc2.BoundText = dtc_codigo2.BoundText

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

' Boton guardar
Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If VAR_SW = "ADD" Then
        VAR_UNI = dtc_codigo1.Text
        var_cod = IIf(txt_codigo.Caption = "", 0, txt_codigo.Caption)
        Set rs_aux1 = New ADODB.Recordset
        SQL_FOR = "select * from ao_compras_cabecera where unidad_codigo = '" & VAR_UNI & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            var_cod = rs_aux1.RecordCount + 1
        
        Else
            'var_cod = rs_datos.RecordCount '+ 1
            var_cod = 1
        End If
        
        txt_codigo.Caption = var_cod
        rs_datos!solicitud_codigo = var_cod
        rs_datos!estado_codigo = "REG"      'no cambia
        rs_datos!ges_gestion = glGestion    ' Year(Date)   'no cambia
        rs_datos!unidad_codigo = VAR_UNI
        'Actualiza correaltivo ...
        db.Execute "Update gc_unidad_ejecutora Set correl_solicitud = " & var_cod & " Where unidad_codigo = '" & VAR_UNI & "'   "
        rs_datos!doc_numero = "0"    'txt_campo1.Caption
        rs_datos!archivo_respaldo = "sin_nombre"
        rs_datos!archivo_respaldo_cargado = "N"
        rs_datos!correl_bitacora = 0
     End If
     rs_datos!compra_fecha = dtpFecha1.Value
     rs_datos!solicitud_tipo = "15"
     rs_datos!venta_tipo = dtc_codigo2.Text
     rs_datos!edif_codigo = dtc_codigo3.Text
     If dtc_codigo4.Text = "" Or dtc_codigo4.Text = "0" Then
        rs_datos!beneficiario_codigo = dtc_aux3.Text
     Else
        rs_datos!beneficiario_codigo = dtc_codigo4.Text
     End If
     Txt_descripcion.Text = lbl_titulo + " - Edificio: " + dtc_desc3.Text + " Cite: " + Txt_campo2.Caption
     rs_datos!compra_descripcion = Txt_descripcion.Text

     Select Case dtc_codigo2.Text
        Case "1"    'SOLO COMPRAS BB y SS
            rs_datos!proceso_codigo = "CMX"
            rs_datos!subproceso_codigo = "CMX-01"
            rs_datos!etapa_codigo = "CMX-01-01"
            rs_datos!clasif_codigo = "CMX"          'IIf(dtc_codigo8.Text = "", "TEC", dtc_codigo8.Text)
            rs_datos!doc_codigo = "R-207"                'IIf(dtc_codigo9.Text = "", "R-XXX", dtc_codigo9.Text)
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
            If VAR_UNI = "DNIREP" Then
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
     End Select
     rs_datos!poa_codigo = dtc_codigo10.Text
     If txt_obs.Text = "" Then
        rs_datos!compra_observaciones = dtc_desc3.Text   'txt_obs.Text
     Else
        rs_datos!compra_observaciones = txt_obs.Text
     End If
     rs_datos!beneficiario_codigo_resp = dtc_codigo11.Text

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
     
        FraDet3.Visible = True
        FraDet2.Visible = True
        FraDet1.Visible = True
        FrmABMDet3.Visible = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True
     
     VAR_SW = ""
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

' Boton editar
Private Sub BtnModificar_Click()
On Error GoTo EditErr
  If Ado_datos.Recordset.RecordCount > 0 Then
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        
        FraDet3.Visible = False
        FraDet2.Visible = False
        FraDet1.Visible = False
        FrmABMDet3.Visible = False
        FrmABMDet2.Visible = False
        FrmABMDet.Visible = False
        
        VAR_SW = "MOD"
        dtc_desc11.SetFocus
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

' Boton anular
Private Sub BtnEliminar_Click()
  On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
    'If ExisteReg(Ado_datos.Recordset!edif_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado previamente ...", vbInformation + vbOKOnly, "Atención": Exit Sub
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

' Boton aprobar
Private Sub BtnAprobar_Click()
On Error GoTo UpdateErr
  If Ado_datos.Recordset.RecordCount > 0 Then
   If Ado_datos.Recordset!beneficiario_codigo = "0" Or Ado_datos.Recordset!beneficiario_codigo = "" Then
        MsgBox "No se puede APROBAR, debe registrar al Propietario del Proyecto de Edificación: " + lbl_campo4.Caption, vbExclamation, "Validación de Registro"
        Exit Sub
   End If
   Set rs_aux1 = New ADODB.Recordset
   rs_aux1.Open "Select * from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
   If rs_aux1.RecordCount > 0 Then
        VAR_CONT2 = rs_aux1.RecordCount
   End If

   If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        
        Dim sqlUp As String
        ' -------------- Actualiza estado proceso
        Select Case Glaux
        Case "PROVI"    'PROVISION DE EQUIPOS
            sqlUp = " UPDATE ao_compra_cabecera SET estado_codigo_eqp = 'APR', estado_codigo_tra = 'REG'  WHERE  compra_codigo = " + Label1.Caption + "  "
        Case "TRANS"    'TRANSPORTE
            sqlUp = " UPDATE ao_compra_cabecera SET estado_codigo_tra = 'APR', estado_codigo_nac = 'REG'  WHERE  compra_codigo = " + Label1.Caption + "  "
        Case "ADUAN"    'DESADUANIZACION
            sqlUp = " UPDATE ao_compra_cabecera SET estado_codigo_nac = 'APR', estado_codigo_des = 'REG'  WHERE  compra_codigo = " + Label1.Caption + "  "
        Case "DESCA"    'DESCARGUIO Y OTROS
            sqlUp = " UPDATE ao_compra_cabecera SET estado_codigo_des = 'APR', estado_codigo_tec = 'REG'  WHERE  compra_codigo = " + Label1.Caption + "  "
        Case "CONTR"    'CONTRATACION TECNICO
            sqlUp = " UPDATE ao_compra_cabecera SET estado_codigo_tec = 'APR', estado_codigo = 'APR'  WHERE  compra_codigo = " + Label1.Caption + "  "
        End Select
       
        db.Execute sqlUp
    
        MsgBox " El registro se aprobo correctamente."
        Call OptFilGral1_Click
      ' ====== Codigo anterior aprovar 01/09/2016
'        Select Case dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "3"    ' COMPRA-VENTA BB Y SS - COMERCIAL
'                Set rs_aux1 = New ADODB.Recordset
'
'                SQL_FOR = "select * from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   "
'                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'
'                    rs_aux2.Open "Select max(trafico_codigo) as Codigo from ao_solicitud_calculo_trafico where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "Select edif_capacidad_min_trafico as Codigo from ao_solicitud_edificacion where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        VAR_AUX = rs_aux2!Codigo
'                    End If
'                    rs_aux1.AddNew
'                    rs_aux1!ges_gestion = Year(Date)
'                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                    rs_aux1!edif_codigo = Ado_detalle1.Recordset!edif_codigo
'                    rs_aux1!trafico_codigo = var_cod
'                    rs_aux1!trafico_h_capacidad_trafico_parametro = Round(VAR_AUX, 2)
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!fecha_registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
'                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
'
'
'            Case "4"    'VENTA DE SERVICIOS (INST, AJUSTE, REP, EMERG, MANT)
'            Case "5"    ' SERVICIO MODERNIZACION
'        End Select
'        Set rs_aux2 = New ADODB.Recordset
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & dtc_codigo9 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        rs_datos!doc_numero = txt_campo1.Caption
'        VAR_ARCH = "COM_" + RTrim(RTrim(dtc_codigo9) + "-") + LTrim(Str(Val(txt_campo1.Caption)))
'
'        rs_datos!archivo_respaldo_cargado = "N"
'        rs_datos!archivo_respaldo = (VAR_ARCH + ".PDF")
'        'rs_datos!estado_codigo = "APR"
'        rs_datos!fecha_registro = Date
'        rs_datos!usr_codigo = glusuario
'
'
'
'        rs_datos.UpdateBatch adAffectAll
        
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene DETALLE ...", vbExclamation, "Validación de Registro"
   End If
  Else
      MsgBox "NO se puede APROBAR !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

' Boton buscar
Private Sub BtnBuscar_Click()
  If Ado_datos.Recordset.RecordCount > 0 Then
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
  Else
      MsgBox "NO se puede Procesar !!. Verifique si existe el registro. ", vbExclamation, "Atención!"
  End If
End Sub

' Boton imprimir
Private Sub BtnImprimir_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
   fw_seg_cobranza_parametro.lblTipoReporte = "pago"
   fw_seg_cobranza_parametro.Show vbModal
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If
End Sub

Private Sub BtnSalir_Click()
 Unload Me
End Sub

' Boton digitalizar
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
         e = NombreCarpeta

      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
      SW0 = 1
    Else
      sino = MsgBox("El archivo ya existe, elija: <SI> para Volver a Cargarlo. <NO> para Visualizarlo. ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "DED2"
            e = NombreCarpeta
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
          SW0 = 1
      Else
        SW0 = 0
        e = ShellExecute(0, vbNullString, App.Path & "\BIENES\EDIFICIOS\" & Trim(VAR_DPTO) & "\" & Trim(Ado_datos.Recordset("edif_codigo")) & "\" & Trim(Ado_datos.Recordset("archivo_respaldo")), vbNullString, vbNullString, vbNormalFocus)
      End If
    End If
    
  Else
       MsgBox "No se puede Guardar el documento PDF, debe APROBAR previamente el registro ...", vbExclamation, "Validación de Registro"
  End If
QError:
    ' Manejo de errores
    If Err.Number > 0 Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
        Screen.MousePointer = vbDefault
    End If
End Sub

' Boton cancelar
Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
        rs_datos.CancelUpdate
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
        
        FraDet3.Visible = True
        FraDet2.Visible = True
        FraDet1.Visible = True
        FrmABMDet3.Visible = True
        FrmABMDet2.Visible = True
        FrmABMDet.Visible = True

        VAR_SW = ""
    End If

End Sub




' ============ 1 COMANDOS fin =================================

' ============ 2 CONTROLES ====================================

' Evento de ado de grilla.
Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Esto mostrará la posición de registro actual para este Recordset
 
End Sub








' Combo 1
Private Sub dtc_aux1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_aux1.BoundText
    dtc_codigo1.BoundText = dtc_aux1.BoundText
End Sub

' Combo 2
Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    dtc_aux1.BoundText = dtc_codigo1.BoundText
End Sub

' Combo edificio 1
Private Sub dtc_aux3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_aux3.BoundText
    dtc_desc3.BoundText = dtc_aux3.BoundText
End Sub


' Combo tipo compra
Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

' Combo edificio 2
Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    dtc_aux3.BoundText = dtc_codigo3.BoundText
End Sub

' Combo representante
Private Sub dtc_codigo4_Click(Area As Integer)
 dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub


' Combo responsable proceso
Private Sub dtc_desc11_Click(Area As Integer)
    dtc_codigo11.BoundText = dtc_desc11.BoundText
End Sub

' Combo responsable proceso codigo
Private Sub dtc_codigo11_Click(Area As Integer)
    dtc_desc11.BoundText = dtc_codigo11.BoundText
End Sub

' Carga de formulario
Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    parametro = Aux

   ' Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click

    mbDataChanged = False
   ' Fra_datos.Enabled = False
    dg_datos.Enabled = True

    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
  
   numAdjudicacion = 0
   numAdjudicacionCargado = 0
	Call SeguridadSet(Me)
End Sub

' Numero compra evento cambio de texto.
Private Sub Label1_Change()
    ' Call ABRIR_TABLA_DET
     'Esto mostrará la posición de registro actual para este Recordset
  If Ado_datos.Recordset.RecordCount > 0 Then
    ' <-- Inicio                Identificación del Cliente                Fin -->   'esto es de Caption
    If VAR_SW <> "ADD" Then
        
        Call ABRIR_TABLA_DET
        
        VAR_COD4 = parametro
        VAR_SOL = Ado_datos.Recordset!solicitud_codigo
    Else
        Set dg_det2.DataSource = rsNada
    End If
    If Ado_datos.Recordset!estado_codigo_eqp = "APR" Then
            'FrmABMDet2.Visible = False
    Else
           ' FrmABMDet2.Visible = True
    End If
  Else
    Set dg_det1.DataSource = rsNada
    Set dg_det2.DataSource = rsNada
    Set dg_det3.DataSource = rsNada
  End If
End Sub

' Radio de grid pendiente
Private Sub OptFilGral1_Click()
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    
    'queryinicial = " EXEC ra_select_aocompracabecera "
    queryinicial = " SELECT * FROM av_compra_cabecera "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    
End Sub

' Radio de grid todos
Private Sub OptFilGral2_Click()
 Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = " Select * from ao_compra_cabecera where unidad_codigo_adm = '" & parametro & "' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    rs_datos.Sort = "solicitud_codigo"
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

' Text concepto
Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

' Combo concepto
Private Sub txt_obs_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

' Text actividad
Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

' Text 2 actividad
Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub


' ============ 2 CONTROLES fin =================================

' ============ 3 FUNCIONES =====================================

' Funcion carga grillas detalle 1
Private Sub ABRIR_TABLA_DET()
    Dim sqlgrid As String
    sqlgrid = " SELECT * FROM av_compra_detalle_tipo WHERE compra_codigo = " & Ado_datos.Recordset!compra_codigo & " and par_codigo = '43340' "
     ' Asigna tipo proceso comex
    txtprocesocomex.Text = "CMX-01"
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open sqlgrid, db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        If Ado_detalle1.Recordset!pais_codigo <> Nulo Then
             VAR_PAIS = Ado_detalle1.Recordset!pais_codigo
        Else
             VAR_PAIS = ""
        End If
        dg_det1.Visible = True
        Set dg_det1.DataSource = Ado_detalle1.Recordset
    Else
        dg_det1.Visible = False
        Set dg_det1.DataSource = rsNada
    End If
    
    Call CargarGridDetalles
        
End Sub

' Funcion carga grid detalle de adjudicacion y orgen
Public Sub CargarGridDetalles()
   Dim sqlAux As String
   
   numAdjudicacionCargado = numAdjudicacion
   
   Set rs_det2 = New ADODB.Recordset
    If rs_det2.State = 1 Then rs_det2.Close
    ' ---- Consulta filtra adjudicaciones por proceso de comex.
    Dim sqlAdju As String
    sqlAdju = " EXEC ra_select_aocompraadjudica " & Ado_datos.Recordset!solicitud_codigo & "  "
    rs_det2.Open sqlAdju, db, adOpenKeyset, adLockOptimistic, adCmdText
    'rs_det2.Open "select * from ao_compra_adjudica where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle2.Recordset = rs_det2
    If Ado_detalle2.Recordset.RecordCount > 0 Then
    
        Set rs_det3 = New ADODB.Recordset
        If rs_det3.State = 1 Then rs_det3.Close
        sqlAux = " EXEC ra_select_pagos " & rs_det2!compra_codigo & ", " & numAdjudicacion & "  "
        rs_det3.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle3.Recordset = rs_det3
        If Ado_detalle3.Recordset.RecordCount > 0 Then
                dg_det3.Visible = True
                Set dg_det3.DataSource = Ado_detalle3.Recordset
            Else
                dg_det3.Visible = False
                Set dg_det3.DataSource = rsNada
        End If
        dg_det2.Visible = True
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    Else
        dg_det3.Visible = False
        Set dg_det3.DataSource = rsNada
        dg_det2.Visible = False
        Set dg_det2.DataSource = rsNada
    End If

End Sub

' Funcion carga grid detalle pago orgen
Public Sub CargarGridDetallesOrden()
   Dim sqlAux As String
   
   numAdjudicacionCargado = numAdjudicacion

    If Ado_detalle2.Recordset.RecordCount > 0 Then
    
        Set rs_det3 = New ADODB.Recordset
        If rs_det3.State = 1 Then rs_det3.Close
        sqlAux = "select * from ao_compra_planilla_pagos where compra_codigo = " & rs_det2!compra_codigo & " and adjudica_codigo = " & numAdjudicacion & "  "
        rs_det3.Open sqlAux, db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle3.Recordset = rs_det3
        If Ado_detalle3.Recordset.RecordCount > 0 Then
                dg_det3.Visible = True
                Set dg_det3.DataSource = Ado_detalle3.Recordset
            Else
                dg_det3.Visible = False
                Set dg_det3.DataSource = rsNada
        End If

    Else
        dg_det3.Visible = False
        Set dg_det3.DataSource = rsNada
      
    End If

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

Private Function ExisteReg(Unidad As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE dgral_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

' ============ 3 FUNCIONES fin =====================================

' ============ 4 COMANDO IMPORTACION ===============================
' Boton importacion nuevo.
'Private Sub BtnAddDetalle1_Click()
'      marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos!estado_codigo = "REG" Then
'    swnuevo = 1
'    fraOpciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet2.Enabled = False
'    FrmABMDet2.Enabled = False
'    FraDet3.Enabled = False
'    FrmABMDet3.Enabled = False
'    Fra_datos.Enabled = False
'    Call ABRIR_TABLA_DET
'    Select Case Glaux
'        Case "PROVI"    'PROVISION DE EQUIPOS
'            'NO HAY
'        Case "TRANS"    'TRANSPORTE
'            Ado_detalle2.Recordset.AddNew
'            frm_solicitud_bienes2.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes2.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes2.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes2.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes2.lbl_det.Caption = Glaux
'            frm_solicitud_bienes2.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes2.Show vbModal
'        Case "ADUAN"    'DESADUANIZACION
'            Ado_detalle2.Recordset.AddNew
'            frm_solicitud_bienes2.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes2.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes2.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes2.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes2.lbl_det.Caption = Glaux
'            frm_solicitud_bienes2.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes2.Show vbModal
'        Case "DESCA"    'DESCARGUIO Y OTROS
'            Ado_detalle2.Recordset.AddNew
'            frm_solicitud_bienes2.txt_codigo.Caption = Me.txt_codigo.Caption
'            frm_solicitud_bienes2.Txt_campo1.Caption = Me.dtc_codigo1.Text
'            frm_solicitud_bienes2.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_solicitud_bienes2.lbl_edif.Caption = dtc_codigo3.Text
'            frm_solicitud_bienes2.lbl_det.Caption = Glaux
'            frm_solicitud_bienes2.Txt_estado.Caption = "REG"
'            frm_solicitud_bienes2.Show vbModal
'    End Select
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraNavega.Enabled = True
'    FraDet2.Enabled = True
'    FrmABMDet2.Enabled = True
'    FraDet3.Enabled = True
'    FrmABMDet3.Enabled = True
''    Fra_datos.Enabled = True
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
'End Sub

' Boton modificar importacion
'Private Sub BtnModDetalle1_Click()
'    If Ado_detalle1.Recordset.RecordCount > 0 Then
'      If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
'        marca1 = Ado_detalle1.Recordset.Bookmark
'        swnuevo = 2
'        fraOpciones.Enabled = False
'        FraNavega.Enabled = False
'        FraDet2.Enabled = False
'        FrmABMDet2.Enabled = False
'        FraDet3.Enabled = False
'        FrmABMDet3.Enabled = False
'        Fra_datos.Enabled = False
'
'        Select Case dtc_codigo2.Text
'            Case "1"    'SOLO COMPRAS BB y SS
'            Case "2"    'SOLO VENTA DE BIENES
'            Case "COM-01"    '3. COMPRA-VENTA BB Y SS - COMERCIAL
'            Case "L"    'IMPORTACION DIRECTA CLIENTE
'                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
'
'                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
'                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
'                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
'                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
'                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
'                frm_solicitud_bienes.Txt_campo19.Text = Me.Ado_detalle1.Recordset("bien_cantidad_por_empaque")
'
'                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
'                frm_solicitud_bienes.Txt_campo15.Text = "10" 'Me.Ado_detalle1.Recordset("fosa_dimension_frente")
'
'                frm_solicitud_bienes.lbl_det.Caption = "43340"
'                frm_solicitud_bienes.Show vbModal
'            Case "V"    'FACTURACION LOCAL
'                frm_solicitud_bienes.txt_codigo.Caption = Me.Ado_detalle1.Recordset("solicitud_codigo")  'cod_cabecera
'                frm_solicitud_bienes.Txt_campo1.Caption = Me.Ado_detalle1.Recordset("unidad_codigo")  'Unidad
'                frm_solicitud_bienes.Txt_descripcion.Caption = Me.dtc_desc1.Text
'
'                frm_solicitud_bienes.lbl_edif.Caption = dtc_codigo3.Text
'                frm_solicitud_bienes.Txt_campo5.Text = Me.Ado_detalle1.Recordset("bien_codigo")
'
'                frm_solicitud_bienes.txt_campo6.Text = Me.Ado_detalle1.Recordset("bien_descripcion")
'                frm_solicitud_bienes.Txt_campo7.Text = Me.Ado_detalle1.Recordset("bien_descripcion_anterior")
'                frm_solicitud_bienes.Txt_campo8.Text = Me.Ado_detalle1.Recordset("marca_codigo")
'                frm_solicitud_bienes.Txt_campo9.Text = Me.Ado_detalle1.Recordset("modelo_codigo")
'
'                frm_solicitud_bienes.Txt_campo16.Text = Me.Ado_detalle1.Recordset("bien_cantidad")
'                frm_solicitud_bienes.Txt_campo10.Text = Me.Ado_detalle1.Recordset("bien_precio_venta_base")
'                frm_solicitud_bienes.Txt_campo11.Caption = Me.Ado_detalle1.Recordset("bien_total_venta")
'
'                frm_solicitud_bienes.Txt_campo14.Text = Me.Ado_detalle1.Recordset("unimed_codigo")
'                frm_solicitud_bienes.lbl_det.Caption = "43340"
'                frm_solicitud_bienes.Show vbModal
'
'        End Select
'        swnuevo = 0
'        fraOpciones.Enabled = True
'        FraNavega.Enabled = True
'        FraDet2.Enabled = True
'        FrmABMDet2.Enabled = True
'        FraDet3.Enabled = True
'        FrmABMDet3.Enabled = True
'        Call ABRIR_TABLA_DET
'        Ado_detalle1.Recordset.Move marca1 - 1
'      Else
'        MsgBox "No se puede MODIFICAR, porque ya está APROBADO o ANULADO, Verifique por favor!! ", vbExclamation
'      End If
'  Else
'     MsgBox "No se puede MODIFICAR, el registro No fue identificado o No Existe, Verifique por favor ...", vbExclamation, "Validación de Registro"
'  End If
'
'End Sub

' ============ 4 COMANDO IMPORTACION fin ===============================

' ============ 5 COMANDO ADJUDICACION ===============================
' Boton nuevo adjudicacion
'Private Sub BtnAddDetalle2_Click()
'
'    If numAdjudicacion = 0 Then
'     numAdjudicacion = 1
'    End If
'  ' ========= GENERA ADJUDICA CODIGO
'    Dim rs_codadj As ADODB.Recordset
'    Set rs_codadj = New ADODB.Recordset
'    Dim codigoA As String
'    If rs_codadj.State = 1 Then rs_codadj.Close
'    rs_codadj.Open "SELECT (ISNULL(MAX(adjudica_codigo),0) + 1) AS codigo FROM ao_compra_adjudica WHERE compra_codigo = " + Label1.Caption + " ", db, adOpenStatic
'    If rs_codadj.RecordCount > 0 Then
'        txtcodigoadj.Text = rs_codadj!Codigo
'    Else
'        txtcodigoadj.Text = "1"
'    End If
'
'
'  If rs_datos!estado_codigo = "REG" Then
'    VAR_PAIS = "BRA"
'    If VAR_PAIS = "NN" Then
'        MsgBox "ERROR, No ha sido registrada la industria. consulte con Gerencia Comercial y vuelva a intentar !! ", vbExclamation
'    Else
'        'FOB + SEG de la Cotizacion
'        Set rs_datos5 = New ADODB.Recordset
'        If rs_datos5.State = 1 Then rs_datos5.Close
'        rs_datos5.Open "Select * from ao_solicitud_cotiza_venta where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & " and pais_codigo = '" & VAR_PAIS & "' ", db, adOpenStatic
'        If rs_datos5.RecordCount > 0 Then
'            If IsNull(rs_datos5!cotiza_fob_seg_dol) Then
'                MsgBox "ERROR, No ha sido registrado el precio FOB. Consulte con Gerencia Comercial y vuelva a intentar !! ", vbExclamation
'                Exit Sub
'            Else
'                VAR_FOBSEG = rs_datos5!cotiza_fob_seg_dol
'                VAR_FOBSEG2 = IIf(IsNull(rs_datos5!cotiza_fob_seg_bs), rs_datos5!cotiza_fob_seg_dol * GlTipoCambioOficial, rs_datos5!cotiza_fob_seg_bs)
'            End If
'        Else
'            MsgBox "ERROR, No ha sido identivicado el registro. Consulte con Gerencia Comercial y vuelva a intentar !! ", vbExclamation
'            Exit Sub
'        End If
'
'        swnuevo = 1
'        fraOpciones.Enabled = False
'        FraNavega.Enabled = False
'        FraDet2.Visible = False
'        'FrmABMDet2.Visible = False
'        FraDet3.Visible = False
'        'FrmABMDet3.Visible = False
'        Fra_datos.Enabled = False
'
'                Call ABRIR_TABLA_DET
'                Ado_detalle2.Recordset.AddNew
'                frm_ao_comex_adjudica.txt_codigo.Caption = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
'                frm_ao_comex_adjudica.Txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
'                frm_ao_comex_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
'                frm_ao_comex_adjudica.txtCodigo1.Caption = Me.Ado_datos.Recordset!compra_codigo
'                frm_ao_comex_adjudica.lbl_adjudica.Caption = Me.Ado_detalle2.Recordset!adjudica_codigo
'                frm_ao_comex_adjudica.txtSW.Text = Me.Ado_datos.Recordset!venta_tipo
'                frm_ao_comex_adjudica.txt_total_dol = VAR_FOBSEG
'                frm_ao_comex_adjudica.txt_total_bs = VAR_FOBSEG2
'                frm_ao_comex_adjudica.txt_pais.Text = VAR_PAIS
'                frm_ao_comex_adjudica.Txtestado.Text = "REG"
'                frm_ao_comex_adjudica.Show vbModal
'
'        swnuevo = 0
'        fraOpciones.Enabled = True
'        FraNavega.Enabled = True
'        FraDet2.Visible = True
'       ' FrmABMDet2.Visible = True
'        FraDet3.Visible = True
'        'FrmABMDet3.Visible = True
'
'    End If
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
'
'End Sub

' Boton modificacion adjudicacion
'Private Sub BtnModDetalle2_Click()
'  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" Then
'     If Ado_detalle2.Recordset.RecordCount > 0 Then
'        swnuevo = 2
'        fraOpciones.Enabled = False
'        FraNavega.Enabled = False
'        FraDet2.Enabled = False
'        FrmABMDet2.Enabled = False
'        FraDet3.Enabled = False
'        FrmABMDet3.Enabled = False
'        Fra_datos.Enabled = False
'
'         txtcodigoadj.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("adjudica_codigo")), "1", Me.Ado_detalle2.Recordset("adjudica_codigo"))
'
'            frm_ao_comex_adjudica.txt_codigo.Caption = Me.Ado_detalle2.Recordset("solicitud_codigo")  'cod_cabecera
'            frm_ao_comex_adjudica.Txt_campo1.Text = Me.Ado_detalle2.Recordset("unidad_codigo")  'Unidad
'            frm_ao_comex_adjudica.Txt_descripcion.Caption = Me.dtc_desc1.Text
'            frm_ao_comex_adjudica.txtCodigo1.Caption = Me.Ado_detalle2.Recordset("compra_codigo")
'            'frm_ao_comex_adjudica.Txt_estado.Caption = "REG"
'            frm_ao_comex_adjudica.lbl_adjudica.Caption = Me.Ado_detalle2.Recordset("adjudica_codigo")
'            frm_ao_comex_adjudica.dtc_codigo5.Text = Me.Ado_detalle2.Recordset("beneficiario_codigo")
'            frm_ao_comex_adjudica.dtc_desc5.BoundText = frm_ao_comex_adjudica.dtc_codigo5.BoundText
'            frm_ao_comex_adjudica.dtc_aux4.BoundText = frm_ao_comex_adjudica.dtc_codigo5.BoundText
'            frm_ao_comex_adjudica.dtc_aux5.BoundText = frm_ao_comex_adjudica.dtc_codigo5.BoundText
'
'            frm_ao_comex_adjudica.txt_Nota.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("nro_nota_remision")), "", Me.Ado_detalle2.Recordset("nro_nota_remision"))
'            frm_ao_comex_adjudica.txt_total_bs.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("adjudica_monto_bs")), 0, Me.Ado_detalle2.Recordset("adjudica_monto_bs"))
'            frm_ao_comex_adjudica.txt_total_dol.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!adjudica_monto_dol), 0, Me.Ado_detalle2.Recordset!adjudica_monto_dol)
'            frm_ao_comex_adjudica.TxtFecha.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_inicio_contrato")), Date, Me.Ado_detalle2.Recordset("fecha_inicio_contrato"))
'            frm_ao_comex_adjudica.TxtFecha2.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_fin_contrato")), Date, Me.Ado_detalle2.Recordset("fecha_fin_contrato"))
'            frm_ao_comex_adjudica.TxtFecha3.Value = IIf(IsNull(Me.Ado_detalle2.Recordset("fecha_envio_proveedor")), Date, Me.Ado_detalle2.Recordset("fecha_envio_proveedor"))
'
'            frm_ao_comex_adjudica.cmb_mes_ini = IIf(IsNull(Me.Ado_detalle2.Recordset!mes_inicio_crono), "ENERO", Me.Ado_detalle2.Recordset!mes_inicio_crono)
'            frm_ao_comex_adjudica.txtCantCuota.Text = IIf(IsNull(Me.Ado_detalle2.Recordset!cantidad_cuotas_pag), "1", Me.Ado_detalle2.Recordset!cantidad_cuotas_pag)
'            frm_ao_comex_adjudica.cmd_unimed2 = IIf(IsNull(Me.Ado_detalle2.Recordset!unimed_codigo_pag), "MES", Me.Ado_detalle2.Recordset!unimed_codigo_pag)
'
'            frm_ao_comex_adjudica.txtSW.Text = Me.Ado_datos.Recordset!venta_tipo
'            frm_ao_comex_adjudica.txt_pais.Text = VAR_PAIS
'
'            frm_ao_comex_adjudica.Show vbModal
'        swnuevo = 0
'        fraOpciones.Enabled = True
'        FraNavega.Enabled = True
'        FraDet2.Enabled = True
'        FrmABMDet2.Enabled = True
'        FraDet3.Enabled = True
'        FrmABMDet3.Enabled = True
'     Else
'        MsgBox "No se puede Modificar un registro inexistente, vuelva a intentar!! ", vbExclamation
'     End If
'  Else
'    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
'
'End Sub

' Boton borrar adjudicacion
Private Sub BtnAnlDetalle2_Click()
   sino = MsgBox("Está Seguro de ANULAR el Registro Activo ? ", vbYesNo + vbQuestion, "Atención")
   If Ado_detalle1.Recordset("estado_codigo") = "REG" Then
      If sino = vbYes Then
        Ado_detalle1.Recordset.Delete 'adAffectAll
      End If
   Else
        MsgBox "No se puede ANULAR un registro Aprobado ...", vbExclamation, "Validación de Registro"
   End If
End Sub

' ============ 5 COMANDO ADJUDICACION fin ===============================

' ============ 6 COMANDO ORDEN PAGO =====================================
' Boton nuevo orden pago
'Private Sub BtnAddDetalle_Click()
'    ' ========= GENERA ADJUDICA CODIGO
'    Dim rs_codadj As ADODB.Recordset
'    Set rs_codadj = New ADODB.Recordset
'    Dim codigoA As String
'    If rs_codadj.State = 1 Then rs_codadj.Close
'    rs_codadj.Open "SELECT (ISNULL(MAX(pago_codigo),0) + 1) AS codigo FROM ao_compra_planilla_pagos WHERE compra_codigo = " + Label1.Caption + " ", db, adOpenStatic
'    If rs_codadj.RecordCount > 0 Then
'        txtcodigopago.Text = rs_codadj!Codigo
'        lblcodigopagosig.Caption = rs_codadj!Codigo
'    Else
'        txtcodigopago.Text = "1"
'    End If
'
'    txtcodigoadj.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("adjudica_codigo")), "1", Me.Ado_detalle2.Recordset("adjudica_codigo"))
'
'
'  If rs_datos!estado_codigo = "REG" Then
'    swnuevo = 1
'    fraOpciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet2.Visible = False
'    FrmABMDet2.Visible = False
'    FraDet3.Visible = False
'    FrmABMDet3.Visible = False
'    Fra_datos.Enabled = False
'
'            Call ABRIR_TABLA_DET
'            Ado_detalle3.Recordset.MoveLast
'            Ado_detalle3.Recordset.AddNew
'
'            frm_ao_comex_pagos.txt_codigo.Caption = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
'    frm_ao_comex_pagos.Txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
'    frm_ao_comex_pagos.Txt_descripcion = Me.dtc_desc1.Text
'    frm_ao_comex_pagos.txtCodigo1.Caption = Me.Ado_datos.Recordset!compra_codigo
'    frm_ao_comex_pagos.lbl_adjudica.Caption = Me.Ado_detalle3.Recordset!adjudica_codigo
'    frm_ao_comex_pagos.Show vbModal
'
'    swnuevo = 0
'    fraOpciones.Enabled = True
'    FraNavega.Enabled = True
'    FraDet2.Visible = True
'    FrmABMDet2.Visible = True
'    FraDet3.Visible = True
'    FrmABMDet3.Visible = True
'
'  Else
'    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
'  End If
'
'End Sub

' Boton modificar orden pago
'Private Sub BtnModDetalle_Click()
'  marca1 = Ado_datos.Recordset.Bookmark
'  If rs_datos.RecordCount > 0 And rs_datos!estado_codigo = "REG" And Ado_detalle1.Recordset.RecordCount > 0 Then
'    swnuevo = 2
'    fraOpciones.Enabled = False
'    FraNavega.Enabled = False
'    FraDet1.Enabled = False
'    FrmABMDet.Enabled = False
'    FraDet2.Enabled = False
'    FrmABMDet2.Enabled = False
'    Fra_datos.Enabled = False
'
'            txtcodigopago.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_fecha_prog), "1", Me.Ado_detalle3.Recordset!pago_codigo)
'            txtcodigoadj.Text = IIf(IsNull(Me.Ado_detalle2.Recordset("adjudica_codigo")), "1", Me.Ado_detalle2.Recordset("adjudica_codigo"))
'
'            frm_ao_comex_pagos.txt_codigo.Caption = Me.Ado_datos.Recordset!solicitud_codigo  'cod_cabecera
'            frm_ao_comex_pagos.Txt_campo1.Text = Me.Ado_datos.Recordset!unidad_codigo  'Unidad
'            frm_ao_comex_pagos.Txt_descripcion = Me.dtc_desc1.Text
'            frm_ao_comex_pagos.txt_codigo.Caption = Me.Ado_datos.Recordset!compra_codigo
'            frm_ao_comex_pagos.lbl_adjudica.Caption = Me.Ado_detalle3.Recordset!adjudica_codigo
'            frm_ao_comex_pagos.txtCodigo1.Caption = Me.Ado_detalle3.Recordset!pago_codigo
'            frm_ao_comex_pagos.Txt_campo1.Text = Me.Ado_detalle3.Recordset!beneficiario_codigo
'            frm_ao_comex_pagos.Txt_descripcion.BoundText = frm_ao_comex_pagos.Txt_campo1.BoundText
'
'            frm_ao_comex_pagos.DTPFechaProg.Value = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_fecha_prog), Date, Me.Ado_detalle3.Recordset!pago_fecha_prog)
'            frm_ao_comex_pagos.DTPFechaPago.Value = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_fecha_efectiva), Date, Me.Ado_detalle3.Recordset!pago_fecha_efectiva)
'            frm_ao_comex_pagos.TxtMontoBs.Text = Me.Ado_detalle3.Recordset!pago_total_bs
'            frm_ao_comex_pagos.TxtMontoDol.Text = Me.Ado_detalle3.Recordset!pago_total_dol
'            frm_ao_comex_pagos.txt_factura.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_nro_cmpbte_factura), 0, Me.Ado_detalle3.Recordset!pago_nro_cmpbte_factura)
'            frm_ao_comex_pagos.txtDoc.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_nro_autorizacion), 0, Me.Ado_detalle3.Recordset!pago_nro_autorizacion)
'
'            frm_ao_comex_pagos.TxtConcepto.Text = Me.Ado_detalle3.Recordset!pago_descripcion
'            frm_ao_comex_pagos.txt_respaldos.Text = IIf(IsNull(Me.Ado_detalle3.Recordset!pago_respaldos), "FACTURA", Me.Ado_detalle3.Recordset!pago_respaldos)
'
''            frm_ao_comex_pagos.Txtestado.Text = "REG"
'            frm_ao_comex_pagos.Show vbModal
'
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
'
'  Else
'    MsgBox "No se puede Modificar el registro, verifique si está Aprobado o fue correctamente identificado !! ", vbExclamation
'  End If
'
'End Sub

' Boton aprobar orden pago
Private Sub BtnAprobar2_Click()
Dim sMsg As String
On Error GoTo ErrHandler

        If Ado_detalle3.Recordset.BOF = False Then
          If Ado_detalle3.Recordset!estado_codigo = "REG" Then
                Dim sqlA, parcodigo As String
                If Ado_detalle1.Recordset.BOF = False Then
                     parcodigo = Ado_detalle1.Recordset!par_codigo
                Else
                     parcodigo = 1
                End If
                'pro_codigo
                sqlA = " EXEC ao_aprobar_compraplanillapago '" & CStr(Ado_detalle3.Recordset!ges_gestion) & "','111', " & CStr(Ado_detalle3.Recordset!compra_codigo) & ", '" & CStr(Ado_detalle3.Recordset!adjudica_codigo) & "', '" & CStr(Ado_detalle3.Recordset!pago_codigo) & "' , '" + parcodigo + "','1','" & CStr(glusuario) & "' "
                db.Execute sqlA
                
                MsgBox "El registro se aprobo correctamente."
                Call CargarGridDetallesOrden
          Else
              MsgBox "El registro no se puede aprobar por que el estado es diferente de REG."
          End If
        Else
          MsgBox "Seleccione una orden de pago"
        End If
ErrHandler:
If Err.Description <> "" Then
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    MsgBox sMsg
End If
End Sub

' Boton imprimir orden pago
Private Sub BtnImprimir1_Click()
  If (Ado_datos.Recordset.RecordCount > 0) Then
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        cr01.ReportFileName = App.Path & "\Reportes\tecnico\tr_identificacion_cliente.rpt"
        cr01.WindowShowPrintSetupBtn = True
        cr01.WindowShowRefreshBtn = True
          cr01.Formulas(0) = "Titulo = '" & lbl_titulo.Caption & "' "
          cr01.Formulas(1) = "Subtitulo = '" & FraDet1.Caption & "' "

        cr01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
        cr01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo
        cr01.StoredProcParam(2) = Me.Ado_datos.Recordset!solicitud_codigo
        iResult = cr01.PrintReport
        If iResult <> 0 Then MsgBox cr01.LastErrorNumber & " : " & cr01.LastErrorString, vbCritical, "Error de impresión"
        cr01.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar datos del Detalle ...", , "Atención"
    End If
  Else
    MsgBox "No se puede Imprimir. Debe elegir el Registro que desea Imprimir ...", , "Atención"
  End If

End Sub
' ============ 6 COMANDO ORDEN PAGO fin =====================================


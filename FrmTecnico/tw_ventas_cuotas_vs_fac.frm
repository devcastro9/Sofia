VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tw_ventas_cuotas_vs_fac 
   Caption         =   "Facturas en Bloque"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16575
   Icon            =   "tw_ventas_cuotas_vs_fac.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   16575
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTA"
      ForeColor       =   &H00C00000&
      Height          =   8955
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   16425
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00404040&
         Height          =   1200
         Left            =   3600
         Picture         =   "tw_ventas_cuotas_vs_fac.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Elije el Registro a Facturar (Envía a DESTINO)"
         Top             =   3600
         Width           =   1245
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00404040&
         Height          =   1200
         Left            =   9360
         Picture         =   "tw_ventas_cuotas_vs_fac.frx":10D1
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Devuelve a ORIGEN ..."
         Top             =   3600
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000018&
         Caption         =   "Facturados y Cobrados"
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
         Left            =   13440
         TabIndex        =   10
         Top             =   8595
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H80000018&
         Caption         =   "Facturados y No Cobrados"
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
         Left            =   9960
         TabIndex        =   9
         Top             =   8595
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H80000018&
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
         Left            =   1080
         TabIndex        =   8
         Top             =   3315
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H00404040&
         Height          =   900
         Left            =   120
         ScaleHeight     =   840
         ScaleWidth      =   16155
         TabIndex        =   2
         Top             =   240
         Width           =   16215
         Begin VB.CommandButton BtnBuscar 
            BackColor       =   &H80000018&
            Caption         =   "1.Buscar"
            Height          =   720
            Left            =   240
            Picture         =   "tw_ventas_cuotas_vs_fac.frx":1D1F
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Busca Registro para Facturar"
            Top             =   60
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton BtnImprimir5 
            BackColor       =   &H80000018&
            Caption         =   "Re-Imprimir"
            Height          =   720
            Left            =   1560
            Picture         =   "tw_ventas_cuotas_vs_fac.frx":2029
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Re-Imprime Factura"
            Top             =   60
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton BtnSalir 
            BackColor       =   &H80000018&
            Caption         =   "Cerrar"
            Height          =   720
            Left            =   14880
            Picture         =   "tw_ventas_cuotas_vs_fac.frx":25B3
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Cerrar Ventana"
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label lbl_titulo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECCIONAR VARIAS CUOTAS PARA UNA FACTURA"
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
            Left            =   4110
            TabIndex        =   5
            Top             =   255
            Width           =   8205
         End
      End
      Begin VB.CommandButton BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Height          =   600
         Left            =   14760
         Picture         =   "tw_ventas_cuotas_vs_fac.frx":2C1C5
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprobar los Registros Elegidos..."
         Top             =   3900
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid dg_datosXXX 
         Bindings        =   "tw_ventas_cuotas_vs_fac.frx":2C9FB
         Height          =   2700
         Left            =   120
         TabIndex        =   6
         Top             =   6720
         Visible         =   0   'False
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777152
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         Caption         =   "SOLICITUDES DE FACTURACION"
         ColumnCount     =   14
         BeginProperty Column00 
            DataField       =   "cobranza_fecha_sol"
            Caption         =   "F.Solicit.Fac"
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
            DataField       =   "cobranza_codigo"
            Caption         =   "No.Cobranza"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Cobrador"
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
            DataField       =   "cobranza_total_bs"
            Caption         =   "Solicitado.Bs."
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
            DataField       =   "edif_codigo_corto"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre del Edificio"
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "estado_codigo_fac"
            Caption         =   "Estado"
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
            DataField       =   "cobranza_fecha_fac"
            Caption         =   "Fecha.Factura"
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
         BeginProperty Column09 
            DataField       =   "cobranza_total_dol"
            Caption         =   "Cobrado en Dol."
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
         BeginProperty Column10 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Respaldo"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "Nro. Factura"
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
            DataField       =   "cobranza_nro_factura"
            Caption         =   "#Factura"
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
            DataField       =   "beneficiario_codigo"
            Caption         =   "NIT/CI del Cliente"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3225.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   4440.189
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_datos1 
         Bindings        =   "tw_ventas_cuotas_vs_fac.frx":2CA13
         Height          =   2700
         Left            =   9120
         TabIndex        =   7
         Top             =   6720
         Visible         =   0   'False
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   4763
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         Caption         =   "FACTURAS EMITIDAS"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "dosifica_autorizacion"
            Caption         =   "#.Autorizacion"
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
            DataField       =   "nro_factura"
            Caption         =   "#Factura"
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
            DataField       =   "fecha_fac"
            Caption         =   "Fecha.Factura"
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
            DataField       =   "total_bs"
            Caption         =   "Facturado.Bs."
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
            DataField       =   "beneficiario_codigo_fac"
            Caption         =   "Factura a Nombre de:"
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
            DataField       =   "beneficiario_nit"
            Caption         =   "NIT/CI.Cliente"
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
            DataField       =   "edif_codigo_corto"
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
         BeginProperty Column07 
            DataField       =   "estado_codigo_fac"
            Caption         =   "Estado"
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
            DataField       =   "beneficiario_RazonSocial"
            Caption         =   "Nombre del Cliente"
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
            DataField       =   "glosa_Descripcion"
            Caption         =   "Concepto de la Factura"
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
            DataField       =   "total_dol"
            Caption         =   "Facurado Dol."
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
         BeginProperty Column11 
            DataField       =   "doc_numero"
            Caption         =   "#.OrdenCobro"
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
            DataField       =   "doc_codigo_fac"
            Caption         =   "Fact/Recibo"
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
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   3555.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   4275.213
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3240
         Width           =   16140
         _ExtentX        =   28469
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
         BackColor       =   -2147483624
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
      Begin MSAdodcLib.Adodc Ado_datos1 
         Height          =   330
         Left            =   9120
         Top             =   8520
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
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
      Begin MSAdodcLib.Adodc Ado_datos2 
         Height          =   330
         Left            =   120
         Top             =   6360
         Visible         =   0   'False
         Width           =   16140
         _ExtentX        =   28469
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
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "tw_ventas_cuotas_vs_fac.frx":2CA2C
         Height          =   1980
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   3493
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         Caption         =   "ORIGEN"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "No.Cuota"
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Mes.Programado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "mmm-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto Programado Bs."
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
         BeginProperty Column03 
            DataField       =   "beneficiario_codigo_resp"
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
         BeginProperty Column04 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Resp."
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
            DataField       =   "cobranza_fecha_conformidad"
            Caption         =   "Fecha.Certif."
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "cobranza_concepto_plazo"
            Caption         =   "Plazo a Cumplir"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "estado_ac"
            Caption         =   "Aviso Cob."
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
            DataField       =   "correl_ac"
            Caption         =   "Nro. Aviso Cob"
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
            DataField       =   "cobranza_programada_dol"
            Caption         =   "Monto a Pagar Dol."
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
         BeginProperty Column12 
            DataField       =   "cobranza_codigo"
            Caption         =   "Cod.Cobranza"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   5955.024
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_datos2 
         Bindings        =   "tw_ventas_cuotas_vs_fac.frx":2CA44
         Height          =   1980
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   16230
         _ExtentX        =   28628
         _ExtentY        =   3493
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
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
         Caption         =   "DESTINO"
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "cobranza_prog_codigo"
            Caption         =   "No.Cuota"
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
            DataField       =   "cobranza_fecha_prog"
            Caption         =   "Mes.Programado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "mmm-yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cobranza_programada_bs"
            Caption         =   "Monto Programado Bs."
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
         BeginProperty Column03 
            DataField       =   "beneficiario_codigo_resp"
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
         BeginProperty Column04 
            DataField       =   "doc_numero"
            Caption         =   "Nro.Doc.Resp."
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
            DataField       =   "cobranza_fecha_conformidad"
            Caption         =   "Fecha.Certif."
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
            DataField       =   "cobranza_observaciones"
            Caption         =   "Concepto de la Cuota"
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
            DataField       =   "cobranza_concepto_plazo"
            Caption         =   "Plazo a Cumplir"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado"
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
            DataField       =   "estado_ac"
            Caption         =   "Aviso Cob."
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
            DataField       =   "correl_ac"
            Caption         =   "Nro. Aviso Cob"
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
            DataField       =   "cobranza_programada_dol"
            Caption         =   "Monto a Pagar Dol."
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
         BeginProperty Column12 
            DataField       =   "cobranza_codigo"
            Caption         =   "Cod.Cobranza"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   5955.024
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1214.929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "3. Aprueba las Cuotas y Envia a Facturación -->"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11160
         TabIndex        =   17
         Top             =   4080
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Devolver.Si el registro elegido no es el correcto --->"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   4080
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Elegir para Agrupar en UNA sola Factura --->"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   3375
      End
   End
End
Attribute VB_Name = "tw_ventas_cuotas_vs_fac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_datos As New ADODB.Recordset     'ORIGEN ao_ventas_cobranza_prog
Dim rs_datos1 As New ADODB.Recordset    'FACTURAS Cabecera
Dim rs_datos2 As New ADODB.Recordset    'DESTINO ao_ventas_cobranza_prog

Dim rs_aux1 As New ADODB.Recordset    'FACTURAS Cabecera (Auxiliar)
Dim rs_aux2 As New ADODB.Recordset    'DESTINO ao_ventas_cobranza_prog (Auxiliar)
Dim rs_aux3 As New ADODB.Recordset    'VENTAS Cabecera (Auxiliar)
Dim rs_aux4 As New ADODB.Recordset    'BENEFICIARIO (Auxiliar)
Dim rs_aux5 As New ADODB.Recordset    'DESTINO ao_ventas_cobranza_prog (suma Importes)

Dim VAR_BENEF, VAR_RAZON, VAR_NIT As String
Dim VAR_BENEF_VTA, VAR_BENEF_RESP As String
Dim var_literal As String

Dim VAR_TOTBS, VAR_TOTDOL As Double

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If Ado_datos.Recordset.RecordCount > 0 Then
       Call abrir_destino
    Else
        If Ado_datos2.Recordset.RecordCount > 0 Then
            Call abrir_destino
            'dg_datos2.Visible = False
        End If
    End If
End Sub

Private Sub abrir_destino()
    Set rs_datos2 = New Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & " and estado_codigo = 'REG' AND es_grupo_fac = 'SI' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos2.Recordset = rs_datos2.DataSource
    Set dg_datos2.DataSource = Ado_datos2.Recordset
    If Ado_datos2.Recordset.RecordCount > 0 Then
         dg_datos2.Visible = True
     Else
         dg_datos2.Visible = False
     End If
End Sub

Private Sub BtnAñadir_Click()
'VALIDAR USUSARIO WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
' If glusuario = "ADMIN" Or glusuario = "FDELGADILLO" Or glusuario = "SQUISPE" Or glusuario = "HMARIN" Or glusuario = "CSALINAS" Then
    If Ado_datos.Recordset.RecordCount > 0 Then
       db.Execute "UPDATE ao_ventas_cobranza_prog SET es_grupo_fac = 'SI' WHERE correl_prog = " & Ado_datos.Recordset!correl_prog & " "
       Call OptFilGral1_Click
       'Call abrir_destino
       
'       Set rs_datos2 = New Recordset
'       If rs_datos2.State = 1 Then rs_datos2.Close
'       rs_datos2.Open "select * From av_ventas_cobranza WHERE (estado_codigo_sol = 'APR' AND estado_codigo_fac = 'REG' AND estado_codigo1 = 'APR' and doc_codigo_fac = 'R-101' AND trans_codigo = 'X' ) ", db, adOpenKeyset, adLockOptimistic
'       Set Ado_datos2.Recordset = rs_datos2.DataSource
'       Set dg_datos2.DataSource = Ado_datos2.Recordset
'       If Ado_datos2.Recordset.RecordCount > 0 Then
'            If Ado_datos2.Recordset!venta_codigo = Ado_datos.Recordset!venta_codigo Then
'                 db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'X'  WHERE cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'            Else
'                 MsgBox "Error, debe Elegir un registro del mismo Edificio, para FACTURAR en bloque, verifique los datos y vuelva a intentar ...", , "Atención"
'                 Exit Sub
'            End If
'        Else
'            db.Execute "UPDATE ao_ventas_cobranza SET estado_codigo1 = 'APR', trans_codigo = 'X'  WHERE cobranza_codigo = " & Ado_datos.Recordset!cobranza_codigo & " "
'            Ado_datos2.Recordset.Requery
'        End If
    End If
' End If
End Sub

Private Sub BtnAprobar_Click()
    nroventa = Ado_datos2.Recordset!venta_codigo
    glGestion = Year(Date)
    VAR_BENEF = Ado_datos2.Recordset!beneficiario_codigo
    
    Set rs_aux3 = New Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "select * from ao_ventas_cabecera where venta_codigo = " & nroventa & " ", db, adOpenKeyset, adLockOptimistic
    If rs_aux3.RecordCount > 0 Then
        VAR_BENEF_VTA = rs_aux3!beneficiario_codigo
        VAR_BENEF_RESP = rs_aux3!beneficiario_codigo_resp
    End If
    
    Set rs_aux4 = New Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "select * from gc_beneficiario where beneficiario_codigo = '" & VAR_BENEF & "'  ", db, adOpenKeyset, adLockOptimistic
    If rs_aux4.RecordCount > 0 Then
        VAR_NIT = rs_aux4!beneficiario_nit
        VAR_RAZON = rs_aux4!beneficiario_denominacion
    End If
    VAR_GLOSA = ""
    Set rs_aux2 = New Recordset
    If rs_aux2.State = 1 Then rs_aux2.Close
    rs_aux2.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & " and estado_codigo = 'REG' AND es_grupo_fac = 'SI' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux2.RecordCount > 0 Then
        'dg_datos2.Visible = True
        db.Execute "update gc_documentos_respaldo set gc_documentos_respaldo.correl_doc = " & nroventa & " Where gc_documentos_respaldo.doc_codigo = '" & rs_aux2!doc_codigo & "' "
        rs_aux2.MoveFirst
        While Not rs_aux2.EOF
            'SI ES CGI o CGE (Falta)
            VAR_GLOSA = VAR_GLOSA + Trim(rs_aux2!cobranza_concepto_plazo) + ". "
            'GRABA DETALLE DE FACTURACION NUEVA (ao_ventas_cobranza)
            db.Execute "INSERT INTO ao_ventas_cobranza (ges_gestion, cobranza_prog_codigo,      venta_codigo,     beneficiario_codigo,       beneficiario_codigo_fac,          beneficiario_codigo_resp,           cobranza_programada_bs,                cobranza_programada_dol,                 cobranza_solicitado_bs,                 cobranza_solicitado_dol,       cobranza_descuento_bs, cobranza_descuento_dol, cobranza_total_bs,              cobranza_total_dol,                       Literal,                  cobranza_fecha_prog,                     cobranza_fecha_cobro,              cobranza_observaciones,     proceso_codigo, subproceso_codigo, etapa_codigo, clasif_codigo, doc_codigo, doc_numero, doc_codigo_fac,         cobranza_nro_factura, cobranza_nro_autorizacion, poa_codigo,  " & _
            " estado_codigo, usr_codigo, fecha_registro, cobranza_fecha_sol, estado_codigo_sol, estado_codigo_fac, venta_codigo_new) " & _
            " VALUES ('" & rs_aux2!ges_gestion & "', " & rs_aux2!cobranza_prog_codigo & ", " & nroventa & ", '" & VAR_BENEF_VTA & "', '" & rs_aux2!beneficiario_codigo & "', '" & VAR_BENEF_RESP & "', " & rs_aux2!cobranza_programada_bs & ", " & rs_aux2!cobranza_programada_dol & ", " & rs_aux2!cobranza_programada_bs & ", " & rs_aux2!cobranza_programada_dol & ",      '0',                '0', " & rs_aux2!cobranza_programada_bs & ", " & rs_aux2!cobranza_programada_dol & ", '" & rs_aux2!Literal & "', '" & rs_aux2!cobranza_fecha_prog & "', '" & rs_aux2!cobranza_fecha_cobro & "', '" & rs_aux2!cobranza_concepto_plazo & "', 'FIN', 'FIN-02',     'FIN-02-02',    'ADM',          'R-105',    '0', '" & rs_aux2!doc_codigo_fac & "', '0',                 '0',                    '3.1.2',  " & _
            " 'REG', '" & glusuario & "', '" & Date & "', '" & Date & "', 'APR', 'REG', '99999')"

            ' APRUEBA ao_ventas_cobranza_prog
            ''db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & Ado_datos16.Recordset!cobranza_prog_codigo & " "
            db.Execute "update ao_ventas_cobranza_prog set estado_codigo = 'APR', fecha_aprueba = '" & Date & "' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & rs_aux2!cobranza_prog_codigo & " "
            ' Actualiza CODIGO_COBRNAZA en el cronogrma
            db.Execute "update ao_ventas_cobranza_prog set ao_ventas_cobranza_prog.cobranza_codigo = ao_ventas_cobranza.cobranza_codigo from ao_ventas_cobranza_prog INNER JOIN ao_ventas_cobranza " & _
            " ON ao_ventas_cobranza_prog.venta_codigo = ao_ventas_cobranza.venta_codigo and ao_ventas_cobranza_prog.cobranza_prog_codigo = ao_ventas_cobranza.cobranza_prog_codigo WHERE (ao_ventas_cobranza_prog.venta_codigo = " & nroventa & " and ao_ventas_cobranza_prog.cobranza_prog_codigo=" & rs_aux2!cobranza_prog_codigo & " )"

            db.Execute "update ao_ventas_cobranza_prog SET fecha_registro = '" & Date & "' Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & rs_aux2!cobranza_prog_codigo & " "
                        
            db.Execute "update ao_ventas_cobranza_prog SET Gestion = YEAR(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & rs_aux2!cobranza_prog_codigo & " "

            db.Execute "update ao_ventas_cobranza_prog SET cobranza_mes = MONTH(cobranza_fecha_prog) Where venta_codigo = " & nroventa & " And cobranza_prog_codigo = " & rs_aux2!cobranza_prog_codigo & " "
            
            'Call ABRIR_TABLAS_AUX
            db.Execute "tp_actualiza_datos_venta " & nroventa
            
            rs_aux2.MoveNext
        Wend
        'GRABA CABECERA DE FACTURACION NUEVA (ao_ventas_cobranza_fac)
        Set rs_aux5 = New Recordset
       If rs_aux5.State = 1 Then rs_aux5.Close
       rs_aux5.Open "select sum(cobranza_total_bs) as totbs2, sum (cobranza_total_dol) as totdl2 from ao_ventas_cobranza where venta_codigo_new = '99999' AND venta_codigo = " & nroventa & " ", db, adOpenKeyset, adLockOptimistic
        If IsNull(rs_aux5!totbs2) Then
            VAR_TOTBS = 0
            VAR_TOTDOL = 0
        Else
            VAR_TOTBS = Round(rs_aux5!totbs2, 2)
            VAR_TOTDOL = Round(rs_aux5!totdl2, 2)
        End If
        var_literal = Literal(CStr(VAR_TOTBS)) + " BOLIVIANOS"
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW

            db.Execute "INSERT INTO ao_ventas_cobranza_fac (ges_gestion, venta_codigo, doc_codigo_fac, beneficiario_codigo_fac,beneficiario_nit, glosa_Descripcion, beneficiario_RazonSocial, nro_dui,  total_bs,     total_dol,                    cambio_oficial, " & _
                        " Importe_ICE, Exportaciones_Exentas, Ventas_tasa_0, Subtotal_ICE, Descuentos_Bonos, Importe_Base_Debito_Fiscal, factura_87_bs,                    factura_87_dol,                  debito_fiscal_13_bs,              debito_fiscal_13_dol,               literal, " & _
                        " clasif_codigo, doc_codigo, doc_numero, factura_impresa, tipo_moneda, cta_codigo, cta_codigo2, correl_contab, estado_fac, estado_codigo_fac, estado_codigo,  " & _
                        " usr_codigo, fecha_registro, edif_codigo_corto, edif_codigo, codigo_empresa ) " & _
            " VALUES ('" & glGestion & "',  " & nroventa & ", '" & Ado_datos2.Recordset!doc_codigo_fac & "', '" & VAR_BENEF & "', '" & VAR_NIT & "', '" & Left(VAR_GLOSA, 245) & "', '" & VAR_RAZON & "',  '0', " & VAR_TOTBS & ",  " & VAR_TOTDOL & ",  " & GlTipoCambioOficial & ",  " & _
                        " '0',          '0',                    '0',            '0',            '0',    " & VAR_TOTBS & ", " & Round(VAR_TOTBS * 0.87, 2) & ", " & Round(VAR_TOTDOL * 0.87, 2) & ", " & Round(VAR_TOTBS * 0.13, 2) & ", " & Round(VAR_TOTDOL * 0.13, 2) & ", '" & var_literal & "',  " & _
                        " 'ADM',        'R-103',        '0',        'N',            'BOB',      'NN',           'NN',        '0',            'REG',      'REG',          'REG',  " & _
                        " '" & glusuario & "', '" & CDate(Date) & "', " & rs_aux3!edif_codigo_corto & ", '" & rs_aux3!edif_codigo & "', " & rs_aux3!codigo_empresa & "  ) "

            'Actualiza CORREO ELECTRONICO
            db.Execute "UPDATE ao_ventas_cobranza_fac SET ao_ventas_cobranza_fac.beneficiario_email  = gc_beneficiario.beneficiario_email FROM ao_ventas_cobranza_fac INNER JOIN gc_beneficiario ON ao_ventas_cobranza_fac.beneficiario_codigo_fac = gc_beneficiario.beneficiario_codigo where ao_ventas_cobranza_fac.beneficiario_email Is Null "
        
            Set rs_aux1 = New ADODB.Recordset
            If rs_aux1.State = 1 Then rs_aux1.Close
            rs_aux1.Open "Select max(IdFactura) as Codigo3 from ao_ventas_cobranza_fac  ", db, adOpenKeyset, adLockOptimistic
            If IsNull(rs_aux1!codigo3) Then
               VAR_IDFAC = 1
            Else
               VAR_IDFAC = rs_aux1!codigo3
            End If
        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        'SE GENERAN CON LA FACTURA (dosifica_autorizacion, nro_factura, fecha_fac, codigo_control, archivo_foto, depto_codigo, Gestion, mes, edif_codigo_corto)
        db.Execute "UPDATE Ao_ventas_cobranza SET venta_codigo_new = " & VAR_IDFAC & " where venta_codigo = " & nroventa & " and  venta_codigo_new = '99999' "
        'estado_codigo1 = 'APR' AND trans_codigo = 'X'
        'GRABA CABECERA DE LA FACTURA (QR)
        db.Execute "INSERT INTO ao_ventas_cobranza_fac_QR (IdFactura, archivo_foto_cargado, estado_codigo, usr_codigo, fecha_registro ) " & _
                " VALUES ('" & VAR_IDFAC & "',  'N',            'REG',   '" & glusuario & "', '" & CDate(Date) & "' ) "

     Else
         dg_datos2.Visible = False
     End If
     MsgBox "Se aprobaron las cuotas y se envió satisfactoriamente la solicitud de Factura ...", , "Atención"
     Unload Me
End Sub

Private Sub btnEliminar_Click()
    If glusuario = "ADMIN" Or glusuario = "FDELGADILLO" Or glusuario = "SQUISPE" Or glusuario = "HMARIN" Or glusuario = "CSALINAS" Or glusuario = "GPALLY" Or glusuario = "CESPINOZA" Then
       If Ado_datos.Recordset.RecordCount > 0 Then
          db.Execute "UPDATE ao_ventas_cobranza_prog SET es_grupo_fac = 'NO' WHERE correl_prog = " & Ado_datos2.Recordset!correl_prog & " "
          Call OptFilGral1_Click
          Call abrir_destino
       End If
    End If
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    db.Execute "UPDATE ao_ventas_cobranza_prog SET es_grupo_fac = 'NO' WHERE venta_codigo = " & NumComp & "  "
    Call OptFilGral1_Click
        Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral1_Click()
    Set rs_datos = New ADODB.Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    rs_datos.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & " and estado_codigo = 'REG' AND es_grupo_fac = 'NO' ", db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos
    'Ado_datos.Recordset.Requery
    If Ado_datos.Recordset.RecordCount > 0 Then
'        FrmCobranza.Visible = True
        Set dg_datos.DataSource = Ado_datos.Recordset
    Else
'        FrmCobranza.Visible = False
    End If
End Sub

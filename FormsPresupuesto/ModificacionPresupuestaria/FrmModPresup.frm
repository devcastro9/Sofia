VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmModPresup 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación - Formulación Presupuestaria"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   12015
   Icon            =   "FrmModPresup.frx":0000
   Moveable        =   0   'False
   ScaleHeight     =   12540
   ScaleWidth      =   17190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H00404040&
      Height          =   1020
      Left            =   120
      Picture         =   "FrmModPresup.frx":0A02
      ScaleHeight     =   960
      ScaleWidth      =   14835
      TabIndex        =   162
      Top             =   120
      Width           =   14900
      Begin VB.CommandButton CmdTransfer 
         BackColor       =   &H00808000&
         Caption         =   "Traspaso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   6000
         Picture         =   "FrmModPresup.frx":6CA34
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Genera una Modificación Presupuestaria"
         Top             =   120
         Width           =   770
      End
      Begin VB.CommandButton BtnAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Aprobar"
         Height          =   720
         Left            =   2640
         Picture         =   "FrmModPresup.frx":6D11E
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Aprueba Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   720
         Left            =   5160
         Picture         =   "FrmModPresup.frx":6D328
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808000&
         Caption         =   "Desapro."
         Height          =   720
         Left            =   2640
         Picture         =   "FrmModPresup.frx":6D76A
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   120
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton BtnBuscar 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         Height          =   720
         Left            =   3480
         Picture         =   "FrmModPresup.frx":6D974
         Style           =   1  'Graphical
         TabIndex        =   169
         ToolTipText     =   "Busca un Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnImprimir 
         BackColor       =   &H00808000&
         Caption         =   "Imprimir"
         Height          =   720
         Left            =   4320
         Picture         =   "FrmModPresup.frx":6DF2C
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Imprime Formulario"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnSalir 
         BackColor       =   &H00808000&
         Caption         =   "Cerrar"
         Height          =   720
         Left            =   6840
         Picture         =   "FrmModPresup.frx":6E4E9
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Cerrar Ventana"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnEliminar 
         BackColor       =   &H00808000&
         Caption         =   "Anular"
         Height          =   720
         Left            =   1800
         Picture         =   "FrmModPresup.frx":6E6F3
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Anula Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnModificar 
         BackColor       =   &H00808000&
         Caption         =   "Modificar"
         Height          =   720
         Left            =   960
         Picture         =   "FrmModPresup.frx":6F3BD
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Modifica Registro Activo"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnAñadir 
         BackColor       =   &H00808000&
         Caption         =   "Nuevo"
         Height          =   720
         Left            =   120
         Picture         =   "FrmModPresup.frx":6F99D
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Nuevo Registro"
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO1"
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
         Left            =   10080
         TabIndex        =   172
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.PictureBox FraGrabarCancelar 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "FrmModPresup.frx":6FFC1
      ScaleHeight     =   915
      ScaleWidth      =   14835
      TabIndex        =   157
      Top             =   120
      Width           =   14900
      Begin VB.CommandButton BtnCancelar 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar"
         Height          =   675
         Left            =   3600
         MaskColor       =   &H00000000&
         Picture         =   "FrmModPresup.frx":DBFF3
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Cancelar"
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton BtnGrabar 
         BackColor       =   &H00808000&
         Caption         =   "Grabar"
         Height          =   675
         Left            =   1560
         Picture         =   "FrmModPresup.frx":DC1FD
         Style           =   1  'Graphical
         TabIndex        =   158
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
         TabIndex        =   160
         Top             =   300
         Width           =   4185
      End
   End
   Begin VB.Frame FraNavega 
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
      Height          =   7830
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   5265
      Begin VB.OptionButton OptIns 
         BackColor       =   &H00000000&
         Caption         =   "OTRA"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2940
         TabIndex        =   151
         Top             =   60
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton OptMin 
         BackColor       =   &H00000000&
         Caption         =   "GERENCIAL"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2940
         TabIndex        =   150
         Top             =   300
         Width           =   1515
      End
      Begin MSAdodcLib.Adodc Adofo_formulacion_gasto 
         Height          =   330
         Left            =   120
         Top             =   7380
         Width           =   5040
         _ExtentX        =   8890
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
         Caption         =   "Formulación"
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
      Begin MSDataGridLib.DataGrid Dgrfo_formulacion_gasto 
         Bindings        =   "FrmModPresup.frx":DC407
         Height          =   6375
         Left            =   120
         TabIndex        =   147
         Top             =   960
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   11245
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
         ColumnCount     =   21
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
            DataField       =   "da_codigo"
            Caption         =   "DA.Codigo"
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
            DataField       =   "dgral_codigo"
            Caption         =   "DGral.Codigo"
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
            DataField       =   "pro_codigo"
            Caption         =   "Proyecto"
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
            Caption         =   "Responsable"
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
            DataField       =   "fte_codigo"
            Caption         =   "Fuente"
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
         BeginProperty Column07 
            DataField       =   "par_codigo"
            Caption         =   "Partida"
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
            DataField       =   "ppto_vigente"
            Caption         =   "Ppto.Vigente"
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
            DataField       =   "ppto_compromiso"
            Caption         =   "Ppto.Compromiso"
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
            DataField       =   "ppto_formulado"
            Caption         =   "Ppto.Formulado"
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
            DataField       =   "ppto_modificaciones"
            Caption         =   "Ppto.Modificaciones"
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
            DataField       =   "ppto_devengado"
            Caption         =   "Ppto.Devengado"
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
            DataField       =   "ppto_pagado"
            Caption         =   "Ppto.Pagado"
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
            DataField       =   "ppto_adiciones"
            Caption         =   "Ppto.Adiciones"
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
            DataField       =   "ppto_devuelto"
            Caption         =   "Ppto.Devuelto"
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
            DataField       =   "ppto_revertido"
            Caption         =   "Ppto.Revertido"
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
            DataField       =   "ppto_anulado"
            Caption         =   "Ppto.Anulado"
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
         BeginProperty Column19 
            DataField       =   "estado_codigo_ppto"
            Caption         =   "Instancia"
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
            DataField       =   "usr_usuario"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   255.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   239.811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column17 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column18 
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column19 
               Object.Visible         =   0   'False
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FORMULACIÓN PRESUPUESTARIA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   315
         Left            =   315
         TabIndex        =   72
         Top             =   635
         Width           =   3930
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFC0&
         BorderWidth     =   2
         X1              =   60
         X2              =   5200
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE RESOLUCIÓN :"
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
         Height          =   240
         Left            =   120
         TabIndex        =   148
         Top             =   180
         Width           =   2550
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE RESOLUCIÓN :"
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
         Left            =   150
         TabIndex        =   149
         Top             =   80
         Width           =   2550
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame FraDatTrans 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7860
      Left            =   5580
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   9345
      Begin VB.Frame Frame3 
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
         Height          =   2790
         Left            =   60
         TabIndex        =   50
         Top             =   60
         Width           =   9255
         Begin VB.TextBox Txtuni_codigo 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2160
            TabIndex        =   57
            Text            =   "UAUDI"
            Top             =   390
            Width           =   1080
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3360
            TabIndex        =   56
            Text            =   "UNIDAD AUDITORIA INTERNA"
            Top             =   390
            Width           =   5355
         End
         Begin VB.TextBox Txtpro_actividad 
            DataField       =   "Pro_actividad"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8640
            TabIndex        =   55
            Top             =   2040
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox TxtPro_programa 
            DataField       =   "Pro_programa"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8640
            TabIndex        =   54
            Top             =   2280
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox Txtfgs_formulado 
            BackColor       =   &H80000014&
            DataField       =   "ppto_formulado"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   285
            Left            =   2160
            TabIndex        =   53
            Text            =   "0"
            Top             =   2280
            Width           =   1200
         End
         Begin VB.TextBox Txtfgs_modificaciones 
            DataField       =   "ppto_modificaciones"
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
            Left            =   4760
            TabIndex        =   52
            Text            =   "0"
            Top             =   2280
            Width           =   1320
         End
         Begin VB.TextBox Txtfgs_vigente 
            DataField       =   "ppto_vigente"
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
            Left            =   7200
            TabIndex        =   51
            Text            =   "0"
            Top             =   2280
            Width           =   1440
         End
         Begin MSDataListLib.DataCombo DtCpar_codigo 
            Bindings        =   "FrmModPresup.frx":DC42D
            DataField       =   "par_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   2160
            TabIndex        =   58
            Top             =   1440
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "par_codigo"
            BoundColumn     =   "Par_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCPar_descripcion_larga 
            Bindings        =   "FrmModPresup.frx":DC465
            DataField       =   "par_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   3360
            TabIndex        =   59
            Top             =   1440
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "Par_descripcion_larga"
            BoundColumn     =   "par_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCFte_codigo 
            Bindings        =   "FrmModPresup.frx":DC47E
            DataField       =   "fte_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   2160
            TabIndex        =   60
            Top             =   720
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "fte_codigo"
            BoundColumn     =   "fte_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCOrg_descripcion 
            Bindings        =   "FrmModPresup.frx":DC49C
            DataField       =   "Org_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   3360
            TabIndex        =   61
            Top             =   1080
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Org_descripcion"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCOrg_codigo 
            Bindings        =   "FrmModPresup.frx":DC4BE
            DataField       =   "Org_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   2160
            TabIndex        =   62
            Top             =   1080
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "Org_codigo"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtCFte_descripcion_larga 
            Bindings        =   "FrmModPresup.frx":DC4EF
            DataField       =   "fte_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   3360
            TabIndex        =   63
            Top             =   720
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Fte_descripcion_larga"
            BoundColumn     =   "fte_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoFte_financia 
            Height          =   330
            Left            =   6915
            Top             =   30
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSAdodcLib.Adodc AdoOrganismo_finan 
            Height          =   330
            Left            =   4890
            Top             =   1155
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
         Begin MSAdodcLib.Adodc Adofc_partida_gasto 
            Height          =   330
            Left            =   7740
            Top             =   1440
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
         Begin MSDataListLib.DataCombo dtc_codigo1 
            Bindings        =   "FrmModPresup.frx":DC50E
            DataField       =   "pro_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   2160
            TabIndex        =   181
            Top             =   1800
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "pro_codigo"
            BoundColumn     =   "pro_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dtc_desc1 
            Bindings        =   "FrmModPresup.frx":DC546
            DataField       =   "pro_codigo"
            DataSource      =   "Adofo_formulacion_gasto"
            Height          =   330
            Left            =   3360
            TabIndex        =   182
            Top             =   1800
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "pro_descripcion"
            BoundColumn     =   "pro_codigo"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label38 
            BackColor       =   &H00808080&
            Caption         =   "     DETALLE DE LA FORMULACIÓN ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   240
            Left            =   60
            TabIndex        =   180
            Top             =   60
            Width           =   9165
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   315
            TabIndex        =   71
            Top             =   410
            Width           =   1545
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financiador :"
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
            Left            =   315
            TabIndex        =   70
            Top             =   1100
            Width           =   1125
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente Financiera:"
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
            Left            =   315
            TabIndex        =   69
            Top             =   740
            Width           =   1620
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto"
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
            Index           =   10
            Left            =   315
            TabIndex        =   68
            Top             =   1850
            Width           =   780
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partida del Gasto:"
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
            Left            =   315
            TabIndex        =   67
            Top             =   1460
            Width           =   1575
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modificación:"
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
            Left            =   3600
            TabIndex        =   66
            Top             =   2300
            Width           =   1140
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vigente:"
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
            Left            =   6460
            TabIndex        =   65
            Top             =   2300
            Width           =   720
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Formulado:"
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
            Left            =   285
            TabIndex        =   64
            Top             =   2300
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   2085
         Left            =   45
         TabIndex        =   18
         Top             =   2840
         Width           =   9255
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Formulado:"
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
            Left            =   300
            TabIndex        =   45
            Top             =   1770
            Width           =   1575
         End
         Begin VB.Label Lblfgs_formuladoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   255
            Left            =   2100
            TabIndex        =   44
            Top             =   1740
            Width           =   1515
         End
         Begin VB.Label Lblpro_actividadO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8760
            TabIndex        =   39
            Top             =   720
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Lblpro_proyectoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2085
            TabIndex        =   38
            Top             =   1350
            Width           =   1080
         End
         Begin VB.Label Lblpro_SubprogramaO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8760
            TabIndex        =   37
            Top             =   1320
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label LblPro_programaO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3300
            TabIndex        =   36
            Top             =   1320
            Width           =   5355
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto"
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
            Index           =   13
            Left            =   315
            TabIndex        =   34
            Top             =   1410
            Width           =   780
         End
         Begin VB.Label Lblfgs_vigenteO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   255
            Left            =   7170
            TabIndex        =   31
            Top             =   1740
            Width           =   1515
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Vigente:"
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
            Left            =   5790
            TabIndex        =   30
            Top             =   1770
            Width           =   1320
         End
         Begin VB.Label Label28 
            BackColor       =   &H00808080&
            Caption         =   "   ORIGEN ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   240
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   9165
         End
         Begin VB.Label LblFte_descripcion_largaO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   27
            Top             =   390
            Width           =   5355
         End
         Begin VB.Label LblOrg_descripcionO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   26
            Top             =   705
            Width           =   5355
         End
         Begin VB.Label LblPar_descripcion_largaO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   25
            Top             =   1020
            Width           =   5355
         End
         Begin VB.Label LblFte_codigoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   24
            Top             =   390
            Width           =   1080
         End
         Begin VB.Label LblOrg_codigoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   23
            Top             =   705
            Width           =   1080
         End
         Begin VB.Label Lblpar_codigoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   22
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente Financiera:"
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
            Left            =   300
            TabIndex        =   21
            Top             =   420
            Width           =   1620
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partida del Gasto:"
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
            Left            =   315
            TabIndex        =   20
            Top             =   1065
            Width           =   1575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
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
            Left            =   300
            TabIndex        =   19
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Lbluni_codigoO 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2550
            TabIndex        =   48
            Top             =   345
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2175
         Left            =   45
         TabIndex        =   7
         Top             =   4920
         Width           =   9270
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Formulado:"
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
            Left            =   300
            TabIndex        =   47
            Top             =   1820
            Width           =   1575
         End
         Begin VB.Label Lblfgs_formuladoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   255
            Left            =   2085
            TabIndex        =   46
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Lblpro_actividadD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8760
            TabIndex        =   43
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Lblpro_proyectoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2085
            TabIndex        =   42
            Top             =   1350
            Width           =   1080
         End
         Begin VB.Label Lblpro_SubprogramaD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   8760
            TabIndex        =   41
            Top             =   1200
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label LblPro_programaD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   3300
            TabIndex        =   40
            Top             =   1350
            Width           =   5355
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto"
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
            Index           =   17
            Left            =   315
            TabIndex        =   35
            Top             =   1395
            Width           =   780
         End
         Begin VB.Label Lblfgs_vigenteD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
            Height          =   255
            Left            =   7120
            TabIndex        =   33
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Vigente:"
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
            Left            =   5760
            TabIndex        =   32
            Top             =   1820
            Width           =   1320
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente Financiera:"
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
            Left            =   315
            TabIndex        =   17
            Top             =   405
            Width           =   1620
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Organismo Finan. :"
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
            Left            =   300
            TabIndex        =   16
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label LblFte_descripcion_largaD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   15
            Top             =   390
            Width           =   5355
         End
         Begin VB.Label LblOrg_descripcionD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   14
            Top             =   705
            Width           =   5355
         End
         Begin VB.Label LblPar_descripcion_largaD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3300
            TabIndex        =   13
            Top             =   1035
            Width           =   5355
         End
         Begin VB.Label LblFte_codigoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   12
            Top             =   390
            Width           =   1080
         End
         Begin VB.Label LblOrg_codigoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   11
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Lblpar_codigoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2085
            TabIndex        =   10
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Partida del Gasto:"
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
            Left            =   315
            TabIndex        =   9
            Top             =   1065
            Width           =   1575
         End
         Begin VB.Label Label25 
            BackColor       =   &H00808080&
            Caption         =   "   DESTINO ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   225
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   9165
         End
         Begin VB.Label Lbluni_codigoD 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2550
            TabIndex        =   49
            Top             =   345
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Frame Framontos 
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
         Height          =   735
         Left            =   45
         TabIndex        =   3
         Top             =   7080
         Width           =   9255
         Begin VB.TextBox TxtNro_resolucionT 
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
            Left            =   2100
            TabIndex        =   0
            Top             =   240
            Width           =   1725
         End
         Begin TDBNumberCtrl.TDBNumber Txtfgs_modificacionesT 
            Height          =   300
            Left            =   6780
            TabIndex        =   1
            Top             =   240
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   529
            _Version        =   65537
            AlignHorizontal =   1
            ClipMode        =   0
            ErrorBeep       =   0   'False
            ReadOnly        =   0   'False
            HighlightText   =   -1  'True
            ZeroAllowed     =   -1  'True
            MinusColor      =   255
            MaxValue        =   999999999
            MinValue        =   -999999999
            Value           =   0
            SelStart        =   1
            SelLength       =   0
            KeyClear        =   "{F2}"
            KeyNext         =   ""
            KeyPopup        =   "{SPACE}"
            KeyPrevious     =   ""
            KeyThreeZero    =   ""
            SepDecimal      =   "."
            SepThousand     =   ","
            Text            =   "0.00"
            Format          =   "###,###,##0.00"
            DisplayFormat   =   ""
            Appearance      =   1
            BackColor       =   -2147483643
            Enabled         =   0   'False
            ForeColor       =   -2147483640
            BorderStyle     =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            DropdownButton  =   0   'False
            SpinButton      =   0   'False
            Caption         =   "&Caption"
            CaptionAlignment=   3
            CaptionColor    =   0
            CaptionWidth    =   2
            CaptionPosition =   0
            CaptionSpacing  =   3
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpinAutowrap    =   0   'False
            _StockProps     =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmModPresup.frx":DC55F
            MousePointer    =   0
         End
         Begin VB.TextBox Txtfgs_modificacionesT1 
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
            Enabled         =   0   'False
            Height          =   285
            Left            =   6780
            TabIndex        =   4
            Top             =   240
            Width           =   1725
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Doc. Respaldo:"
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
            Left            =   300
            TabIndex        =   6
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Modificación:"
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
            Left            =   4920
            TabIndex        =   5
            Top             =   300
            Width           =   1740
         End
      End
   End
   Begin VB.Frame FraModpptoDat 
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
      Height          =   7896
      Left            =   5940
      TabIndex        =   73
      Top             =   1200
      Width           =   8535
      Begin VB.OptionButton OptTipo_resolucion2 
         Caption         =   "MINISTERIAL"
         Height          =   195
         Left            =   3660
         TabIndex        =   153
         Top             =   480
         Width           =   1515
      End
      Begin VB.OptionButton OptTipo_resolucion1 
         Caption         =   "INSTITUCIONAL"
         Height          =   195
         Left            =   1920
         TabIndex        =   152
         Top             =   480
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.Frame FraDES 
         Caption         =   "DESTINO :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3375
         Left            =   60
         TabIndex        =   109
         Top             =   4500
         Visible         =   0   'False
         Width           =   8430
         Begin VB.TextBox Txtuni_codigo_des 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1605
            TabIndex        =   119
            Text            =   "VEIPS"
            Top             =   225
            Width           =   1065
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2685
            TabIndex        =   118
            Text            =   "VICEMINISTERIO DE EDUCACION INICAL, PRIMARIA Y SECUNDARIA"
            Top             =   225
            Width           =   5685
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Proyecto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   7560
            Picture         =   "FrmModPresup.frx":DC57B
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Despliega lista de Proyectos"
            Top             =   1965
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox Txtpro_actividad_des 
            DataField       =   "Pro_actividad"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            Height          =   270
            Left            =   6750
            TabIndex        =   116
            Top             =   1905
            Width           =   450
         End
         Begin VB.TextBox Txtpro_proyecto_des 
            DataField       =   "Pro_proyecto"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            Height          =   270
            Left            =   5280
            TabIndex        =   115
            Top             =   1905
            Width           =   450
         End
         Begin VB.TextBox Txtpro_Subprograma_des 
            DataField       =   "Pro_subprograma"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            Height          =   270
            Left            =   3885
            TabIndex        =   114
            Top             =   1905
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox TxtPro_programa_des 
            DataField       =   "Pro_programa"
            DataSource      =   "AdoDetalle"
            Enabled         =   0   'False
            Height          =   270
            Left            =   2205
            TabIndex        =   113
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox Txtfgs_formulado_des 
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
            Left            =   1890
            TabIndex        =   112
            Text            =   "0"
            Top             =   2850
            Width           =   1440
         End
         Begin VB.TextBox Txtfgs_modificaciones_des 
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
            Left            =   4515
            TabIndex        =   111
            Text            =   "0"
            Top             =   2850
            Width           =   1440
         End
         Begin VB.TextBox Txtfgs_vigente_des 
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
            Left            =   6825
            TabIndex        =   110
            Text            =   "0"
            Top             =   2850
            Width           =   1440
         End
         Begin MSDataListLib.DataCombo DtCpar_codigo_des 
            Bindings        =   "FrmModPresup.frx":DC6C5
            Height          =   315
            Left            =   1605
            TabIndex        =   120
            Top             =   1440
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "par_codigo"
            BoundColumn     =   "par_descripcion_larga"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCPar_descripcion_larga_des 
            Bindings        =   "FrmModPresup.frx":DC70A
            Height          =   315
            Left            =   2685
            TabIndex        =   121
            Top             =   1440
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Par_descripcion_larga"
            BoundColumn     =   "par_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCFte_codigo_des 
            Bindings        =   "FrmModPresup.frx":DC730
            DataField       =   "fte_codigo"
            Height          =   315
            Left            =   1605
            TabIndex        =   122
            Top             =   630
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "fte_codigo"
            BoundColumn     =   "Fte_descripcion_larga"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCOrg_descripcion_des 
            Bindings        =   "FrmModPresup.frx":DC752
            DataField       =   "Org_descripcion"
            Height          =   315
            Left            =   2685
            TabIndex        =   123
            Top             =   1035
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Org_descripcion"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCOrg_codigo_des 
            Bindings        =   "FrmModPresup.frx":DC778
            DataField       =   "Org_codigo"
            Height          =   315
            Left            =   1605
            TabIndex        =   124
            Top             =   1020
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Org_codigo"
            BoundColumn     =   "Org_descripcion"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCFte_descripcion_larga_des 
            Bindings        =   "FrmModPresup.frx":DC7AD
            DataField       =   "Fte_descripcion_larga"
            Height          =   315
            Left            =   2685
            TabIndex        =   125
            Top             =   615
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Fte_descripcion_larga"
            BoundColumn     =   "fte_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc AdoFte_financia_des 
            Height          =   330
            Left            =   3600
            Top             =   540
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSAdodcLib.Adodc AdoOrganismo_finan_des 
            Height          =   330
            Left            =   5745
            Top             =   930
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
         Begin MSAdodcLib.Adodc Adofc_partida_gasto_des 
            Height          =   330
            Left            =   4080
            Top             =   1380
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
         Begin VB.Label Label7 
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
            Left            =   60
            TabIndex        =   139
            Top             =   255
            Width           =   1320
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Organismo Finan. :"
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
            TabIndex        =   138
            Top             =   1095
            Width           =   1530
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fuente Finan. :"
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
            Left            =   60
            TabIndex        =   137
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label Label1_des 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1245
            TabIndex        =   136
            Top             =   2220
            Visible         =   0   'False
            Width           =   6270
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actividad"
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
            Index           =   4
            Left            =   5955
            TabIndex        =   135
            Top             =   1950
            Width           =   750
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programa"
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
            Index           =   5
            Left            =   1365
            TabIndex        =   134
            Top             =   1950
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto"
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
            Index           =   6
            Left            =   4530
            TabIndex        =   133
            Top             =   1950
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubPrograma"
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
            Index           =   7
            Left            =   2745
            TabIndex        =   132
            Top             =   1950
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Categoría Programática:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   75
            TabIndex        =   131
            Top             =   2025
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Partida:"
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
            Left            =   60
            TabIndex        =   130
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Formulado:"
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
            Left            =   930
            TabIndex        =   129
            Top             =   2895
            Width           =   930
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Modificación:"
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
            Left            =   3435
            TabIndex        =   128
            Top             =   2895
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Vigente:"
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
            Left            =   6135
            TabIndex        =   127
            Top             =   2895
            Width           =   690
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "MONTOS:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   126
            Top             =   2745
            Width           =   750
         End
      End
      Begin VB.Frame FraORI 
         Caption         =   "ORIGEN :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3345
         Left            =   60
         TabIndex        =   77
         Top             =   1095
         Width           =   8430
         Begin VB.TextBox Txtfgs_adicion_ori 
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
            Left            =   2880
            TabIndex        =   155
            Text            =   "0"
            Top             =   2940
            Width           =   1440
         End
         Begin VB.TextBox Txtuni_codigo_ori 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1575
            TabIndex        =   93
            Text            =   "VEIPS"
            Top             =   210
            Width           =   1065
         End
         Begin VB.TextBox Txtfgs_vigente_ori 
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
            Left            =   6840
            TabIndex        =   86
            Text            =   "0"
            Top             =   2950
            Width           =   1440
         End
         Begin VB.TextBox Txtfgs_modificaciones_ori 
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
            Left            =   4740
            TabIndex        =   85
            Text            =   "0"
            Top             =   2950
            Width           =   1440
         End
         Begin VB.TextBox Txtfgs_formulado_ori 
            BackColor       =   &H80000018&
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
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   84
            Text            =   "0"
            Top             =   2950
            Width           =   1440
         End
         Begin VB.TextBox TxtPro_programa_ori 
            DataField       =   "Pro_programa"
            DataSource      =   "AdoDetalle"
            Height          =   270
            Left            =   2220
            MaxLength       =   2
            TabIndex        =   83
            Top             =   1905
            Width           =   435
         End
         Begin VB.TextBox Txtpro_Subprograma_ori 
            DataField       =   "Pro_subprograma"
            DataSource      =   "AdoDetalle"
            Height          =   270
            Left            =   3885
            MaxLength       =   2
            TabIndex        =   82
            Top             =   1905
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox Txtpro_proyecto_ori 
            DataField       =   "Pro_proyecto"
            DataSource      =   "AdoDetalle"
            Height          =   270
            Left            =   5280
            MaxLength       =   2
            TabIndex        =   81
            Top             =   1905
            Width           =   450
         End
         Begin VB.TextBox Txtpro_actividad_ori 
            DataField       =   "Pro_actividad"
            DataSource      =   "AdoDetalle"
            Height          =   270
            Left            =   6750
            MaxLength       =   2
            TabIndex        =   80
            Top             =   1905
            Width           =   450
         End
         Begin VB.CommandButton CmdProyecto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Proyecto"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   7560
            Picture         =   "FrmModPresup.frx":DC7D0
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Despliega lista de Proyectos"
            Top             =   1965
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2685
            TabIndex        =   78
            Text            =   "VICEMINISTERIO DE EDUCACION INICAL, PRIMARIA Y SECUNDARIA"
            Top             =   225
            Width           =   5685
         End
         Begin MSDataListLib.DataCombo DtCpar_codigo_ori 
            Bindings        =   "FrmModPresup.frx":DC91A
            Height          =   315
            Left            =   1605
            TabIndex        =   87
            Top             =   1440
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "par_codigo"
            BoundColumn     =   "Par_descripcion_larga"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCPar_descripcion_larga_ori 
            Bindings        =   "FrmModPresup.frx":DC95F
            Height          =   315
            Left            =   2685
            TabIndex        =   88
            Top             =   1440
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Par_descripcion_larga"
            BoundColumn     =   "par_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCFte_codigo_ori 
            Bindings        =   "FrmModPresup.frx":DC985
            Height          =   315
            Left            =   1605
            TabIndex        =   89
            Top             =   630
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "fte_codigo"
            BoundColumn     =   "Fte_descripcion_larga"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCOrg_descripcion_ori 
            Bindings        =   "FrmModPresup.frx":DC9A7
            Height          =   315
            Left            =   2685
            TabIndex        =   90
            Top             =   1035
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Org_descripcion"
            BoundColumn     =   "Org_codigo"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCOrg_codigo_ori 
            Bindings        =   "FrmModPresup.frx":DC9CD
            Height          =   315
            Left            =   1605
            TabIndex        =   91
            Top             =   1035
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Org_codigo"
            BoundColumn     =   "Org_descripcion"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DtCFte_descripcion_larga_ori 
            Bindings        =   "FrmModPresup.frx":DCA02
            Height          =   315
            Left            =   2685
            TabIndex        =   92
            Top             =   630
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "Fte_descripcion_larga"
            BoundColumn     =   "fte_codigo"
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc AdoFte_financia_ori 
            Height          =   330
            Left            =   3600
            Top             =   540
            Visible         =   0   'False
            Width           =   1320
            _ExtentX        =   2328
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
         Begin MSAdodcLib.Adodc AdoOrganismo_finan_ori 
            Height          =   330
            Left            =   2895
            Top             =   975
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
         Begin MSAdodcLib.Adodc Adofc_partida_gasto_ori 
            Height          =   330
            Left            =   4080
            Top             =   1380
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
         Begin MSAdodcLib.Adodc Adofc_unidad_ejecutora_ori 
            Height          =   330
            Left            =   2325
            Top             =   135
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
         Begin MSDataListLib.DataCombo dcbuni_codigo_ori 
            Bindings        =   "FrmModPresup.frx":DCA25
            Height          =   315
            Left            =   1590
            TabIndex        =   94
            Top             =   165
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "fte_codigo"
            BoundColumn     =   "Fte_descripcion_larga"
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
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Adición:"
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
            TabIndex        =   156
            Top             =   2760
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "MONTOS:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   75
            TabIndex        =   108
            Top             =   2745
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Vigente:"
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
            Left            =   6900
            TabIndex        =   107
            Top             =   2760
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Modificación:"
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
            Left            =   4740
            TabIndex        =   106
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Formulado:"
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
            Left            =   975
            TabIndex        =   105
            Top             =   2760
            Width           =   930
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Partida:"
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
            Left            =   60
            TabIndex        =   104
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Categoría Programática:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   75
            TabIndex        =   103
            Top             =   1980
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SubPrograma"
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
            Index           =   3
            Left            =   2745
            TabIndex        =   102
            Top             =   1950
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proyecto"
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
            Index           =   2
            Left            =   4530
            TabIndex        =   101
            Top             =   1950
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programa"
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
            Index           =   1
            Left            =   1365
            TabIndex        =   100
            Top             =   1950
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actividad"
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
            Index           =   0
            Left            =   5955
            TabIndex        =   99
            Top             =   1950
            Width           =   750
         End
         Begin VB.Label Label1_ori 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1245
            TabIndex        =   98
            Top             =   2220
            Visible         =   0   'False
            Width           =   6270
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Fuente Finan. :"
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
            Left            =   60
            TabIndex        =   97
            Top             =   690
            Width           =   1185
         End
         Begin VB.Label LblCod_Poa 
            AutoSize        =   -1  'True
            Caption         =   "Organismo Finan. :"
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
            TabIndex        =   96
            Top             =   1095
            Width           =   1530
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
            Left            =   60
            TabIndex        =   95
            Top             =   255
            Width           =   1320
         End
      End
      Begin VB.TextBox TxtNro_resolucion 
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
         Left            =   1755
         TabIndex        =   75
         Top             =   795
         Width           =   1725
      End
      Begin VB.ComboBox CmbTipo_modificacion 
         Height          =   315
         ItemData        =   "FrmModPresup.frx":DCA47
         Left            =   6765
         List            =   "FrmModPresup.frx":DCA4E
         TabIndex        =   74
         Top             =   750
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker DTPFecha_Ingreso 
         DataField       =   "fecha_ingreso"
         Height          =   285
         Left            =   6840
         TabIndex        =   76
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   57344001
         CurrentDate     =   36541
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE RESOLUCIÓN :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   154
         Top             =   480
         Width           =   2070
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   3405
         Left            =   60
         Picture         =   "FrmModPresup.frx":DCA5B
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   8430
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
         Left            =   105
         TabIndex        =   146
         Top             =   165
         Width           =   1560
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
         Left            =   2940
         TabIndex        =   145
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Modificación:"
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
         Left            =   4920
         TabIndex        =   144
         Top             =   165
         Width           =   1860
      End
      Begin VB.Label Lblcodigo_mod_ppto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "correlativo_ingreso"
         Height          =   255
         Left            =   1680
         TabIndex        =   143
         Top             =   120
         Width           =   735
      End
      Begin VB.Label LblGes_Gestion 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2000"
         DataField       =   "ges_gestion"
         Height          =   255
         Left            =   3720
         TabIndex        =   142
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Resolución :"
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
         Left            =   135
         TabIndex        =   141
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Comprobante :"
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
         Left            =   4815
         TabIndex        =   140
         Top             =   795
         Width           =   1890
      End
   End
   Begin Crystal.CrystalReport Cry 
      Left            =   1560
      Top             =   9840
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
   Begin VB.PictureBox FraCmdTrans 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      Picture         =   "FrmModPresup.frx":DD90F
      ScaleHeight     =   915
      ScaleWidth      =   14835
      TabIndex        =   174
      Top             =   120
      Width           =   14900
      Begin VB.CommandButton CmdTransNoTot 
         BackColor       =   &H00808000&
         Caption         =   "Cancelar TODO"
         Height          =   720
         Left            =   3600
         Picture         =   "FrmModPresup.frx":149941
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   120
         Width           =   750
      End
      Begin VB.CommandButton CmdTransOk 
         BackColor       =   &H00808000&
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   2760
         Picture         =   "FrmModPresup.frx":14A80B
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   120
         Width           =   765
      End
      Begin VB.CommandButton CmdBuscar1 
         BackColor       =   &H00808000&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1920
         Picture         =   "FrmModPresup.frx":14AB15
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   120
         Width           =   770
      End
      Begin VB.CommandButton CmdTransDes 
         BackColor       =   &H00808000&
         Caption         =   "Destino"
         Height          =   720
         Left            =   1080
         MousePointer    =   4  'Icon
         Picture         =   "FrmModPresup.frx":14AD1F
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   120
         Width           =   770
      End
      Begin VB.CommandButton CmdTransOri 
         BackColor       =   &H00808000&
         Caption         =   "Origen"
         Height          =   720
         Left            =   240
         MousePointer    =   4  'Icon
         Picture         =   "FrmModPresup.frx":14B409
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   120
         Width           =   770
      End
      Begin VB.Label Label50 
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
         Left            =   9900
         TabIndex        =   175
         Top             =   300
         Width           =   1305
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   120
      Top             =   9120
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
   Begin MSAdodcLib.Adodc Adofo_cmbte_mod_ppto 
      Height          =   330
      Left            =   2640
      Top             =   9120
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
      Caption         =   "Adofo_cmbte_mod_ppto"
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
   End
End
Attribute VB_Name = "FrmModPresup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Sistema:                  ADFIN-2002 / FE
' Módulo:                   Momdificación Presupuestaria de ModPpto
' Base de Datos:            SQL SERVER 7.0 (español)
' Formulario :              FrmModPpto.frm
' Descipción :              Registro de ModPpto Presupuestarios
' Formularios relacionados: MainMenu.frm (Padre)
'                           ComproModPpto.rpt (Crystal Reports ver. 7.0)
' Versión:                  2.0
'========================================================================================

Option Explicit
'==== recordset princ
Dim rstfo_cmbte_mod_ppto As New ADODB.Recordset
Dim rstfc_cmbte_mod_ppto_Correl As New ADODB.Recordset
Dim rstfc_unidad_ejecutora_ori As New ADODB.Recordset
Dim rstfc_unidad_ejecutora_des As New ADODB.Recordset

Dim rstFte_financia_ori As New ADODB.Recordset
Dim rstFte_financia_des As New ADODB.Recordset
Dim rstOrganismo_finan_ori As New ADODB.Recordset
Dim rstOrganismo_finan_des As New ADODB.Recordset
Dim rstfc_partida_gasto_ori As New ADODB.Recordset
Dim rstfc_partida_gasto_des As New ADODB.Recordset
Dim rstdestino As New ADODB.Recordset
Dim rstfo_cmbte_mod_ppto_rep As New ADODB.Recordset
'==== recordset fo_formulacion_gasto
  Dim rstfo_formulacion_gasto As New ADODB.Recordset
  Dim rstTipo_moneda As New ADODB.Recordset
  Dim rstFte_financia As New ADODB.Recordset
  Dim rstOrganismo_finan As New ADODB.Recordset
  
  Dim rs_datos1 As New ADODB.Recordset

'==== variables
Dim correlativo1 As Integer
Dim sino As String
Dim swgraba As Integer
Dim v_añadir As Integer
Dim marca1 As BookmarkEnum
Dim swcopiar As Integer
Dim V_accion As String
Dim ges_gestion1 As String
Dim swmodificar As Integer
Dim codigo_mod_ppto1 As Integer

Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim EntrarAdo As Boolean 'Para que al aprobar no muestre uno por uno
Dim queryinicial As String
Dim PosibleApliqueFiltro As Boolean
Dim msgSalir As String
Dim swvalida_trans As Integer
Dim swigual As Integer
Dim fgs_vigente1 As Double
Dim v1, amod

'Dim ClBuscaGrid As  ClBuscaEnGridExterno

Private Sub Adofo_cmbte_mod_ppto_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Call asigna

  If (Not Adofo_cmbte_mod_ppto.Recordset.BOF) And (Not Adofo_cmbte_mod_ppto.Recordset.EOF) Then
'    Adofo_cmbte_mod_ppto.Recordset.MoveFirst

    If Adofo_cmbte_mod_ppto.Recordset("estado_aprobacion") = "S" Or Adofo_cmbte_mod_ppto.Recordset("estado_aprobacion") = "E" Then
      BtnEliminar.Enabled = False
      BtnAprobar.Enabled = False
      BtnModificar.Enabled = False
      BtnEliminar.Enabled = False
    Else
      BtnEliminar.Enabled = False
      BtnAprobar.Enabled = True
      BtnModificar.Enabled = True
      BtnEliminar.Enabled = True
    End If

    '===== origen
    TxtNro_resolucion = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Nro_resolucion")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("Nro_resolucion"))
    Lblcodigo_mod_ppto = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("codigo_mod_ppto")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("codigo_mod_ppto"))
    LblGes_Gestion = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Ges_Gestion")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("Ges_Gestion"))
    If Adofo_cmbte_mod_ppto.Recordset!tipo_resolucion = "INS" Then
      OptTipo_resolucion1.Value = True
    End If
    If Adofo_cmbte_mod_ppto.Recordset!tipo_resolucion = "MIN" Then
      OptTipo_resolucion2.Value = True
    End If
    Select Case Adofo_cmbte_mod_ppto.Recordset("Tipo_modificacion")
      Case "A"
        CmbTipo_modificacion.ListIndex = 0
        CmbTipo_modificacion.Enabled = True
        FraDES.Visible = False
'      Case "R"
'        CmbTipo_modificacion.ListIndex = 1
'        CmbTipo_modificacion.Enabled = True
'        FraDES.Visible = False
      Case "T"
        CmbTipo_modificacion.Text = "TRASNFERENCIA"
        CmbTipo_modificacion.Enabled = False
        FraDES.Visible = True
    End Select
    
    Txtuni_codigo_ori = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("uni_codigo_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("uni_codigo_ori"))
    DtCFte_codigo_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fte_codigo_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("fte_codigo_ori"))
    DtCFte_descripcion_larga_ori.Text = DtCFte_codigo_ori.BoundText
    DtCOrg_codigo_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("org_codigo_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("org_codigo_ori"))
    DtCOrg_descripcion_ori.Text = DtCOrg_codigo_ori.BoundText
    DtCpar_codigo_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("par_codigo_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("par_codigo_ori"))
    DtCPar_descripcion_larga_ori.Text = DtCpar_codigo_ori.BoundText
    TxtPro_programa_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Pro_programa_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("Pro_programa_ori"))
'    Txtpro_Subprograma_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Pro_subprograma_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("Pro_subprograma_ori"))
    Txtpro_proyecto_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("pro_proyecto_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("pro_proyecto_ori"))
    Txtpro_actividad_ori.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("pro_actividad_ori")) = True, "", Adofo_cmbte_mod_ppto.Recordset("pro_actividad_ori"))
    Txtfgs_vigente_ori = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_vigente_ori")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_vigente_ori"))
    Txtfgs_adicion_ori = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset!fgs_adicion_ori) = True, 0, Adofo_cmbte_mod_ppto.Recordset!fgs_adicion_ori)
    Txtfgs_modificaciones_ori = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_modificaciones_ori")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_modificaciones_ori"))
    Txtfgs_formulado_ori = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_formulado_ori")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_formulado_ori"))
    '===== destino
    Txtuni_codigo_des = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("uni_codigo_des")) = True, "", Adofo_cmbte_mod_ppto.Recordset("uni_codigo_des"))
    DtCFte_codigo_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fte_codigo_des")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("fte_codigo_des"))
    DtCFte_descripcion_larga_des.Text = DtCFte_codigo_des.BoundText
    DtCOrg_codigo_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("org_codigo_des")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("org_codigo_des"))
    DtCOrg_descripcion_des.Text = DtCOrg_codigo_des.BoundText
    DtCpar_codigo_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("par_codigo_des")) = True, " ", Adofo_cmbte_mod_ppto.Recordset("par_codigo_des"))
    DtCPar_descripcion_larga_des.Text = DtCpar_codigo_des.BoundText
    TxtPro_programa_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Pro_programa_des")) = True, "", Adofo_cmbte_mod_ppto.Recordset("Pro_programa_des"))
'    Txtpro_Subprograma_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("Pro_subprograma_des")) = True, "", Adofo_cmbte_mod_ppto.Recordset("Pro_subprograma_des"))
    Txtpro_proyecto_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("pro_proyecto_des")) = True, "", Adofo_cmbte_mod_ppto.Recordset("pro_proyecto_des"))
    Txtpro_actividad_des.Text = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("pro_actividad_des")) = True, "", Adofo_cmbte_mod_ppto.Recordset("pro_actividad_des"))
    Txtfgs_vigente_des = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_vigente_des")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_vigente_des"))
    Txtfgs_modificaciones_des = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_modificaciones_des")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_modificaciones_des"))
    Txtfgs_formulado_des = IIf(IsNull(Adofo_cmbte_mod_ppto.Recordset("fgs_formulado_des")) = True, 0, Adofo_cmbte_mod_ppto.Recordset("fgs_formulado_des"))

    'Call activar_Obj
    'Call desactivar_Obj
  End If

End Sub

Private Sub Adofo_formulacion_gasto_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If (Not Adofo_formulacion_gasto.Recordset.BOF) And (Not Adofo_formulacion_gasto.Recordset.EOF) Then
    
    Adofo_formulacion_gasto.Caption = "Registro N°: " & Adofo_formulacion_gasto.Recordset.AbsolutePosition & " de " & Adofo_formulacion_gasto.Recordset.RecordCount
    
'    Txtuni_codigo = IIf(IsNull(Adofo_formulacion_gasto.Recordset("uni_codigo")) = True, "", Adofo_formulacion_gasto.Recordset("uni_codigo"))
    
'    DtCFte_codigo.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("fte_codigo")) = True, "", Adofo_formulacion_gasto.Recordset("fte_codigo"))
'    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
'    DtCOrg_codigo.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("org_codigo")) = True, "", Adofo_formulacion_gasto.Recordset("org_codigo"))
'    DtCOrg_descripcion.Text = DtCOrg_codigo.BoundText
'    DtCpar_codigo.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("par_codigo")) = True, "", Adofo_formulacion_gasto.Recordset("par_codigo"))
'    DtCPar_descripcion_larga.Text = DtCpar_codigo.BoundText
'    TxtPro_programa.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("Pro_programa")) = True, "", Adofo_formulacion_gasto.Recordset("Pro_programa"))
''    Txtpro_Subprograma.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("Pro_subprograma")) = True, "", Adofo_formulacion_gasto.Recordset("Pro_subprograma"))
'    dtc_codigo1.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("pro_proyecto")) = True, "", Adofo_formulacion_gasto.Recordset("pro_proyecto"))
'    Txtpro_actividad.Text = IIf(IsNull(Adofo_formulacion_gasto.Recordset("pro_actividad")) = True, "", Adofo_formulacion_gasto.Recordset("pro_actividad"))
'    Txtfgs_vigente = IIf(IsNull(Adofo_formulacion_gasto.Recordset("fgs_vigente")) = True, "", Adofo_formulacion_gasto.Recordset("fgs_vigente"))
'    Txtfgs_modificaciones = IIf(IsNull(Adofo_formulacion_gasto.Recordset("fgs_modificaciones")) = True, "", Adofo_formulacion_gasto.Recordset("fgs_modificaciones"))
'    Txtfgs_formulado = IIf(IsNull(Adofo_formulacion_gasto.Recordset("fgs_formulado")) = True, "", Adofo_formulacion_gasto.Recordset("fgs_formulado"))
  End If

End Sub

Private Sub BtnAñadir_Click()
'===== Proceso para Añadir y/o modificar Datos
    v_añadir = 1
    V_accion = "NORMAL"
'    FraModpptoNav.Enabled = False
    FraModpptoDat.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    FraDES.Visible = False
            
    ' blanquear
    '===== origen
    Lblcodigo_mod_ppto = ""
    LblGes_Gestion = Year(Date)
    DTPFecha_Ingreso = Date
    TxtNro_resolucion.Text = ""
    CmbTipo_modificacion.Text = ""
    CmbTipo_modificacion.Enabled = True
'    Txtuni_codigo_ori.Text = ""
    DtCFte_codigo_ori.Text = ""
    DtCFte_descripcion_larga_ori.Text = ""
    DtCOrg_codigo_ori.Text = ""
    DtCOrg_descripcion_ori.Text = ""
    DtCpar_codigo_ori.Text = ""
    DtCPar_descripcion_larga_ori.Text = ""
    TxtPro_programa_ori.Text = ""
'    Txtpro_Subprograma_ori = ""
    Txtpro_proyecto_ori.Text = ""
    Txtpro_actividad_ori.Text = ""
    Txtfgs_formulado_ori = 0
    Me.Txtfgs_adicion_ori = 0
    Txtfgs_modificaciones_ori = 0
    Txtfgs_vigente_ori = 0
'    Txtfgs_formulado_ori.Enabled = True
    Txtfgs_modificaciones_ori.Enabled = False
    Txtfgs_vigente_ori.Enabled = False
    DtCFte_codigo_ori.Enabled = True
    FraORI.Enabled = True
    
            
End Sub

Private Sub BtnAprobar_Click()
  Dim rstfo_formulacion_gasto As New ADODB.Recordset
  Set rstfo_formulacion_gasto = New ADODB.Recordset
  If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
  If Adofo_cmbte_mod_ppto.Recordset("tipo_modificacion") = "A" Then
    rstfo_formulacion_gasto.Open "select * from fo_ppto_formulacion_gasto where ges_gestion = '" & LblGes_Gestion & "' and uni_codigo = '" & Txtuni_codigo_ori.Text & "' and Pro_programa = '" & TxtPro_programa_ori.Text & "' and pro_proyecto = '" & Txtpro_proyecto_ori.Text & "' and pro_actividad = '" & Txtpro_actividad_ori.Text & "' and Fte_codigo = '" & DtCFte_codigo_ori.Text & "' and Org_codigo = '" & DtCOrg_codigo_ori.Text & "' and par_codigo ='" & DtCpar_codigo_ori.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rstfo_formulacion_gasto.RecordCount < 1 Then
      db.BeginTrans
      rstfo_formulacion_gasto.AddNew
      rstfo_formulacion_gasto("ges_gestion") = LblGes_Gestion
      rstfo_formulacion_gasto("uni_codigo") = Txtuni_codigo_ori.Text
      rstfo_formulacion_gasto("pro_programa") = TxtPro_programa_ori.Text
'      rstfo_formulacion_gasto("pro_subprograma") = Txtpro_Subprograma_ori.Text
      rstfo_formulacion_gasto("pro_proyecto") = Txtpro_proyecto_ori.Text
      rstfo_formulacion_gasto("pro_actividad") = Txtpro_actividad_ori.Text
      rstfo_formulacion_gasto("fte_codigo") = DtCFte_codigo_ori.Text
      rstfo_formulacion_gasto("org_codigo") = DtCOrg_codigo_ori.Text
      rstfo_formulacion_gasto("par_codigo") = DtCpar_codigo_ori.Text
      rstfo_formulacion_gasto("ent_codigo") = Adofo_cmbte_mod_ppto.Recordset("ent_codigo_ori")
      rstfo_formulacion_gasto("fgs_formulado") = CDbl(Txtfgs_formulado_ori)
      rstfo_formulacion_gasto("fgs_modificaciones") = CDbl(Txtfgs_modificaciones_ori)
      rstfo_formulacion_gasto("fgs_vigente") = CDbl(Txtfgs_vigente_ori)
      rstfo_formulacion_gasto!fgs_adicion = CDbl(Txtfgs_adicion_ori)
      rstfo_formulacion_gasto("fgs_compromiso") = 0
      rstfo_formulacion_gasto("fgs_devengado") = 0
      rstfo_formulacion_gasto("fgs_pagado") = 0
      rstfo_formulacion_gasto("fgs_acum_dev") = 0
      rstfo_formulacion_gasto("fgs_acum_rev") = 0
      rstfo_formulacion_gasto("fgs_acum_anl") = 0
      rstfo_formulacion_gasto("fecha_registro") = Format(Date, "dd/mm/yyyy")
      rstfo_formulacion_gasto("hora_registro") = Format(Time, "hh/mm/ss")
      rstfo_formulacion_gasto("usr_usuario") = GlUsuario
      rstfo_formulacion_gasto.Update
      
      rstfo_cmbte_mod_ppto("estado_aprobacion") = "S"
      rstfo_cmbte_mod_ppto.Update
      
      db.CommitTrans
    Else
      MsgBox "La Estructura Presupuestaria ya existe", vbCritical + vbYesNo, "ERROR en la Creación de Estructura Presupuestaria ..."
'      sino = MsgBox("La Estructura Presupuestaria ya existe, ¿ Desea proceder a acumular el monto adicional ? ", vbCritical + vbYesNo, "Creación de Estructura Presupuestaria ...")
'      If sino = vbYes Then
'        db.BeginTrans
'        rstfo_formulacion_gasto!fgs_adicion = rstfo_formulacion_gasto!fgs_adicion + CDbl(Me.Txtfgs_adicion_ori)
'        rstfo_formulacion_gasto!fgs_vigente = rstfo_formulacion_gasto!fgs_vigente + (rstfo_formulacion_gasto!fgs_adicion)
'        rstfo_formulacion_gasto.Update
'
'        rstfo_cmbte_mod_ppto("estado_aprobacion") = "S"
'        rstfo_cmbte_mod_ppto.Update
'
'        db.CommitTrans
'      End If
    End If
    If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
  End If
  
  If Adofo_cmbte_mod_ppto.Recordset("tipo_modificacion") = "T" Then
    If Txtfgs_vigente_ori < 0 Then
      MsgBox "El Monto Vigente no puede ser negativo, por favor cambie el monto a modificar.", vbInformation + vbOKOnly, "Error Monto a cambiar, muy grande..."
      Exit Sub
    End If
    
'==
    If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    rstfo_formulacion_gasto.Open "select * from fo_ppto_formulacion_gasto where pro_programa = '" & TxtPro_programa_ori.Text & "' and pro_proyecto = '" & Txtpro_proyecto_ori.Text & "' and pro_actividad = '" & Txtpro_actividad_ori.Text & "' and fte_codigo = '" & DtCFte_codigo_ori.Text & "' and org_codigo = '" & DtCOrg_codigo_ori.Text & "' and par_codigo = '" & DtCpar_codigo_ori.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rstfo_formulacion_gasto.RecordCount > 0 Then
      Dim fgs_modificaciones1 As Double
      amod = (rstfo_formulacion_gasto!fgs_modificaciones + CDbl(Txtfgs_modificaciones_ori))
      v1 = (rstfo_formulacion_gasto!fgs_formulado + rstfo_formulacion_gasto!fgs_adicion + amod)
      If v1 < rstfo_formulacion_gasto!FGS_compromiso Then
        MsgBox "La modificación no puede ser aprobada, debido a que el monto que se quiere restar es mayor al monto comprometido", vbOKOnly + vbCritical, "Error al realizar la modificación..."
        If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
        Exit Sub
      End If
      'fgs_modificaciones1 = rstfo_formulacion_gasto("fgs_modificaciones") + CDbl(Txtfgs_modificaciones_ori)
      'fgs_vigente1 = rstfo_formulacion_gasto("fgs_formulado") - (IIf(fgs_modificaciones1 < 0, fgs_modificaciones1 * -1, fgs_modificaciones1))
      'If fgs_vigente1 < 0 Then
      ' MsgBox "La modificación no puede ser aprobada, debido a que el monto que se quiere restar es mayor al monto vigente en la tabla actualizada", vbOKOnly + vbCritical, "Error al realizar la modificación..."
      ' If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
      ' Exit Sub
      'End If
    End If
    'If valida_ppto(Me.DtCOrg_codigo_ori, Me.DtCpar_codigo_ori, Me.TxtPro_programa_ori, Txtpro_Subprograma_ori, Txtpro_proyecto_ori, Txtpro_actividad_ori) = 0 Then
    ' Exit Sub
    'End If
'==
    db.BeginTrans
    If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    rstfo_formulacion_gasto.Open "select * from fo_ppto_formulacion_gasto where pro_programa = '" & TxtPro_programa_ori.Text & "' and pro_proyecto = '" & Txtpro_proyecto_ori.Text & "' and pro_actividad = '" & Txtpro_actividad_ori.Text & "' and fte_codigo = '" & DtCFte_codigo_ori.Text & "' and org_codigo = '" & DtCOrg_codigo_ori.Text & "' and par_codigo = '" & DtCpar_codigo_ori.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rstfo_formulacion_gasto.RecordCount > 0 Then
      'rstfo_formulacion_gasto("fgs_formulado") = CDbl(Txtfgs_formulado_ori)
      rstfo_formulacion_gasto("fgs_modificaciones") = amod ' NOW rstfo_formulacion_gasto("fgs_modificaciones") - CDbl(Txtfgs_modificaciones_ori)
      rstfo_formulacion_gasto("fgs_vigente") = v1 'NOW rstfo_formulacion_gasto("fgs_formulado") - (IIf(rstfo_formulacion_gasto("fgs_modificaciones") < 0, rstfo_formulacion_gasto("fgs_modificaciones") * -1, rstfo_formulacion_gasto("fgs_modificaciones")))
      'rstfo_formulacion_gasto("fgs_compromiso") = 0
      'rstfo_formulacion_gasto("fgs_devengado") = 0
      'rstfo_formulacion_gasto("fgs_pagado") = 0
      'rstfo_formulacion_gasto("fgs_acum_dev") = 0
      'rstfo_formulacion_gasto("fgs_acum_rev") = 0
      'rstfo_formulacion_gasto("fgs_acum_anl") = 0
      rstfo_formulacion_gasto("fecha_registro") = Format(Date, "dd/mm/yyyy")
      rstfo_formulacion_gasto("hora_registro") = Format(Time, "hh/mm/ss")
      rstfo_formulacion_gasto("usr_usuario") = GlUsuario
      rstfo_formulacion_gasto.Update
      If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    End If
    If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    rstfo_formulacion_gasto.Open "select * from fo_ppto_formulacion_gasto where pro_programa = '" & TxtPro_programa_des.Text & "' and pro_proyecto = '" & Txtpro_proyecto_des.Text & "' and pro_actividad = '" & Txtpro_actividad_des.Text & "' and fte_codigo = '" & DtCFte_codigo_des.Text & "' and org_codigo = '" & DtCOrg_codigo_des.Text & "' and par_codigo = '" & DtCpar_codigo_des.Text & "'", db, adOpenKeyset, adLockOptimistic
    If rstfo_formulacion_gasto.RecordCount > 0 Then
      'rstfo_formulacion_gasto("fgs_formulado") = CDbl(Txtfgs_formulado_ori)
      amod = 0
      v1 = 0
      amod = (rstfo_formulacion_gasto!fgs_modificaciones + CDbl(Txtfgs_modificaciones_des))
      v1 = (rstfo_formulacion_gasto!fgs_formulado + amod)

      rstfo_formulacion_gasto!fgs_modificaciones = amod 'NOW rstfo_formulacion_gasto!fgs_modificaciones + CDbl(Txtfgs_modificaciones_des)
      rstfo_formulacion_gasto!FGS_VIGENTE = v1 ' NOW rstfo_formulacion_gasto!fgs_formulado + rstfo_formulacion_gasto!fgs_modificaciones  'rstfo_formulacion_gasto("fgs_formulado") + (rstfo_formulacion_gasto("fgs_modificaciones") + CDbl(Txtfgs_modificaciones_des))
      'rstfo_formulacion_gasto("fgs_compromiso") = 0
      'rstfo_formulacion_gasto("fgs_devengado") = 0
      'rstfo_formulacion_gasto("fgs_pagado") = 0
      'rstfo_formulacion_gasto("fgs_acum_dev") = 0
      'rstfo_formulacion_gasto("fgs_acum_rev") = 0
      'rstfo_formulacion_gasto("fgs_acum_anl") = 0
      rstfo_formulacion_gasto("fecha_registro") = Format(Date, "dd/mm/yyyy")
      rstfo_formulacion_gasto("hora_registro") = Format(Time, "hh/mm/ss")
      rstfo_formulacion_gasto("usr_usuario") = GlUsuario
      rstfo_formulacion_gasto.Update
      If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    End If
      codigo_mod_ppto1 = rstfo_cmbte_mod_ppto("codigo_mod_ppto")
      rstfo_cmbte_mod_ppto("estado_aprobacion") = "S"
      rstfo_cmbte_mod_ppto("fecha_registro_aprueba") = Format(Date, "dd/mm/yyyy")
      rstfo_cmbte_mod_ppto("hora_registro_aprueba") = Format(Time, "hh/mm/ss")
      rstfo_cmbte_mod_ppto("usr_usuario_aprueba") = GlUsuario
      
      rstfo_cmbte_mod_ppto.Update
'      rstfo_cmbte_mod_ppto.Requery
      rstfo_cmbte_mod_ppto.Find "codigo_mod_ppto = " & codigo_mod_ppto1, , adSearchForward, 1
      If rstfo_cmbte_mod_ppto.EOF Then rstfo_cmbte_mod_ppto.MoveLast
    db.CommitTrans
  End If
  rstfo_cmbte_mod_ppto.Requery

End Sub

Private Sub BtnEliminar_Click()
' ===== Proceso para confirmar el eliminado de registros
  v_añadir = 3
  sino = MsgBox("¿Está seguro de ANULAR este registro?", vbYesNo + vbQuestion, "Atención...")
  If sino = vbYes Then
    'Call elimina
    Call errado
  End If
End Sub

Private Sub BtnBuscar_Click()
'JQA
'  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'
'  Dim ClBuscaSec As  ClBuscaSecuencialEnRS
'  PosibleApliqueFiltro = False
'  Dim RSNADA As ADODB.Recordset
'  Dim GrSqlAux As String
'
'  Set ClBuscaGrid = New  ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.EsTdbGrid = False
'  Set ClBuscaGrid.GridTrabajo = DtGIngresos  'DtGIngresos
'  ClBuscaGrid.QueryUtilizado = queryinicial
'  Set ClBuscaGrid.RecordsetTrabajo = Adofo_cmbte_mod_ppto.Recordset
'  ClBuscaGrid.CamposVisibles = "110"
'  ClBuscaGrid.Ejecutar
'  PosibleApliqueFiltro = True
'JQA
End Sub

Private Sub CmdBuscar1_Click()
'JQA
'  Dim ClVBusca As  ClBuscaEnGridPropio 'Componente de busquedas
'
'  Dim ClBuscaSec As  ClBuscaSecuencialEnRS
'  PosibleApliqueFiltro = False
'  Dim RSNADA As ADODB.Recordset
'  Dim GrSqlAux As String
'
'  Set ClBuscaGrid = New  ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.EsTdbGrid = False
'  Set ClBuscaGrid.GridTrabajo = Dgrfo_formulacion_gasto  'DtGIngresos
'  ClBuscaGrid.QueryUtilizado = queryinicial
'  Set ClBuscaGrid.RecordsetTrabajo = Adofo_formulacion_gasto.Recordset
'  ClBuscaGrid.CamposVisibles = "110"
'  ClBuscaGrid.Ejecutar
'  PosibleApliqueFiltro = True
'JQA
End Sub

Private Sub BtnCancelar_Click()
'===== Ini cancela actualizaciones ==========
  FraGrabarCancelar.Visible = False
  fraOpciones.Visible = True
'  FraModpptoNav.Enabled = True
  FraModpptoDat.Enabled = False
  rstfo_cmbte_mod_ppto.Requery
  v_añadir = 0
End Sub

Private Sub BtnGrabar_Click()
'======= Ini grabado de datos
  If V_accion = "NORMAL" Then
    swgraba = 0
    Call valida
  End If
  If V_accion = "TRANSFERENCIA" Then
    swgraba = 1
  End If
  If swgraba = 1 Then
    FraGrabarCancelar.Visible = False
    fraOpciones.Visible = True
'    FraModpptoNav.Enabled = True
    FraModpptoDat.Enabled = False
    Set rstdestino = New ADODB.Recordset
    db.BeginTrans
    If v_añadir = 1 Then
      'Call add_correl
      Dim rstcorrModPpto As New ADODB.Recordset
      Set rstcorrModPpto = New ADODB.Recordset
      If rstcorrModPpto.State = 1 Then rstcorrModPpto.Close
      rstcorrModPpto.Open "select * from fc_cmbte_mod_ppto_correl", db, adOpenDynamic, adLockOptimistic ' where org_codigo
      If (rstcorrModPpto.EOF) Then
      'rstcorrModPpto.Find "org_codigo = '" & (DtCOrg_codigo.Text) & "' ", , adSearchForward
      'If rstcorrModPpto.EOF Then
        rstcorrModPpto.AddNew
        'rstcorrModPpto("org_codigo") = Trim(DtCOrg_codigo.Text)
        'rstcorrModPpto("ges_gestion") = Trim(LblGes_Gestion.Caption)
        rstcorrModPpto("codigo_mod_ppto") = 1
        rstcorrModPpto.Update
        correlativo1 = rstcorrModPpto("codigo_mod_ppto")
        FrmModPresup.Lblcodigo_mod_ppto.Caption = rstcorrModPpto("codigo_mod_ppto")
      Else
        rstcorrModPpto.MoveFirst
        rstcorrModPpto("codigo_mod_ppto") = rstcorrModPpto("codigo_mod_ppto") + 1
        rstcorrModPpto.Update
        correlativo1 = rstcorrModPpto("codigo_mod_ppto")
        'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
      End If
      If rstcorrModPpto.State = 1 Then rstcorrModPpto.Close

      rstdestino.Open "select * from fo_cmbte_mod_ppto where codigo_mod_ppto = 0", db, adOpenDynamic, adLockOptimistic
      rstdestino.AddNew
      rstdestino("codigo_mod_ppto") = correlativo1
      rstdestino("Ges_Gestion") = Trim(LblGes_Gestion.Caption)
    End If
    If v_añadir = 2 Then
      rstdestino.Open "select * from fo_cmbte_mod_ppto where codigo_mod_ppto = " & Adofo_cmbte_mod_ppto.Recordset("codigo_mod_ppto"), db, adOpenDynamic, adLockOptimistic
    End If
    codigo_mod_ppto1 = rstdestino("codigo_mod_ppto")
    rstdestino("ges_gestion") = LblGes_Gestion
'    rstdestino("codigo_mod_ppto") = Lblcodigo_mod_ppto
    If OptTipo_resolucion1.Value = True Then
      rstdestino!tipo_resolucion = "INS"
    End If
    If OptTipo_resolucion2.Value = True Then
      rstdestino!tipo_resolucion = "MIN"
    End If
    rstdestino("tipo_modificacion") = Left(Trim(CmbTipo_modificacion.Text), 1)
    rstdestino("Nro_resolucion") = TxtNro_resolucion
    rstdestino("fecha_mod") = Format(Date, "dd/mm/yyyy")
    rstdestino("estado_aprobacion") = "N"
    
    rstdestino("uni_codigo_ori") = Txtuni_codigo_ori.Text
    rstdestino("pro_programa_ori") = TxtPro_programa_ori.Text
'    rstdestino("pro_subprograma_ori") = Txtpro_Subprograma_ori.Text
    rstdestino("pro_proyecto_ori") = Txtpro_proyecto_ori.Text
    rstdestino("pro_actividad_ori") = Txtpro_actividad_ori.Text
    rstdestino("fte_codigo_ori") = DtCFte_codigo_ori.Text
    rstdestino("org_codigo_ori") = DtCOrg_codigo_ori.Text
    rstdestino("par_codigo_ori") = DtCpar_codigo_ori.Text
    rstdestino("fgs_formulado_ori") = CDbl(Txtfgs_formulado_ori)
    rstdestino!fgs_adicion_ori = CDbl(Txtfgs_adicion_ori)
    rstdestino("fgs_modificaciones_ori") = CDbl(Txtfgs_modificaciones_ori)
    rstdestino("fgs_vigente_ori") = CDbl(Txtfgs_vigente_ori)
    'aqui rstdestino ("ent_codigo_ori")
    If V_accion = "TRANSFERENCIA" Then
      rstdestino("uni_codigo_des") = Txtuni_codigo_des.Text
      rstdestino("pro_programa_des") = TxtPro_programa_des.Text
'      rstdestino("pro_subprograma_des") = Txtpro_Subprograma_des.Text
      rstdestino("pro_proyecto_des") = Txtpro_proyecto_des.Text
      rstdestino("pro_actividad_des") = Txtpro_actividad_des.Text
      rstdestino("fte_codigo_des") = DtCFte_codigo_des.Text
      rstdestino("org_codigo_des") = DtCOrg_codigo_des.Text
      rstdestino("par_codigo_des") = DtCpar_codigo_des.Text
      rstdestino("fgs_formulado_des") = Txtfgs_formulado_des
      rstdestino("fgs_modificaciones_des") = CDbl(Txtfgs_modificaciones_des)
      rstdestino("fgs_vigente_des") = CDbl(Txtfgs_vigente_des)
      'aqui rstdestino ("ent_codigo_des")
    End If
    
    rstdestino("fecha_registro") = Format(Date, "dd/mm/yyyy")
    rstdestino("hora_registro") = Format(Time, "hh:mm:ss")
    rstdestino("usr_usuario") = GlUsuario
    rstdestino.Update
    db.CommitTrans
  End If
  
  rstfo_cmbte_mod_ppto.Requery
  rstfo_cmbte_mod_ppto.Find "codigo_mod_ppto = " & codigo_mod_ppto1, , adSearchForward, 1
  If rstfo_cmbte_mod_ppto.EOF Then rstfo_cmbte_mod_ppto.MoveLast
End Sub

Private Sub BtnImprimir_Click()
If rstfo_cmbte_mod_ppto.RecordCount > 0 Then
'===== Ini comando para iniciar impresión
  Dim iResult As Integer
  '  Cry.Reset
  Cry.WindowShowRefreshBtn = True
  If Adofo_cmbte_mod_ppto.Recordset!tipo_modificacion = "A" Then
    Cry.Formulas(1) = "add = '" & IIf(Adofo_cmbte_mod_ppto.Recordset!estado_aprobacion = "S", "A", "S") & "' "
    Cry.Formulas(2) = "mod = '" & " " & "' "
  End If
  If Adofo_cmbte_mod_ppto.Recordset!tipo_modificacion = "T" Then
    Cry.Formulas(1) = "add = '" & " " & "' "
    Cry.Formulas(2) = "mod = '" & IIf(Adofo_cmbte_mod_ppto.Recordset!estado_aprobacion = "S", "A", "S") & "' "
  End If
  Cry.ReportFileName = App.Path & "\FormsPresupuesto\ModificacionPresupuestaria\ComproModPpto.rpt"  ' App.Path & "\ModificacionPresupuestaria\ComproModPpto.rpt"
  
  db.BeginTrans
  Set rstfo_cmbte_mod_ppto_rep = New ADODB.Recordset
  rstfo_cmbte_mod_ppto_rep.Open "select * from fo_cmbte_mod_ppto_rep where maquina = '" & GlMaquina & "'", db, adOpenKeyset, adLockOptimistic
  While Not rstfo_cmbte_mod_ppto_rep.EOF
    rstfo_cmbte_mod_ppto_rep.Delete
    rstfo_cmbte_mod_ppto_rep.Update
    rstfo_cmbte_mod_ppto_rep.MoveNext
  Wend
  rstfo_cmbte_mod_ppto_rep.AddNew
  rstfo_cmbte_mod_ppto_rep("ges_gestion") = Trim(LblGes_Gestion.Caption)
  rstfo_cmbte_mod_ppto_rep("codigo_mod_ppto") = CInt(Lblcodigo_mod_ppto)
  rstfo_cmbte_mod_ppto_rep("tipo_modificacion") = Left(Trim(CmbTipo_modificacion.Text), 1)
  rstfo_cmbte_mod_ppto_rep("Nro_resolucion") = TxtNro_resolucion
  rstfo_cmbte_mod_ppto_rep("fecha_mod") = CDate(Adofo_cmbte_mod_ppto.Recordset("fecha_mod"))
  rstfo_cmbte_mod_ppto_rep("estado_aprobacion") = Adofo_cmbte_mod_ppto.Recordset("estado_aprobacion")
  
  rstfo_cmbte_mod_ppto_rep("uni_codigo_ori") = Txtuni_codigo_ori.Text
  rstfo_cmbte_mod_ppto_rep("uni_descripcion_ori") = Txtuni_codigo_ori.Text
  
  rstfo_cmbte_mod_ppto_rep("pro_programa_ori") = TxtPro_programa_ori.Text
'  rstfo_cmbte_mod_ppto_rep("pro_subprograma_ori") = Txtpro_Subprograma_ori.Text
  rstfo_cmbte_mod_ppto_rep("pro_proyecto_ori") = Txtpro_proyecto_ori.Text
  rstfo_cmbte_mod_ppto_rep("pro_actividad_ori") = Txtpro_actividad_ori.Text
  
  rstfo_cmbte_mod_ppto_rep("fte_codigo_ori") = DtCFte_codigo_ori.Text
  rstfo_cmbte_mod_ppto_rep("Fte_descripcion_larga_ori") = DtCFte_descripcion_larga_ori.Text
  
  rstfo_cmbte_mod_ppto_rep("org_codigo_ori") = DtCOrg_codigo_ori.Text
  rstfo_cmbte_mod_ppto_rep("Org_descripcion_ori") = DtCOrg_descripcion_ori
  
  rstfo_cmbte_mod_ppto_rep("par_codigo_ori") = DtCpar_codigo_ori.Text
  rstfo_cmbte_mod_ppto_rep("Par_descripcion_larga_ori") = Trim(DtCPar_descripcion_larga_ori.Text)
  
  rstfo_cmbte_mod_ppto_rep("fgs_formulado_ori") = CDbl(Txtfgs_formulado_ori)
  rstfo_cmbte_mod_ppto_rep("fgs_adicion_ori") = CDbl(Txtfgs_adicion_ori) 'fgs_adicion_ori
  rstfo_cmbte_mod_ppto_rep("fgs_modificaciones_ori") = CDbl(Txtfgs_modificaciones_ori)
  rstfo_cmbte_mod_ppto_rep("fgs_vigente_ori") = CDbl(Txtfgs_vigente_ori)
  'aqui rstfo_cmbte_mod_ppto_rep("ent_codigo_ori")
  
  If Left(Trim(CmbTipo_modificacion.Text), 1) <> "A" Then
    rstfo_cmbte_mod_ppto_rep("uni_codigo_des") = Txtuni_codigo_des.Text
    rstfo_cmbte_mod_ppto_rep("uni_descripcion_des") = Txtuni_codigo_des.Text
    
    rstfo_cmbte_mod_ppto_rep("pro_programa_des") = TxtPro_programa_des.Text
'    rstfo_cmbte_mod_ppto_rep("pro_subprograma_des") = Txtpro_Subprograma_des.Text
    rstfo_cmbte_mod_ppto_rep("pro_proyecto_des") = Txtpro_proyecto_des.Text
    rstfo_cmbte_mod_ppto_rep("pro_actividad_des") = Txtpro_actividad_des.Text
    
    rstfo_cmbte_mod_ppto_rep("fte_codigo_des") = DtCFte_codigo_des.Text
    rstfo_cmbte_mod_ppto_rep("Fte_descripcion_larga_des") = DtCFte_descripcion_larga_des.Text
    
    rstfo_cmbte_mod_ppto_rep("org_codigo_des") = DtCOrg_codigo_des.Text
    rstfo_cmbte_mod_ppto_rep("Org_descripcion_des") = DtCOrg_descripcion_des.Text
    
    rstfo_cmbte_mod_ppto_rep("par_codigo_des") = DtCpar_codigo_des.Text
    rstfo_cmbte_mod_ppto_rep("Par_descripcion_larga_des") = DtCPar_descripcion_larga_des.Text
    
    rstfo_cmbte_mod_ppto_rep("fgs_formulado_des") = Txtfgs_formulado_des
    rstfo_cmbte_mod_ppto_rep("fgs_modificaciones_des") = CDbl(Txtfgs_modificaciones_des)
    rstfo_cmbte_mod_ppto_rep("fgs_vigente_des") = CDbl(Txtfgs_vigente_des)
    'aqui rstfo_cmbte_mod_ppto_rep("ent_codigo_des")
  End If
  rstfo_cmbte_mod_ppto_rep("maquina") = GlMaquina
  rstfo_cmbte_mod_ppto_rep.Update
  db.CommitTrans
  
  Cry.SelectionFormula = "{Fo_cmbte_mod_ppto_rep.Maquina} = '" & GlMaquina & "'"
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

Private Sub BtnModificar_Click()
    v_añadir = 2
    If Adofo_cmbte_mod_ppto.Recordset("tipo_modificacion") = "A" Then
      fraOpciones.Visible = False
      FraGrabarCancelar.Visible = True
'      FraModpptoNav.Enabled = False
      FraModpptoDat.Enabled = True
      DtCFte_codigo_ori.Enabled = False
      DtCOrg_codigo_ori.Enabled = False
      
'      Txtfgs_formulado_ori.Enabled = True
      Txtfgs_modificaciones_ori.Enabled = False
      Txtfgs_vigente_ori.Enabled = False
      DtCFte_codigo_ori.Enabled = True
      FraORI.Enabled = True
  
      swmodificar = 1
      If swcopiar = 1 Then
        marca1 = Adofo_cmbte_mod_ppto.Recordset.Bookmark
      Else
        marca1 = Adofo_cmbte_mod_ppto.Recordset.Bookmark
      End If
      correlativo1 = Adofo_cmbte_mod_ppto.Recordset("codigo_mod_ppto")
      ges_gestion1 = Adofo_cmbte_mod_ppto.Recordset("ges_gestion")
    End If
    
    If Adofo_cmbte_mod_ppto.Recordset("tipo_modificacion") = "T" Then
      v_añadir = 2
      CmdTransfer_Click
    End If
    
    V_accion = "NORMAL"

End Sub

Private Sub BtnSalir_Click()
  sino = MsgBox("¿Está seguro de Salir?", vbQuestion + vbYesNo, "Confirmando...")
  If sino = vbYes Then
    Call cerrar
    Unload Me
  End If

End Sub

Private Sub cerrar()

End Sub

Private Sub add_correl()
'  Dim rstcorrModPpto As New ADODB.Recordset
'  Set rstcorrModPpto = New ADODB.Recordset
'  If rstcorrModPpto.State = 1 Then rstcorrModPpto.Close
'  rstcorrModPpto.Open "select * from fc_cmbte_mod_ppto_correl", db, adOpenDynamic, adLockOptimistic ' where org_codigo
'  If (rstcorrModPpto.EOF) Then
'  'rstcorrModPpto.Find "org_codigo = '" & (DtCOrg_codigo.Text) & "' ", , adSearchForward
'  'If rstcorrModPpto.EOF Then
'    rstcorrModPpto.AddNew
'    'rstcorrModPpto("org_codigo") = Trim(DtCOrg_codigo.Text)
'    'rstcorrModPpto("ges_gestion") = Trim(LblGes_Gestion.Caption)
'    rstcorrModPpto("codigo_mod_ppto") = 1
'    rstcorrModPpto.Update
'    correlativo1 = rstcorrModPpto("codigo_mod_ppto")
'    FrmModPresup.Lblcodigo_mod_ppto.Caption = rstcorrModPpto("codigo_mod_ppto")
'  Else
'    rstcorrModPpto.MoveFirst
'    rstcorrModPpto("codigo_mod_ppto") = rstcorrModPpto("codigo_mod_ppto") + 1
'    rstcorrModPpto.Update
'    correlativo1 = rstcorrModPpto("codigo_mod_ppto")
'    'FrmIngresosabm.LblCorrelativo_ingreso.Caption = rstcorrel_ing("correlativo")
'  End If
'  If rstcorrModPpto.State = 1 Then rstcorrModPpto.Close
End Sub

Private Sub DataCombo4_Click(Area As Integer)

End Sub


Private Sub CmdTransDes_Click()
  FraDatTrans.Visible = True
  FraDES.Visible = True
'  swtransfer = 2
  Lbluni_codigoD = Txtuni_codigo.Text
  LblFte_codigoD = DtCFte_codigo.Text
  LblFte_descripcion_largaD = DtCFte_descripcion_larga.Text
  LblOrg_codigoD = DtCorg_codigo.Text
  LblOrg_descripcionD = DtCOrg_descripcion.Text
  Lblpar_codigoD = DtCpar_codigo.Text
  LblPar_descripcion_largaD = DtCPar_descripcion_larga.Text
  LblPro_programaD = TxtPro_programa.Text
'  Lblpro_SubprogramaD = Txtpro_Subprograma.Text
  Lblpro_proyectoD = dtc_codigo1.Text
  Lblpro_actividadD = Txtpro_actividad.Text
  Lblfgs_formuladoD = Txtfgs_formulado
  Lblfgs_vigenteD = Txtfgs_vigente.Text
  If (Len(Trim(LblFte_codigoO)) > 0) And (Len(Trim(LblFte_codigoD)) > 0) Then
    CmdTransOk.Enabled = True
  Else
    CmdTransOk.Enabled = False
  End If
End Sub

Private Sub CmdTransfer_Click()

  '===== carga datos de fo_formulacion_gasto
  If rstFte_financia.State = 1 Then rstFte_financia.Close
  rstFte_financia.Open "Select * from Fc_fuente_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoFte_financia.Recordset = rstFte_financia
  AdoFte_financia.Refresh
  If Not AdoFte_financia.Recordset.BOF Then AdoFte_financia.Recordset.MoveFirst
  
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  rstOrganismo_finan.Open "Select * from Fc_organismo_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoOrganismo_finan.Recordset = rstOrganismo_finan
  AdoOrganismo_finan.Refresh
  If Not rstOrganismo_finan.BOF Then rstOrganismo_finan.MoveFirst
  
  Set rstfo_formulacion_gasto = New ADODB.Recordset
  queryinicial = "select * from fo_ppto_formulacion_gasto "
  If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
  rstfo_formulacion_gasto.Open queryinicial, db, adOpenKeyset, adLockReadOnly
'  rstIngresos.Sort = rstIngresos("correlativo_ingreso") & " " & "org_codigo"  '"correlativo_ingreso" & " " & "org_codigo"
  Set Adofo_formulacion_gasto.Recordset = rstfo_formulacion_gasto
  
  If v_añadir = 2 Then
    TxtNro_resolucionT.Text = TxtNro_resolucion.Text
    Txtfgs_modificacionesT = IIf(CDbl(Txtfgs_modificaciones_ori) < 0, CDbl(Txtfgs_modificaciones_ori) * -1, CDbl(Txtfgs_modificaciones_ori))
    '===== origen
    Lbluni_codigoO = Txtuni_codigo_ori.Text
    LblFte_codigoO = DtCFte_codigo_ori.Text
    LblFte_descripcion_largaO = DtCFte_descripcion_larga_ori.Text
    LblOrg_codigoO = DtCOrg_codigo_ori.Text
    LblOrg_descripcionO = DtCOrg_descripcion_ori.Text
    Lblpar_codigoO = DtCpar_codigo_ori.Text
    LblPar_descripcion_largaO = DtCPar_descripcion_larga_ori.Text
    LblPro_programaO = TxtPro_programa_ori.Text
'    Lblpro_SubprogramaO = Txtpro_Subprograma_ori
    Lblpro_proyectoO = Txtpro_proyecto_ori.Text
    Lblpro_actividadO = Txtpro_actividad_ori.Text
    
    Lblfgs_formuladoO = CDbl(Txtfgs_formulado_ori)
    'Txtfgs_modificacionesT = Txtfgs_modificaciones_ori
    Lblfgs_vigenteO = CDbl(Txtfgs_vigente_ori) + CDbl(Txtfgs_modificaciones_ori)
    'Txtfgs_vigente_ori = CDbl(Lblfgs_formuladoO) - CDbl(Txtfgs_modificacionesT)
    
    '===== destino
    Lbluni_codigoD = Txtuni_codigo_des.Text
    LblFte_codigoD = DtCFte_codigo_des
    LblFte_descripcion_largaD = DtCFte_descripcion_larga_des.Text
    LblOrg_codigoD = DtCOrg_codigo_des.Text
    LblOrg_descripcionD = DtCOrg_descripcion_des.Text
    Lblpar_codigoD = DtCpar_codigo_des.Text
    LblPar_descripcion_largaD = DtCPar_descripcion_larga_des.Text
    LblPro_programaD = TxtPro_programa_des.Text
'    Lblpro_SubprogramaD = Txtpro_Subprograma_des.Text
    Lblpro_proyectoD = Txtpro_proyecto_des.Text
    Lblpro_actividadD = Txtpro_actividad_des.Text
    Lblfgs_formuladoD = Txtfgs_formulado_des
    'Txtfgs_modificaciones_des = CDbl(Txtfgs_modificacionesT)
    Lblfgs_vigenteD = CDbl(Txtfgs_vigente_des) - CDbl(Txtfgs_modificaciones_des) 'Txtfgs_vigente_des = Lblfgs_formuladoD) + CDbl(Txtfgs_modificacionesT)
  Else
    Lbluni_codigoO = ""
    LblFte_codigoO = ""
    LblFte_descripcion_largaO = ""
    LblOrg_codigoO = ""
    LblOrg_descripcionO = ""
    Lblpar_codigoO = ""
    LblPar_descripcion_largaO = ""
    LblPro_programaO = ""
'    Lblpro_SubprogramaO = ""
    Lblpro_proyectoO = ""
    Lblpro_actividadO = ""
    Lblfgs_formuladoO = 0
    Lblfgs_vigenteO = 0
  
    Lbluni_codigoD = ""
    LblFte_codigoD = ""
    LblFte_descripcion_largaD = ""
    LblOrg_codigoD = ""
    LblOrg_descripcionD = ""
    Lblpar_codigoD = ""
    LblPar_descripcion_largaD = ""
    LblPro_programaD = ""
'    Lblpro_SubprogramaD = ""
    Lblpro_proyectoD = ""
    Lblpro_actividadD = ""
    Lblfgs_formuladoD = 0
    Lblfgs_vigenteD = 0
  
    Txtfgs_modificacionesT = 0
    TxtNro_resolucionT = ""
    v_añadir = 1
  End If
  V_accion = "TRANSFERECIA"
  fraOpciones.Visible = False
  fraOpciones.Enabled = False
  FraCmdTrans.Visible = True
  FraCmdTrans.Enabled = True
  FraNavega.Visible = True
  FraNavega.Enabled = True
  
'  FraModpptoNav.Visible = False
  FraNavega.Visible = True
  
  FraDatTrans.Visible = True
  FraDatTrans.Enabled = True
  If (Len(Trim(LblFte_codigoO)) > 0) And (Len(Trim(LblFte_codigoD)) > 0) Then
    CmdTransOk.Enabled = True
  Else
    CmdTransOk.Enabled = False
  End If

End Sub

Private Sub CmdTransNoTot_Click()
  FraNavega.Visible = False
  FraNavega.Enabled = False
  FraDatTrans.Visible = False
  FraDatTrans.Enabled = False
  FraCmdTrans.Visible = False
  FraCmdTrans.Enabled = True
  fraOpciones.Visible = True
  fraOpciones.Enabled = True
'  FraModpptoNav.Visible = True
  FraNavega.Visible = False
  If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.CancelUpdate
  If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
  If rstTipo_moneda.State = 1 Then rstTipo_moneda.Close
  If rstFte_financia.State = 1 Then rstFte_financia.Close
  If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
  v_añadir = 0
End Sub


Private Sub CmdTransOk_Click()
  swigual = 1
  If LblFte_codigoO <> LblFte_codigoD Then swigual = 0
  If LblOrg_codigoO <> LblOrg_codigoD Then swigual = 0
  If Lblpar_codigoO <> Lblpar_codigoD Then swigual = 0
  If LblPar_descripcion_largaO <> LblPar_descripcion_largaD Then swigual = 0
  If LblPro_programaO <> LblPro_programaD Then swigual = 0
'  If Lblpro_SubprogramaO <> Lblpro_SubprogramaD Then swigual = 0
  If Lblpro_proyectoO <> Lblpro_proyectoD Then swigual = 0
  If Lblpro_actividadO <> Lblpro_actividadD Then swigual = 0
  
  If swigual = 0 Then
    Call valida_trans
  Else
    MsgBox "El origen no puede ser el mismo que el destino", vbOKOnly + vbExclamation, "Error ..."
    swigual = 0
  End If
  
'  v_añadir = 0

End Sub

Private Sub CmdTransOri_Click()
  FraDatTrans.Visible = True
'  swtransfer = 1
  If Txtfgs_vigente > 0 Then
    Lbluni_codigoO = Txtuni_codigo.Text
    LblFte_codigoO = DtCFte_codigo.Text
    LblFte_descripcion_largaO = DtCFte_descripcion_larga.Text
    LblOrg_codigoO = DtCorg_codigo.Text
    LblOrg_descripcionO = DtCOrg_descripcion.Text
    Lblpar_codigoO = DtCpar_codigo.Text
    LblPar_descripcion_largaO = DtCPar_descripcion_larga.Text
    LblPro_programaO = TxtPro_programa.Text
'    Lblpro_SubprogramaO = Txtpro_Subprograma.Text
    Lblpro_proyectoO = dtc_codigo1.Text
    Lblpro_actividadO = Txtpro_actividad.Text
    Lblfgs_formuladoO = Txtfgs_formulado
    Lblfgs_vigenteO = Txtfgs_vigente.Text
    If (Len(Trim(LblFte_codigoO)) > 0) And (Len(Trim(LblFte_codigoD)) > 0) Then
      CmdTransOk.Enabled = True
    Else
      CmdTransOk.Enabled = False
    End If
  Else
    MsgBox "La estructura no tiene monto vigente ...", vbOKOnly + vbInformation, "Error ..."
  End If
End Sub

Private Sub DtCFte_codigo_Click(Area As Integer)
    DtCFte_descripcion_larga.BoundText = DtCFte_codigo.BoundText
End Sub

Private Sub DtCFte_codigo_ori_Click(Area As Integer)
   DtCFte_descripcion_larga_ori.Text = DtCFte_codigo_ori.BoundText
'    DtCFte_descripcion_larga.Text = DtCFte_codigo.BoundText
    DtCOrg_codigo_ori.Enabled = True
    Call pfil_Org_Fte(DtCFte_codigo_ori.Text)
End Sub

Private Sub DtCFte_descripcion_larga_Click(Area As Integer)
    DtCFte_codigo.BoundText = DtCFte_descripcion_larga.BoundText
End Sub

Private Sub DtCFte_descripcion_larga_ori_Click(Area As Integer)
   DtCFte_codigo_ori.Text = DtCFte_descripcion_larga_ori.BoundText
End Sub

Private Sub DtCOrg_codigo_Click(Area As Integer)
    DtCOrg_descripcion.BoundText = DtCorg_codigo.BoundText
End Sub

Private Sub DtCOrg_codigo_ori_Click(Area As Integer)
  DtCOrg_descripcion_ori.Text = DtCOrg_codigo_ori.BoundText
End Sub

Private Sub DtCOrg_descripcion_Click(Area As Integer)
    DtCorg_codigo.BoundText = DtCOrg_descripcion.BoundText
End Sub

Private Sub DtCOrg_descripcion_ori_Click(Area As Integer)
  DtCOrg_codigo_ori.Text = DtCOrg_descripcion_ori.BoundText
End Sub

Private Sub DtCpar_codigo_Click(Area As Integer)
    DtCPar_descripcion_larga.BoundText = DtCpar_codigo.BoundText
End Sub

Private Sub DtCpar_codigo_ori_Click(Area As Integer)
  DtCPar_descripcion_larga_ori.Text = DtCpar_codigo_ori.BoundText
End Sub

Private Sub DtCPar_descripcion_larga_Click(Area As Integer)
    DtCpar_codigo.BoundText = DtCPar_descripcion_larga.BoundText
End Sub

Private Sub DtCPar_descripcion_larga_ori_Click(Area As Integer)
  DtCpar_codigo_ori.Text = DtCPar_descripcion_larga_ori.BoundText
End Sub

Private Sub Form_Load()
  '===== Ini cargado de tablas de consulta y de datos de despliegue
'  LblUsuario.Caption = LblUsuario.Caption + GlUsuario
  swgraba = 0
  marca1 = 0
  swcopiar = 0
  V_accion = "TRANSFERENCIA"
  
    'fc_estructura_programatica
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from fc_estructura_programatica order by pro_descripcion", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
'  Set rstTipo_moneda = New ADODB.Recordset
'  If rstTipo_moneda.State = 1 Then rstTipo_moneda.Close
'  rstTipo_moneda.Open "select * from Tipo_moneda order by denominacion_moneda", db, adOpenKeyset, adLockReadOnly
'  Set AdoTipo_moneda.Recordset = rstTipo_moneda
'  AdoTipo_moneda.Refresh
'  If Not AdoTipo_moneda.Recordset.BOF Then AdoTipo_moneda.Recordset.MoveFirst
  
'  Set rstTipo_comprobante = New ADODB.Recordset
'  If rstTipo_comprobante.State = 1 Then rstTipo_comprobante.Close
'  rstTipo_comprobante.Open "select * from Tipo_comprobante where ingresos = 'A' order by denominacion_tipo", db, adOpenKeyset, adLockReadOnly
'  Set AdoTipo_comprobante.Recordset = rstTipo_comprobante
'  AdoTipo_comprobante.Refresh
'  If Not AdoTipo_comprobante.Recordset.BOF Then AdoTipo_comprobante.Recordset.MoveFirst
  
'  If rstfc_unidad_ejecutora_ori.State = 1 Then rstfc_unidad_ejecutora_ori.Close
'  rstfc_unidad_ejecutora_ori.Open "select * from fc_unidad_ejecutora", db, adOpenKeyset, adLockReadOnly
'  Set Adofc_unidad_ejecutora_ori.Recordset = rstfc_unidad_ejecutora_ori
'  Adofc_unidad_ejecutora_ori.Refresh
'  If Not Adofc_unidad_ejecutora_ori.Recordset.BOF Then Adofc_unidad_ejecutora_ori.Recordset.MoveFirst
'
'  If rstfc_unidad_ejecutora_des.State = 1 Then rstfc_unidad_ejecutora_des.Close
'  rstfc_unidad_ejecutora_des.Open "select * from fc_unidad_ejecutora", db, adOpenKeyset, adLockReadOnly
'  Set Adofc_unidad_ejecutora_des.Recordset = rstfc_unidad_ejecutora_des
'  Adofc_unidad_ejecutora_des.Refresh
'  If Not Adofc_unidad_ejecutora_des.Recordset.BOF Then Adofc_unidad_ejecutora_des.Recordset.MoveFirst
  
  If rstFte_financia_ori.State = 1 Then rstFte_financia_ori.Close
  rstFte_financia_ori.Open "Select * from fc_fuente_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoFte_financia_ori.Recordset = rstFte_financia_ori
  AdoFte_financia_ori.Refresh
  If Not AdoFte_financia_ori.Recordset.BOF Then AdoFte_financia_ori.Recordset.MoveFirst
  
  If rstFte_financia_des.State = 1 Then rstFte_financia_des.Close
  rstFte_financia_des.Open "Select * from fc_fuente_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoFte_financia_des.Recordset = rstFte_financia_des
  AdoFte_financia_des.Refresh
  If Not AdoFte_financia_des.Recordset.BOF Then AdoFte_financia_des.Recordset.MoveFirst
  
  If rstOrganismo_finan_ori.State = 1 Then rstOrganismo_finan_ori.Close
  rstOrganismo_finan_ori.Open "Select * from Fc_organismo_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoOrganismo_finan_ori.Recordset = rstOrganismo_finan_ori
  AdoOrganismo_finan_ori.Refresh
  If Not AdoOrganismo_finan_ori.Recordset.BOF Then AdoOrganismo_finan_ori.Recordset.MoveFirst
  
  If rstOrganismo_finan_des.State = 1 Then rstOrganismo_finan_des.Close
  rstOrganismo_finan_des.Open "Select * from Fc_organismo_financiamiento", db, adOpenKeyset, adLockReadOnly
  Set AdoOrganismo_finan_des.Recordset = rstOrganismo_finan_des
  AdoOrganismo_finan_des.Refresh
  If Not AdoOrganismo_finan_des.Recordset.BOF Then AdoOrganismo_finan_des.Recordset.MoveFirst
  
  If rstfc_partida_gasto_ori.State = 1 Then rstfc_partida_gasto_ori.Close
  rstfc_partida_gasto_ori.Open "Select * from fc_partida_gasto", db, adOpenKeyset, adLockReadOnly
  Set Adofc_partida_gasto_ori.Recordset = rstfc_partida_gasto_ori
  Adofc_partida_gasto_ori.Refresh
  If Not Adofc_partida_gasto_ori.Recordset.BOF Then Adofc_partida_gasto_ori.Recordset.MoveFirst
  
  If rstfc_partida_gasto_des.State = 1 Then rstfc_partida_gasto_des.Close
  rstfc_partida_gasto_des.Open "Select * from fc_partida_gasto", db, adOpenKeyset, adLockReadOnly
  Set Adofc_partida_gasto_des.Recordset = rstfc_partida_gasto_des
  Adofc_partida_gasto_des.Refresh
  If Not Adofc_partida_gasto_des.Recordset.BOF Then Adofc_partida_gasto_des.Recordset.MoveFirst
  
  'Adofc_partida_gasto_ori
  
'  If rstac_documento_respaldo.State = 1 Then rstac_documento_respaldo.Close
'  Set rstac_documento_respaldo = New ADODB.Recordset
'  rstac_documento_respaldo.Open "select * from ac_documento_respaldo", db, adOpenKeyset, adLockReadOnly
'  Set Adoac_documento_respaldo.Recordset = rstac_documento_respaldo
'  Adoac_documento_respaldo.Refresh
'  If Not Adoac_documento_respaldo.Recordset.BOF Then Adoac_documento_respaldo.Recordset.MoveFirst
  
  Set rstfo_formulacion_gasto = New ADODB.Recordset
  queryinicial = "select * from fo_ppto_formulacion_gasto "
  If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
  rstfo_formulacion_gasto.Open queryinicial, db, adOpenKeyset, adLockReadOnly
'  rstIngresos.Sort = rstIngresos("correlativo_ingreso") & " " & "org_codigo"  '"correlativo_ingreso" & " " & "org_codigo"
  Set Adofo_formulacion_gasto.Recordset = rstfo_formulacion_gasto
  
'  Set rstfo_cmbte_mod_ppto = New ADODB.Recordset
'  ' pa busqueda QueryInicial = "select * from fo_ingresos where estado_aprobacion <> 'S'" 'ORDER BY correlativo_ingreso , org_codigo
'  queryinicial = "select * from fo_cmbte_mod_ppto where estado_aprobacion <> 'S' and estado_aprobacion <> 'E'" ' ORDER BY codigo_mod_ppto"
'  If rstfo_cmbte_mod_ppto.State = 1 Then rstfo_cmbte_mod_ppto.Close
'  rstfo_cmbte_mod_ppto.Open queryinicial & " ORDER BY codigo_mod_ppto", db, adOpenDynamic, adLockOptimistic
'  Set Adofo_cmbte_mod_ppto.Recordset = rstfo_cmbte_mod_ppto
'
'  If (Not Adofo_cmbte_mod_ppto.Recordset.BOF) And (Not Adofo_cmbte_mod_ppto.Recordset.EOF) Then
'
'  End If
  '===== fin cargado de tablas de consulta y de datos de despliegue

	Call SeguridadSet(Me)
End Sub
Private Sub valida()
  swgraba = 1
  If Len(Trim(TxtNro_resolucion.Text)) < 1 Then swgraba = 0
  If Len(Trim(CmbTipo_modificacion.Text)) < 1 Then swgraba = 0
  If Len(Trim(Txtuni_codigo_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCFte_codigo_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCFte_descripcion_larga_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCOrg_codigo_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCOrg_descripcion_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCpar_codigo_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(DtCPar_descripcion_larga_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(TxtPro_programa_ori.Text)) < 1 Then swgraba = 0
'  If Len(Trim(Txtpro_Subprograma_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(Txtpro_proyecto_ori.Text)) < 1 Then swgraba = 0
  If Len(Trim(Txtpro_actividad_ori.Text)) < 1 Then swgraba = 0
'  If Len(Trim(Txtfgs_formulado_ori.Text)) < 1 Then swgraba = 0
'  If Len(Trim(Txtfgs_modificaciones_ori.Text)) < 1 = 0 Then swgraba = 0
'  If Len(Trim(Txtfgs_vigente_ori.Text)) < 1 Then swgraba = 0
  If swgraba = 0 Then
    MsgBox "Los datos están incompletos, Por favor revíselos, o cancele el proceso", vbInformation + vbOKOnly, "Error al grabar los datos"
  End If
End Sub

Private Sub LblFte_codigoD_Change()
  If Len(Trim(LblFte_codigoD)) > 0 Then
    Label25.BackColor = &HFFC0C0
  Else
    Label25.BackColor = &H808080
  End If
End Sub

Private Sub LblFte_codigoO_Change()
  If Len(Trim(LblFte_codigoO)) > 0 Then
    Label28.BackColor = &HFFC0C0
  Else
    Label28.BackColor = &H808080
  End If
End Sub

Private Sub OptFilGral1_Click()
  queryinicial = "select * from fo_cmbte_mod_ppto where estado_aprobacion <> 'S' and estado_aprobacion <> 'E'"
  If rstfo_cmbte_mod_ppto.State = 1 Then rstfo_cmbte_mod_ppto.CancelUpdate
  If rstfo_cmbte_mod_ppto.State = 1 Then rstfo_cmbte_mod_ppto.Close
  rstfo_cmbte_mod_ppto.Open queryinicial & " ORDER BY codigo_mod_ppto", db, adOpenDynamic, adLockOptimistic
  rstfo_cmbte_mod_ppto.Requery
  Set Adofo_cmbte_mod_ppto.Recordset = rstfo_cmbte_mod_ppto
End Sub

Private Sub OptFilGral2_Click()
  queryinicial = "select * from fo_cmbte_mod_ppto"
  If rstfo_cmbte_mod_ppto.State = 1 Then rstfo_cmbte_mod_ppto.CancelUpdate
  If rstfo_cmbte_mod_ppto.State = 1 Then rstfo_cmbte_mod_ppto.Close
  rstfo_cmbte_mod_ppto.Open queryinicial & " ORDER BY codigo_mod_ppto ", db, adOpenDynamic, adLockOptimistic
  rstfo_cmbte_mod_ppto.Requery
  Set Adofo_cmbte_mod_ppto.Recordset = rstfo_cmbte_mod_ppto

End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
End Sub

Private Sub Txtfgs_adicion_ori_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtfgs_adicion_ori_KeyUp(KeyCode As Integer, Shift As Integer)
  Txtfgs_vigente_ori = Txtfgs_formulado_ori + Txtfgs_adicion_ori
End Sub

Private Sub Txtfgs_formulado_ori_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
  Else
    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtfgs_formulado_ori_KeyUp(KeyCode As Integer, Shift As Integer)
  Txtfgs_vigente_ori = Txtfgs_formulado_ori
End Sub

Private Sub Txtfgs_modificacionesT_Change()
  If Txtfgs_modificacionesT > CDbl(Lblfgs_vigenteO) Then
    Txtfgs_modificacionesT_KeyPress (0)
  Else
    'MsgBox "mayor"
'    KeyAscii = Asc(UCase(Chr(0)))
  End If
End Sub

Private Sub Txtfgs_modificacionesT_KeyPress(KeyAscii As Integer)
'  If (KeyAscii < 58) And (KeyAscii > 47) Or (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 44) Then
'  Else
'    KeyAscii = Asc(UCase(Chr(0)))
'  End If


  If Txtfgs_modificacionesT > CDbl(Lblfgs_vigenteO) Then
'    KeyCode
    KeyAscii = Asc(UCase(Chr(8)))
    MsgBox "No se puede restar un monto mayor al vigente", vbInformation + vbOKOnly, "Error en el monto a modificar..."
    Txtfgs_modificacionesT = 0
  Else
    'MsgBox "mayor"
'    KeyAscii = Asc(UCase(Chr(0)))
  End If

End Sub

Private Sub Txtfgs_modificacionesT_KeyUp(KeyCode As Integer, Shift As Integer)
  If Txtfgs_modificacionesT > CDbl(Lblfgs_vigenteO) Then
    KeyCode = 18
    'MsgBox "No se puede restar un monto mayor al vigente", vbInformation + vbOKOnly,"Error en el monto a modificar..."
    'KeyAscii = Asc(UCase(Chr(0)))
  Else
    'MsgBox "mayor"
  End If
End Sub

Private Sub TxtNro_resolucion_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNro_resolucionT_Change()
  If Len(TxtNro_resolucionT) > 0 Then
    Txtfgs_modificacionesT.Enabled = True
  Else
    Txtfgs_modificacionesT.Enabled = False
  End If

End Sub

Private Sub valida_trans()
  swvalida_trans = 1
  If Len(Trim(TxtNro_resolucionT)) < 1 Then swvalida_trans = 0
  If Txtfgs_modificacionesT < 1 Then swvalida_trans = 0
  If Len(Trim(LblFte_codigoO)) < 1 Then swvalida_trans = 0
  If Len(Trim(LblFte_codigoD)) < 1 Then swvalida_trans = 0
  If valida_ppto(LblOrg_codigoO, Lblpar_codigoO, LblPro_programaO, Lblpro_proyectoO, Lblpro_actividadO) = 0 Then swvalida_trans = 0
  
  If swvalida_trans = 1 Then
'    v_añadir = 1
    FraDES.Visible = True
    TxtNro_resolucion.Text = TxtNro_resolucionT.Text
    CmbTipo_modificacion.Text = "TRANSFERENCIA"
    If OptIns.Value = True Then
      OptTipo_resolucion1.Value = True
    End If
    If OptMin.Value = True Then
      OptTipo_resolucion2.Value = True
    End If
    
    '===== origen
    Txtuni_codigo_ori.Text = Lbluni_codigoO
    DtCFte_codigo_ori.Text = LblFte_codigoO
    DtCFte_descripcion_larga_ori.Text = LblFte_descripcion_largaO
    DtCOrg_codigo_ori.Text = LblOrg_codigoO
    DtCOrg_descripcion_ori.Text = LblOrg_descripcionO
    DtCpar_codigo_ori.Text = Lblpar_codigoO
    DtCPar_descripcion_larga_ori.Text = LblPar_descripcion_largaO
    TxtPro_programa_ori.Text = LblPro_programaO
'    Txtpro_Subprograma_ori = Lblpro_SubprogramaO
    Txtpro_proyecto_ori.Text = Lblpro_proyectoO
    Txtpro_actividad_ori.Text = Lblpro_actividadO
    Txtfgs_formulado_ori = CDbl(Lblfgs_formuladoO)
    Txtfgs_adicion_ori = 0
    Txtfgs_modificaciones_ori = CDbl(Txtfgs_modificacionesT) * -1
    Txtfgs_vigente_ori = IIf(CDbl(Lblfgs_formuladoO) = 0, 0, CDbl(Lblfgs_formuladoO) - CDbl(Txtfgs_modificacionesT)) 'CDbl(Txtfgs_modificacionesT)
    '===== destino
    Txtuni_codigo_des.Text = Lbluni_codigoD
    DtCFte_codigo_des = LblFte_codigoD
    DtCFte_descripcion_larga_des.Text = LblFte_descripcion_largaD
    DtCOrg_codigo_des.Text = LblOrg_codigoD
    DtCOrg_descripcion_des.Text = LblOrg_descripcionD
    DtCpar_codigo_des.Text = Lblpar_codigoD
    DtCPar_descripcion_larga_des.Text = LblPar_descripcion_largaD
    TxtPro_programa_des.Text = LblPro_programaD
'    Txtpro_Subprograma_des.Text = Lblpro_SubprogramaD
    Txtpro_proyecto_des.Text = Lblpro_proyectoD
    Txtpro_actividad_des.Text = Lblpro_actividadD
    Txtfgs_formulado_des = CDbl(Lblfgs_formuladoD)
    Txtfgs_modificaciones_des = CDbl(Txtfgs_modificacionesT)
    Txtfgs_vigente_des = CDbl(Lblfgs_formuladoD) + CDbl(Txtfgs_modificacionesT)
    
    FraNavega.Visible = False
    FraNavega.Enabled = False
    FraDatTrans.Visible = False
    FraDatTrans.Enabled = False
    FraCmdTrans.Visible = False
    FraCmdTrans.Enabled = True
    fraOpciones.Visible = True
    fraOpciones.Enabled = True
        
'    FraModpptoNav.Visible = True
    FraNavega.Visible = False

    If rstfo_formulacion_gasto.State = 1 Then rstfo_formulacion_gasto.Close
    If rstTipo_moneda.State = 1 Then rstTipo_moneda.Close
    If rstFte_financia.State = 1 Then rstFte_financia.Close
    If rstOrganismo_finan.State = 1 Then rstOrganismo_finan.Close
    V_accion = "TRANSFERENCIA"
    Call BtnGrabar_Click
  Else
    MsgBox "Por Favor Complete los datos", vbExclamation + vbOKOnly, "ERROR al intentar grabar los la transferencia..."
  End If
  
End Sub

Private Sub errado()
'===== proceso para eliminar registros
  Dim rsterrado As New ADODB.Recordset
  If rsterrado.State = 1 Then rsterrado.Close
  Set rsterrado = New ADODB.Recordset
  rsterrado.Open "select * from fo_cmbte_mod_ppto where codigo_mod_ppto = " & Adofo_cmbte_mod_ppto.Recordset("codigo_mod_ppto"), db, adOpenKeyset, adLockOptimistic
    If rsterrado.RecordCount > 0 Then
        rsterrado("estado_aprobacion") = "E"
    End If
    rsterrado.Update
  If rsterrado.State = 1 Then rsterrado.Close
  rstfo_cmbte_mod_ppto.Update
  rstfo_cmbte_mod_ppto.Requery
  Set Adofo_cmbte_mod_ppto.Recordset = rstfo_cmbte_mod_ppto
  Adofo_cmbte_mod_ppto.Refresh
End Sub

Private Sub pfil_Org_Fte(Codfte As String)
'===== Proceso para filtrar los Organismos en base a la Fuente de financiamiento
  If rstOrganismo_finan_ori.State = 1 Then rstOrganismo_finan_ori.Close
  rstOrganismo_finan_ori.Open "Select * from Fc_organismo_financiamiento where fte_codigo = '" & Codfte & "'", db, adOpenDynamic, adLockReadOnly
  If rstOrganismo_finan_ori.RecordCount < 1 Then
    DtCOrg_codigo_ori.Text = ""
    DtCOrg_descripcion_ori.Text = ""
  End If
  Set AdoOrganismo_finan_ori.Recordset = rstOrganismo_finan_ori
  AdoOrganismo_finan_ori.Refresh
  If Not rstOrganismo_finan_ori.BOF Then rstOrganismo_finan_ori.MoveFirst
End Sub

Private Sub TxtNro_resolucionT_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Public Function valida_ppto(varOrg_codigoO, varpar_codigoO, varPro_programaO, varpro_proyectoO, varpro_actividadO)
  Dim rsFO_formulacion_gasto As New ADODB.Recordset
  Set rsFO_formulacion_gasto = New ADODB.Recordset
  If rsFO_formulacion_gasto.State = 1 Then rsFO_formulacion_gasto.Close
  rsFO_formulacion_gasto.Open "select * from fo_ppto_formulacion_gasto where Org_codigo = '" & varOrg_codigoO & "' and par_codigo = '" & varpar_codigoO & "' and Pro_programa = '" & varPro_programaO & "' and pro_proyecto = '" & varpro_proyectoO & "' and pro_actividad = '" & varpro_actividadO & "' ", db, adOpenKeyset, adLockReadOnly
  If rsFO_formulacion_gasto.RecordCount > 0 Then
'    fgs_vigente
'    fgs_compromiso
'    Txtfgs_modificacionesT
    If (rsFO_formulacion_gasto!FGS_VIGENTE - rsFO_formulacion_gasto!FGS_compromiso) >= Txtfgs_modificacionesT Then
      valida_ppto = 1
    Else
      valida_ppto = 0
      MsgBox "El saldo presupuestario ya está comprometido", vbCritical + vbOKOnly, "Error en búsqueda... "
    End If
  Else
    MsgBox " Error en Estructura Presupuestaria", vbCritical + vbOKOnly, "Error en búsqueda... "
  End If
  If rsFO_formulacion_gasto.State = 1 Then rsFO_formulacion_gasto.Close
End Function

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

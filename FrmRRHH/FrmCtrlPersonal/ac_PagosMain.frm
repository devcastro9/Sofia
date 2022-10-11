VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ac_PagosMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal - Orden de Pago"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11910
   Icon            =   "ac_PagosMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Beneficiarios del pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3975
      Left            =   0
      TabIndex        =   13
      Top             =   4320
      Width           =   7575
      Begin VB.CommandButton cmdConformidad 
         Caption         =   "Conformidad"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   36
         Top             =   -120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdEmiteFactura 
         Caption         =   "Emite Factura ?"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   26
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdCapMontos 
         Caption         =   "Captura Montos"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelACYD 
         Caption         =   "CYD"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdRelADevengado 
         Caption         =   "DEV"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdRelAComprometido 
         Caption         =   "COM"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdEliminarBeneficiario 
         Caption         =   "Eliminar beneficiario"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdicionarBeneficiario 
         Caption         =   "Adicionar beneficiario"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdJalaMismasPersonas 
         Caption         =   "Adicionar las mismas personas del pago anterior"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "ac_PagosMain.frx":0ECA
         Height          =   2775
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648384
         HeadLines       =   2
         RowHeight       =   22
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
         ColumnCount     =   18
         BeginProperty Column00 
            DataField       =   "idfuncionario"
            Caption         =   "Id Pers."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "paterno"
            Caption         =   "Primer Apellido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "materno"
            Caption         =   "Segundo Apellido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "nombres"
            Caption         =   "Nombres"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "ci"
            Caption         =   "Doc. Identidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "devengado"
            Caption         =   "Aprob"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "conformidad"
            Caption         =   "Confor- midad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "emitefactura"
            Caption         =   "Emite Factura?"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "monto_dol_ext"
            Caption         =   "Monto $US financ. externo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "monto_dol_nal"
            Caption         =   "Monto $US financ. nacional"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "ncite_conformidad"
            Caption         =   "Nro. cite conformidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "nchist"
            Caption         =   "File"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "item"
            Caption         =   "Item"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   ""
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   ""
            Caption         =   "Monto Bs financ. externo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   ""
            Caption         =   "Monto Bs financ. nacional"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "fcite_conformidad"
            Caption         =   "Fecha cite conformidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "fte_financiamientohist"
            Caption         =   "Fte Financiamiento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column08 
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "ac_PagosMain.frx":0EE9
         Height          =   1110
         Left            =   120
         TabIndex        =   14
         Top             =   2745
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1958
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
   Begin VB.Frame FraGrupo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planillas / Grupos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   7575
      Left            =   7680
      TabIndex        =   7
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton cmdImport 
         Caption         =   "Importar consultoria desde 2000"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdCopiaCrono 
         Caption         =   "Copia cronograma"
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dgGrupos 
         Bindings        =   "ac_PagosMain.frx":0F08
         Height          =   5655
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
         HeadLines       =   2
         RowHeight       =   23
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "ges_gestion"
            Caption         =   "gestion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "codigo_unidad"
            Caption         =   "Unidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "codigo_grupo"
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
         BeginProperty Column03 
            DataField       =   "descripcion_grupo"
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
         BeginProperty Column04 
            DataField       =   "espagogrupo"
            Caption         =   "Es Pago por Grupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            Locked          =   -1  'True
            BeginProperty Column00 
               Object.Visible         =   0   'False
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810.142
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkFiltroMain 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Filtrar por Unidad"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdicionarGrupo 
         Caption         =   "Adiciona Grupo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdBorraGrupo 
         Caption         =   "Elimina Grupo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdModificarGrupo 
         Caption         =   "Modifica Grupo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "ac_PagosMain.frx":0F20
         Height          =   2220
         Left            =   120
         TabIndex        =   8
         Top             =   4875
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3916
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.CommandButton cmdPrintPlanilla 
         Caption         =   "Imprime Grupo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprime Planilla / Grupo"
         Top             =   480
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dtcCodigo_unidad 
         Bindings        =   "ac_PagosMain.frx":0F38
         Height          =   315
         Left            =   1920
         TabIndex        =   38
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Uni_codigo"
         Text            =   "dtcCodigo_unidad"
         Object.DataMember      =   "dbo_edListaUnidadEjecutora"
      End
      Begin MSAdodcLib.Adodc adoGrupos 
         Height          =   330
         Left            =   120
         Top             =   7080
         Width           =   4020
         _ExtentX        =   7091
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
         Caption         =   "Grupos / Planillas"
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
      Begin VB.Label labCerrar 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   495
         Left            =   3360
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Orden de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      Begin VB.CommandButton cmdAdicionarPago 
         Caption         =   "Adicionar pago"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdEliminarPago 
         Caption         =   "Eliminar pago"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdModificarPago 
         Caption         =   "Modificar pago"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintPago 
         Caption         =   "Print detalle pago"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Envio a la unidad"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdAprobadoParaEnvio 
         Caption         =   "Aprobar Envío"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   21
         ToolTipText     =   "Envio a la unidad"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGenF02 
         Caption         =   "Aprueba Pago"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   39
         ToolTipText     =   "Aprueba Pago y Genera Cmpbtes."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdPrDetHonorarios 
         Caption         =   "Detalle Honorario"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdPrintOP 
         Caption         =   "Imprime Orden de Pago"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdRelCronoCompro 
         Caption         =   "Crono vs Compro"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdActConformidad 
         Caption         =   "Actualiza conformidad se trslado al f02"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   23
         ToolTipText     =   "Actualiza conformidades recibidas de la unidad"
         Top             =   -120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDown 
         Height          =   435
         Left            =   6120
         Picture         =   "ac_PagosMain.frx":0F49
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Permite mover el registro del postulante una posición más abajo"
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdUp 
         Height          =   435
         Left            =   5520
         Picture         =   "ac_PagosMain.frx":138B
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Permite mover el registro del postulante una posición más arriba"
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdGenDEVHist 
         Caption         =   "Genera DEV historico"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGenDevengado 
         Caption         =   "Generar Devengado se trslado al f02"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   -120
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dgPagos 
         Bindings        =   "ac_PagosMain.frx":17CD
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
         HeadLines       =   2
         RowHeight       =   19
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "numero_pago"
            Caption         =   "Numero de pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "concepto"
            Caption         =   "Concepto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "antecedente"
            Caption         =   "Antecente de pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "pago_aprobado"
            Caption         =   "Pago aprobado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "tipo_moneda"
            Caption         =   "Tipo Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "monto_dolares"
            Caption         =   "Monto Dolares"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "devengadogenerado"
            Caption         =   "Pago Aprobado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "codigo_orden"
            Caption         =   "Orden de pago"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            Locked          =   -1  'True
            BeginProperty Column00 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2670.236
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid6 
         Bindings        =   "ac_PagosMain.frx":17E4
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
   Begin MSAdodcLib.Adodc adoBeneficiarios 
      Height          =   330
      Left            =   5160
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "adoBeneficiarios"
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
   Begin MSAdodcLib.Adodc adoPagos 
      Height          =   330
      Left            =   2880
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "adoPagos"
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
   Begin Crystal.CrystalReport cr 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCancelBtn=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Orden de Pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   120
      Width           =   3720
   End
   Begin VB.Image Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   0
      Picture         =   "ac_PagosMain.frx":17FB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11730
   End
End
Attribute VB_Name = "ac_PagosMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------
'programa principal del submodulo de GESTION DE PAGOS                                          -
'se divide en tres partes                                                                      -
' 1  gestion de planillas                                                                      -
'       adiciona, modifica, elimina planillas de pago                                          -
' 2  administrador de pagos                                                                    -
'       adiciona, modifica, elimina pagos por planilla o individuales                          -
'       administra la actualizacion de conformidades de pago (F02) que lleguen de la UNIDAD    -
'       genera el comprometido de pago en base a la conformidad                                -
' 3  administrador de beneficiarios de cada pago                                               -
'       adiciona, modifica, elimina beneficiarios del pago                                     -
'       actualiza datos personales                                                             -
'       actualiza unidad a la que pertenece el beneficiario                                    -
'       establece el monto del pago para cada beneficiairo                                     -
'-----------------------------------------------------------------------------------------------
Option Explicit
Dim XCOLINDEX As Integer
Dim codigo_solicitudF02 As Integer

Private Sub adoBeneficiarios_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Call CuidaBotonesBeneficiario
'''Me.cmdEliminarBeneficiario.Enabled = False
'''Me.cmdRelAComprometido.Enabled = False
'''Me.cmdRelADevengado.Enabled = False
'''Me.cmdRelACYD.Enabled = False
'''Me.cmdCapMontos.Enabled = False
'''Me.cmdEmiteFactura.Enabled = False
'''Me.cmdConformidad.Enabled = False
'''If Me.adoBeneficiarios.Recordset!DEVENGADO <> "S" Then
'''    Me.cmdEliminarBeneficiario.Enabled = True
'''    Me.cmdRelAComprometido.Enabled = True
'''    Me.cmdRelADevengado.Enabled = True
'''    Me.cmdRelACYD.Enabled = True
'''    Me.cmdCapMontos.Enabled = True
'''    Me.cmdEmiteFactura.Enabled = True
'''    Me.cmdConformidad.Enabled = True
'''End If
End Sub

Private Sub adoGrupos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Call RefrescaPagos
End Sub


Private Sub adoPagos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Call RefrescaBeneficiarios
End Sub


Private Sub chkFiltroMain_Click()
'conmuta si o no se usa el filtro de unidades
If chkFiltroMain.Value = 0 Then
    Me.dtcCodigo_unidad.Enabled = False
Else
    Me.dtcCodigo_unidad.Enabled = True
End If
Call RefrescaListaPRincipal
End Sub

Private Sub cmdActConformidad_Click()
'procesa las conformidades recibidas mediante el cargado del F02
'actualiza el campo conformidad =S y emiteFactura=S/N
Dim Qantos%
Qantos = 0
If MsgBox("Desea procesar las conformidades recibidas ?", vbYesNo) = vbYes Then
    DE.dbo_edPagosActualizaConformidades Qantos
'    Call ImprimeReporteDeActualizacion
    If Qantos > 0 Then
        MsgBox "Se ha(n) procesado " & CStr(Qantos) & " registro(s) con todo éxito"
        Call RefrescaBeneficiarios
    Else
        MsgBox "Lo siento, no hubieron conformidades para procesar"
    End If
End If
End Sub

Private Sub cmdAdicionarGrupo_Click()
'adicionar un grupo (planilla) de pago
ac_PagosADiGrupo.Para_Codigo_Grupo = 0
ac_PagosADiGrupo.Show vbModal
Call RefrescaListaPRincipal
End Sub


Private Sub cmdAdicionarBeneficiario_Click()
'adiciona un beneficiario al pago
If Me.adoGrupos.Recordset!espagogrupo = "N" Then
    If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
        MsgBox "Solo se permite una persona en el grupo por que es una planilla individual", vbInformation, "Gestión de pagos"
        Exit Sub
    End If
End If
ac_PagosAdiMiembroGrupo.Para_det_correlativo = 0
ac_PagosAdiMiembroGrupo.Show vbModal
Call RefrescaBeneficiarios
End Sub


Private Sub cmdAdicionarPago_Click()
'adiciona un pago para la planilla correspondiente
ac_PagosAdiPago.Para_Numero_Pago = 0
ac_PagosAdiPago.Show vbModal
Call RefrescaPagos
End Sub


Private Sub cmdAprobadoParaEnvio_Click()
'marca el pago como "DETALLE DE BENEFICIARIOS ENVIADO A LA UNIDAD"
If Me.adoPagos.Recordset!AprobadoParaEnvioAUnidad = "N" Then
    If MsgBox("Desea aprobar el pago indicado para enviarlo a la unidad ejecutora ?", vbYesNo) = vbYes Then
        DE.dbo_edGeneralSearching "update ac_pagos_cronograma set AprobadoParaEnvioAUnidad='S' where ges_gestion='" & Me.adoPagos.Recordset!ges_gestion & "' and codigo_unidad='" & Me.adoPagos.Recordset!codigo_unidad & "' and codigo_grupo=" & Me.adoPagos.Recordset!codigo_grupo & " and numero_pago=" & Me.adoPagos.Recordset!numero_pago
        Call RefrescaPagos
    End If
Else
    If MsgBox("Desea anular la aprobación del envio del pago indicado ?", vbYesNo) = vbYes Then
        DE.dbo_edGeneralSearching "update ac_pagos_cronograma set AprobadoParaEnvioAUnidad='N' where ges_gestion='" & Me.adoPagos.Recordset!ges_gestion & "' and codigo_unidad='" & Me.adoPagos.Recordset!codigo_unidad & "' and codigo_grupo=" & Me.adoPagos.Recordset!codigo_grupo & " and numero_pago=" & Me.adoPagos.Recordset!numero_pago
        Call RefrescaPagos
    End If
End If
End Sub


Function TodoBienPAraDevengar() As Boolean
'valida que todos los datos esten completos para poder generar el devengado del comprometido
TodoBienPAraDevengar = True
If Me.adoBeneficiarios.Recordset.RecordCount < 1 Then
    MsgBox "No hay registrado ningún beneficiario"
    TodoBienPAraDevengar = False
End If
End Function


Private Sub cmdBorraGrupo_Click()
'elimina un grupo (planilla) de pago
If TodoBienPAraBorrarPlanilla() Then
    If MsgBox("Desea eliminar el grupo y el compromiso de pago relacionado ?", vbYesNo) = vbYes Then
        DE.dbo_edPagosBorraGrupo Me.adoGrupos.Recordset!ges_gestion, Me.adoGrupos.Recordset!codigo_unidad, Me.adoGrupos.Recordset!codigo_grupo
        Call RefrescaListaPRincipal
    End If
End If
End Sub

Function TodoBienPAraBorrarPlanilla() As Boolean
Dim rs As New ADODB.Recordset
Dim rsa As New ADODB.Recordset
Dim Qantos%
'para contar sus compromisos aprobados
    With Me.adoGrupos.Recordset
    rsa.Open "select q=count(*) from ac_ben_comprDeven where gp_ges_gestion='" & !ges_gestion & "' and gp_codigo_unidad='" & !codigo_unidad & "' and gp_codigo_grupo=" & !codigo_grupo & " and tipoComprobante='COM' and aprobotesoreria='S'", db, adOpenStatic, adLockReadOnly
    Qantos = rsa!q
    rsa.Close
    End With

TodoBienPAraBorrarPlanilla = True
If Qantos > 0 Then
    TodoBienPAraBorrarPlanilla = False
    MsgBox "No puede eliminar la planilla por tiene compromiso de pago APROBADO", vbCritical, "Atencion"
Else
    Set rs = Me.adoPagos.Recordset.Clone
    If rs.RecordCount > 0 Then
        rs.Find "devengadogenerado='S'"
        If Not rs.EOF Then
            TodoBienPAraBorrarPlanilla = False
            MsgBox "No puede eliminar una planilla que tiene ordenes de pago elaboradas", vbCritical, "Atencion"
        End If
    End If
End If
End Function

Private Sub cmdCapMontos_Click()
'edita los montos del pago asignado al beneficiario
'siempre que no exista la conformida de la unidad
If Me.adoBeneficiarios.Recordset!conformidad = "S" Then
    MsgBox "Lo siento, no puede modificar los montos del registro por que su pago cuenta con la conformidad de la unidad", vbCritical
Else
    ac_PagosAdiPagoPersona.Para_GES_GESTION = Me.adoBeneficiarios.Recordset!ges_gestion1
    ac_PagosAdiPagoPersona.Para_Codigo_Unidad = Me.adoBeneficiarios.Recordset!codigo_unidad1
    ac_PagosAdiPagoPersona.Para_Codigo_Grupo = Me.adoBeneficiarios.Recordset!codigo_grupo
    ac_PagosAdiPagoPersona.Para_Numero_Pago = Me.adoBeneficiarios.Recordset!numero_pago
    ac_PagosAdiPagoPersona.Para_IdFuncionario = Me.adoBeneficiarios.Recordset!idfuncionario
    ac_PagosAdiPagoPersona.Show vbModal
    Call RefrescaBeneficiarios
    Call RefrescaPagos
End If
End Sub



Private Sub cmdConformidad_Click()
'conmuta la bandera que indica si el beneficiario TIENE CONFORMIDAD O NO
If Me.adoPagos.Recordset!DEVENGADOGENERADO = "S" Then
    MsgBox "No se puede cambiar el estado de la bandera CONFORMIDAD por que ya se ha generado los Devengados del Compromiso de Pago", vbCritical, "Atencion"
Else
    With Me.adoBeneficiarios.Recordset
        If !conformidad = "S" Then
            If MsgBox("ATENCION" & Chr(13) & Chr(13) & "El beneficiario [" & !paterno & " " & !materno & " " & !NombreS & "] tiene registrado CONFORMIDAD = 'SI', desea cambiar a 'NO' ?", vbYesNo) = vbYes Then
                db.Execute "update ac_pagos_cronograma_detalle_1 set CONFORMIDAD='N' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
                Call RefrescaBeneficiarios
            End If
        Else
            If MsgBox("El beneficiario [" & !paterno & " " & !materno & " " & !NombreS & "] tiene registrado CONFORMIDAD='N', desea cambiar a 'SI' ?", vbYesNo) = vbYes Then
                db.Execute "update ac_pagos_cronograma_detalle_1 set CONFORMIDAD='S' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
                Call RefrescaBeneficiarios
            End If
        End If
    End With
End If
End Sub

Private Sub cmdCopiaCrono_Click()
Dim DES_ges_gestion$, DES_codigo_unidad$, DES_codigo_grupo%
    
If Me.adoGrupos.Recordset!espagogrupo = "N" And Me.adoPagos.Recordset.RecordCount = 0 Then
    With Me.adoGrupos.Recordset
    ac_PagosSelPlanillaPaClonar.DES_ges_gestion = !ges_gestion
    ac_PagosSelPlanillaPaClonar.DES_codigo_unidad = !codigo_unidad
    ac_PagosSelPlanillaPaClonar.DES_codigo_grupo = !codigo_grupo
    ac_PagosSelPlanillaPaClonar.Show vbModal
    End With
Else
    MsgBox "Este utilitario solo funciona con planillas individuales y con planillas que aun no tienen Cronograma de pagos definido", vbInformation
End If
End Sub

Private Sub cmdEliminarBeneficiario_Click()
'quita un beneficiario del pago indicado
If Me.adoBeneficiarios.Recordset!conformidad = "S" Then
    MsgBox "Lo siento, no puede eliminar el registro por que su pago cuenta con la conformidad de la unidad", vbCritical
Else
    If MsgBox("Desea eliminar el registro del Beneficiario ?", vbYesNo) = vbYes Then
        With Me.adoBeneficiarios.Recordset
            DE.dbo_edPagosBorraPagoPersona !ges_gestion1, !codigo_unidad1, !codigo_grupo, !numero_pago, !idfuncionario
        End With
        Call RefrescaBeneficiarios
        Call RefrescaPagos
    End If
End If
End Sub


Private Sub cmdEliminarPago_Click()
'elimina el pago incluyendo sus beneficiarios relacionados
Dim rsX As New ADODB.Recordset
Set rsX = Me.adoBeneficiarios.Recordset.Clone
If Not rsX.EOF Then
    rsX.MoveFirst
    rsX.Find "conformidad='S'"
    If Not rsX.EOF Then
        MsgBox "Lo siento, no puede eliminar el pago por que tiene personas que cuentan con la conformidad de la unidad", vbCritical
    Else
    If MsgBox("Desea eliminar el registro del pago ?", vbYesNo) = vbYes Then
        DE.dbo_edPagosBorraPago Me.adoPagos.Recordset!ges_gestion, Me.adoPagos.Recordset!codigo_unidad, Me.adoPagos.Recordset!codigo_grupo, Me.adoPagos.Recordset!numero_pago
        Call RefrescaPagos
    End If
    End If
Else
    If MsgBox("Desea eliminar el registro del pago ?", vbYesNo) = vbYes Then
        DE.dbo_edPagosBorraPago Me.adoPagos.Recordset!ges_gestion, Me.adoPagos.Recordset!codigo_unidad, Me.adoPagos.Recordset!codigo_grupo, Me.adoPagos.Recordset!numero_pago
        Call RefrescaPagos
    End If
End If
rsX.Close
End Sub


Private Sub cmdEmiteFactura_Click()
Dim Resp As Integer

'conmuta la bandera que indica si el beneficiario EMITE FACTURA O NO
If Me.adoPagos.Recordset!DEVENGADOGENERADO = "S" Then
    MsgBox "No se puede cambiar el estado de la bandera EMITE FACTURA por que ya se ha generado los Devengados del Compromiso de Pago", vbCritical, "Atencion"
Else
'    With Me.adoBeneficiarios.Recordset
'        If !emitefactura = "S" Then
'            If MsgBox("ATENCION" & Chr(13) & Chr(13) & "El beneficiario [" & !paterno & " " & !materno & " " & !nombres & "] indica que SI EMITE FACTURA, desea cambiar a NO EMITE FACTURA ?", vbYesNo) = vbYes Then
'                db.Execute "update ac_pagos_cronograma_detalle_1 set emitefactura='N' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
'                Call RefrescaBeneficiarios
'            End If
'        Else
'            If MsgBox("El beneficiario [" & !paterno & " " & !materno & " " & !nombres & "] indica que NO EMITE FACTURA, desea cambiar a SI EMITE FACTURA ?", vbYesNo) = vbYes Then
'                db.Execute "update ac_pagos_cronograma_detalle_1 set emitefactura='S' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
'                Call RefrescaBeneficiarios
'            End If
'        End If
'    End With
    With Me.adoBeneficiarios.Recordset
         If MsgBox("ATENCION" & Chr(13) & Chr(13) & "El beneficiario [" & !paterno & " " & !materno & " " & !NombreS & "] EMITE FACTURA?", vbYesNo) = vbYes Then
            db.Execute "update ac_pagos_cronograma_detalle_1 set emitefactura='S' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
            Call RefrescaBeneficiarios
        Else
            db.Execute "update ac_pagos_cronograma_detalle_1 set emitefactura='N' where codigo_grupo=" & !codigo_grupo & " and numero_pago=" & !numero_pago & " and idfuncionario=" & !idfuncionario
            Call RefrescaBeneficiarios
            
        End If
    End With
End If
End Sub

Private Sub cmdGenDevengado_Click()
'genera el devengado del comprometido para todo el pago
Dim HayConformidades As Boolean
If TodoBienPAraDevengar Then
    HayConformidades = False
    Me.adoBeneficiarios.Recordset.MoveFirst
    'recorre todos los beneficiarios para determinar si existen conformidades para procesar
    Do While Not Me.adoBeneficiarios.Recordset.EOF
        If Me.adoBeneficiarios.Recordset!conformidad = "S" And Me.adoBeneficiarios.Recordset!devengado <> "S" Then
            HayConformidades = True
            codigo_solicitudF02 = Me.adoBeneficiarios.Recordset!codigo_solicitudF02
            Exit Do
        End If
        Me.adoBeneficiarios.Recordset.MoveNext
    Loop
    If HayConformidades = False Then
        MsgBox "No se registraron conformidades para poder generar el Registro del Devengado", vbCritical, "Atencion"
    Else
        If MsgBox("Desea generar el Devengado del Compromiso de Pago ?", vbYesNo) = vbYes Then
            Call GeneraDevengado
            Call RefrescaPagos
        End If
    End If
    Me.adoBeneficiarios.Recordset.MoveFirst
End If
End Sub

Sub GeneraDevengado()
'genera el devengado de esa persona y lo graba a la tabla temporal
Dim montocontrol As Double
Dim sesion$
Dim Mal As Integer
Dim SeAsigno As Boolean
Dim rsOrden As New ADODB.Recordset
Dim rsTMP As New ADODB.Recordset
Dim rsc As New ADODB.Recordset
Dim co As New ADODB.Command
Mal = 0
sesion = Left("S" & CStr(Rnd()), 10)
rsTMP.Open "select * from ac_Ben_Devengado_TMP where sesion='" & sesion & "'", db, adOpenDynamic, adLockOptimistic
'recorre todos los beneficiarios para generar su devengado en tabla temporal
Me.adoBeneficiarios.Recordset.MoveFirst
Do While Not Me.adoBeneficiarios.Recordset.EOF
    If Me.adoBeneficiarios.Recordset!conformidad = "S" Then
        DE.dbo_apGeneralSearching "update ac_ben_comprdeven set monto_dolares_acum=0 where idfuncionario=" & Me.adoBeneficiarios.Recordset!idfuncionario
        If rsOrden.State = 1 Then rsOrden.Close
        rsOrden.Open "select distinct ordencomprobante From AC_BEN_COMPRDEVEN WHERE idFuncionario = " & Me.adoBeneficiarios.Recordset!idfuncionario & " and aprobotesoreria   = 'S' and tipocomprobante = 'COM' order by ordencomprobante", db, adOpenStatic, adLockReadOnly
        If rsOrden.RecordCount = 0 Then
            Mal = 1
        Else
            SeAsigno = False
            Do While Not rsOrden.EOF
                If HaySaldoParaDevengar(Me.adoBeneficiarios.Recordset!idfuncionario, rsOrden!ordenComprobante) Then
                    If HayEspacioEnComprometido(sesion, rsOrden!ordenComprobante) Then
                        'devenga lo que se pueda del comprometido --> TMP
                        If COMEstaSoloEs100yEs258o222(Me.adoBeneficiarios.Recordset!idfuncionario, rsOrden!ordenComprobante) And glProceso = "CONSULTORIA" Then
                            DE.dbo_edGeneraDevEnTmp100 sesion, Me.adoBeneficiarios.Recordset!idfuncionario, rsOrden!ordenComprobante, Me.adoBeneficiarios.Recordset!codigo_unidad, Me.adoBeneficiarios.Recordset!codigo_grupo, Me.adoBeneficiarios.Recordset!numero_pago, Me.adoBeneficiarios.Recordset!emitefactura
                        Else
                            DE.dbo_edGeneraDevengadoEnTmp sesion, Me.adoBeneficiarios.Recordset!idfuncionario, rsOrden!ordenComprobante, Me.adoBeneficiarios.Recordset!codigo_unidad, Me.adoBeneficiarios.Recordset!codigo_grupo, Me.adoBeneficiarios.Recordset!numero_pago, Me.adoBeneficiarios.Recordset!emitefactura
                        End If
                        SeAsigno = True
                    End If
                End If
                rsOrden.MoveNext
            Loop
            If SeAsigno = False Then
                Mal = 2
            End If
        End If
    End If
    'toma siguiente beneficiario
    Me.adoBeneficiarios.Recordset.MoveNext
    If Mal > 0 Then Exit Do
Loop
If Mal = 0 Then
    'este SP genera el devengado en base a la tabla ac_ben_devengado_tmp teniendo como aghrupador la sesión
    co.ActiveConnection = db
    co.CommandText = "edGeneraDevengado"
    co.CommandType = adCmdStoredProc
    co.Parameters("@sesion") = sesion
    co.Parameters("@GP_ges_gestion") = Me.adoPagos.Recordset!ges_gestion
    co.Parameters("@GP_codigo_unidad") = Me.adoPagos.Recordset!codigo_unidad
    co.Parameters("@GP_codigo_grupo") = Me.adoPagos.Recordset!codigo_grupo
    co.Parameters("@GP_numero_pago") = Me.adoPagos.Recordset!numero_pago
    co.Parameters("@codigo_solicitud") = codigo_solicitudF02
    co.Parameters("@usr_usuario") = GlUsuario
    co.Execute
    MsgBox "Se ha generado el Devengado del Comprometido con todo éxito", vbInformation, "Atencion"
ElseIf Mal = 1 Then
    MsgBox "Existe conformidad de parte de la Unidad, pero el Compromiso de Pago no está aprobado", vbCritical, "Atencion"
ElseIf Mal = 2 Then
    MsgBox "No se genero Devengado por que existe error en los saldos del Compromiso de Pago, revise por favor", vbCritical, "Atencion"
End If
'elimina registros de la sesion en la tabla temporal
DE.dbo_apGeneralSearching "delete ac_ben_devengado_tmp where sesion='" & sesion & "'"
Me.adoBeneficiarios.Recordset.MoveFirst
End Sub

Function COMEstaSoloEs100yEs258o222(idf As Integer, oc As Integer) As Boolean
Dim Respuesta$
COMEstaSoloEs100yEs258o222 = False
If glProceso = "CONSULTORIA" Then
    DE.dbo_edComSolo100y258o222 Me.adoPagos.Recordset!codigo_unidad, Me.adoPagos.Recordset!codigo_grupo, idf, oc, Respuesta
    If Respuesta = "S" Then COMEstaSoloEs100yEs258o222 = True
End If
End Function

Function HaySaldoParaDevengar(idfuncionario As Integer, ordenComprobante As Integer) As Boolean
'verifica si existe saldo del comprometido para devengar
HaySaldoParaDevengar = False
Dim rs As New ADODB.Recordset
rs.Open "select saldo = monto_dolares - monto_dolares_acum from ac_ben_comprdeven where idfuncionario=" & idfuncionario & " and ordencomprobante=" & ordenComprobante & " and tipocomprobante='COM' and aprobotesoreria='S' AND GP_GES_GESTION='" & Me.adoPagos.Recordset!ges_gestion & "' AND gp_codigo_unidad='" & Me.adoPagos.Recordset!codigo_unidad & "' and  GP_CODIGO_GRUPO=" & Me.adoPagos.Recordset!codigo_grupo, db, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    Do While Not rs.EOF
        If rs!saldo > 0 Then HaySaldoParaDevengar = True
        rs.MoveNext
    Loop
End If
rs.Close
End Function

Function HayEspacioEnComprometido(sesion As String, ordenComprobante As Integer) As Boolean
'determina si existe espacio en el comprometido considerando los devengados benerados en la tabla temporal
HayEspacioEnComprometido = True
Dim rsc As New ADODB.Recordset
Dim rsR As New ADODB.Recordset
Dim rsT As New ADODB.Recordset
Dim montoTMP As Double

rsR.Open "SELECT ges_gestion, org_codigo, codigo_pago " & _
         "From AC_BEN_COMPRDEVEN WHERE   idFuncionario     = " & Me.adoBeneficiarios.Recordset!idfuncionario & " and " & _
                                        "aprobotesoreria   = 'S'         and " & _
                                        "tipocomprobante   = 'COM'       and " & _
                                        "ordencomprobante  = " & ordenComprobante & " and " & _
                                        "gp_ges_Gestion    = '" & Me.adoPagos.Recordset!ges_gestion & "' and " & _
                                        "gp_codigo_unidad  = '" & Me.adoPagos.Recordset!codigo_unidad & "' and " & _
                                        "gp_codigo_grupo   = " & Me.adoPagos.Recordset!codigo_grupo & " " & _
         "order by ges_gestion, org_codigo desc, codigo_pago", db, adOpenStatic, adLockReadOnly
If rsR.RecordCount > 0 Then
    rsc.Open "SELECT monto_Dolares " & _
         "From pagos WHERE ges_Gestion = '" & rsR!ges_gestion & "' and " & _
                                 "org_codigo = '" & rsR!org_codigo & "' and " & _
                                 "codigo_pago = " & rsR!codigo_pago & " and " & _
                                 "tipo_formulario = 'COM' ", db, adOpenStatic, adLockReadOnly
    If rsc.EOF Then
        HayEspacioEnComprometido = False
    Else
        rsT.Open "SELECT monto_dolares From ac_ben_devengado_TMP " & _
                 "where sesion       = '" & sesion & "' and " & _
                       "Cges_gestion = '" & rsR!ges_gestion & "' and " & _
                       "Corg_codigo  = '" & rsR!org_codigo & "' and " & _
                       "Ccodigo_pago = " & rsR!codigo_pago, db, adOpenStatic, adLockReadOnly
        If rsT.RecordCount > 0 Then
            montoTMP = rsT!monto_dolares
        Else
            montoTMP = 0
        End If
        If rsc!monto_dolares - montoTMP > 0 Then
            HayEspacioEnComprometido = True
        Else
            HayEspacioEnComprometido = False
        End If
    End If
Else
    HayEspacioEnComprometido = False
End If
End Function

Private Sub cmdGenDEVHist_Click()
''''''''genera el devengado del comprometido para todo el pago
''''''''este procedimiento enlaza a los comprometidos manualmente
''''''''solo sirve para consultorias individuales
'''''''Dim HayConformidades As Boolean
'''''''Dim MontoTotADevengar As Double
'''''''If TodoBienPAraDevengar Then
''''''''    If Me.adoBeneficiarios.Recordset.RecordCount > 1 Then
''''''''        MsgBox "Se puede emplear este procedimiento solo en planillas individuales", vbInformation, "Atencion"
''''''''    Else
'''''''        HayConformidades = False
'''''''        Me.adoBeneficiarios.Recordset.MoveFirst
'''''''        'recorre todos los beneficiarios para determinar si existen conformidades para procesar
'''''''        Do While Not Me.adoBeneficiarios.Recordset.EOF
'''''''            If Me.adoBeneficiarios.Recordset!conformidad = "S" And Me.adoBeneficiarios.Recordset!devengado <> "S" Then
'''''''                HayConformidades = True
'''''''                MontoTotADevengar = MontoTotADevengar + Me.adoBeneficiarios.Recordset!monto_dol_ext + Me.adoBeneficiarios.Recordset!monto_dol_nal
'''''''                Exit Do
'''''''            End If
'''''''            Me.adoBeneficiarios.Recordset.MoveNext
'''''''        Loop
'''''''        Me.adoBeneficiarios.Recordset.MoveFirst
'''''''        If HayConformidades = False Then
'''''''            MsgBox "No se registraron conformidades para poder generar el Registro del Devengado", vbCritical, "Atencion"
'''''''        Else
'''''''            Hist_GeneraDevengado.Para_MontoADevengar = MontoTotADevengar
'''''''            Hist_GeneraDevengado.Show vbModal
'''''''            Call RefrescaPagos
'''''''''''''            'marca devengadogenerado=S en pagos
'''''''''''''            DE.dbo_edPagosMarcaPagoDevengado Me.adoPagos.Recordset!ges_gestion, Me.adoPagos.Recordset!codigo_unidad, Me.adoPagos.Recordset!codigo_grupo, Me.adoPagos.Recordset!numero_pago
'''''''''''''            'marca devengado=S para cada beneficiario del pago
'''''''''''''            DE.dbo_edGeneralSearching "update ac_pagos_cronograma_detalle_1 set Devengado='S' where ges_gestion='" & Me.adoPagos.Recordset!ges_gestion & "' and codigo_unidad='" & Me.adoPagos.Recordset!codigo_unidad & "' and codigo_grupo=" & Me.adoPagos.Recordset!codigo_grupo & " and numero_pago=" & Me.adoPagos.Recordset!numero_pago & " and conformidad='S'"
'''''''''''''            Call RefrescaPagos
'''''''        End If
'''''''    End If
''''''''End If
End Sub


Private Sub cmdGenF02_Click()
'ac_PagosGenF02.Show vbModal
End Sub

Private Sub cmdImport_Click()
'ac_ImportaDesde2000.Show vbModal
'Call refrescaListaPrincipal
End Sub

Private Sub cmdPrDetHonorarios_Click()
ac_PagosPrintDetHonor.Show vbModal
End Sub

Private Sub cmdPrintOP_Click()
ac_PagosPrintOrdenPago.Show vbModal
End Sub


Private Sub cmdPrintPago_Click()
'imprime el detalle del pago
If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
    Dim IResult As Variant, i%
    CR.StoredProcParam(0) = Me.adoPagos.Recordset!ges_gestion
    CR.StoredProcParam(1) = Me.adoPagos.Recordset!codigo_unidad
    CR.StoredProcParam(2) = Val(Me.adoPagos.Recordset!codigo_grupo)
    CR.StoredProcParam(3) = Val(Me.adoPagos.Recordset!numero_pago) 'para que filtre unicamente el pago indicado
    CR.ReportFileName = App.Path & "\consultoria edson\rptDetPlanilla.rpt"
    IResult = CR.PrintReport
    If IResult <> 0 Then MsgBox CR.LastErrorNumber & " : " & CR.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No existen registros para imprimir", vbInformation, "Atencion"
End If
End Sub

Private Sub cmdPrintPlanilla_Click()
'imprime el detalle de la planilla agrupados por pagos
Dim IResult As Variant, i%
If Me.adoPagos.Recordset.RecordCount > 0 Then
    CR.StoredProcParam(0) = Me.adoPagos.Recordset!ges_gestion
    CR.StoredProcParam(1) = Me.adoPagos.Recordset!codigo_unidad
    CR.StoredProcParam(2) = Val(Me.adoPagos.Recordset!codigo_grupo)
    CR.StoredProcParam(3) = 0   ' ES PARA QUE NO FILTRE UN SOLO PAGO SINO MUESTRE TODA LA PLANILLA
    CR.WindowShowGroupTree = True
    
    If Me.adoGrupos.Recordset!espagogrupo = "S" Then
        CR.ReportFileName = App.Path & "\consultoria edson\rptDetPlanilla.rpt"
    Else
        CR.ReportFileName = App.Path & "\consultoria edson\rptDetPlanillaUni.rpt"
    End If
    IResult = CR.PrintReport
    If IResult <> 0 Then MsgBox CR.LastErrorNumber & " : " & CR.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No existen registros para imprimir", vbInformation, "Atencion"
End If
End Sub

Private Sub cmdJalaMismasPersonas_Click()
'toma los beneficiarios del pago anterior y los copia al pago indicado
If MsgBox("Desea copiar los mismos beneficiarios del pago [" & CStr(Me.adoPagos.Recordset!numero_pago - 1) & "] ?", vbYesNo) = vbYes Then
    With Me.adoPagos.Recordset
        DE.dbo_edPagosClonaBeneficiarios !ges_gestion, !codigo_unidad, !codigo_grupo, !numero_pago - 1, !numero_pago
        Call RefrescaBeneficiarios
    End With
End If
End Sub


Private Sub cmdModificarGrupo_Click()
'para modificar los datos del grupo (planilla)
ac_PagosADiGrupo.Para_Codigo_Grupo = Me.adoGrupos.Recordset!codigo_grupo
ac_PagosADiGrupo.Show vbModal
Call RefrescaListaPRincipal
End Sub


Private Sub cmdModificarPago_Click()
'modifica los datos del pago
ac_PagosAdiPago.Para_Codigo_Grupo = Me.adoPagos.Recordset!codigo_grupo
ac_PagosAdiPago.Para_Numero_Pago = Me.adoPagos.Recordset!numero_pago
ac_PagosAdiPago.Show vbModal
Call RefrescaPagos
End Sub

Private Sub cmdRelAComprometido_Click()
''''''relaciona a la perosna con un registro de comprobante de pago existente
'''''Dim rs As New ADODB.Recordset
'''''rs.Open "select p.ges_Gestion, p.codigo_unidad, p.codigo_solicitud from pagos p, ac_ben_comprdeven b " & _
'''''         "where   p.ges_Gestion   = b.ges_Gestion and " & _
'''''         "p.org_codigo    = b.org_codigo  and " & _
'''''         "p.codigo_pago   = b.codigo_pago and " & _
'''''         "b.tipocomprobante= 'COM' and " & _
'''''         "b.idfuncionario = " & Me.adoBeneficiarios.Recordset!idfuncionario, db, adOpenStatic, adLockReadOnly
'''''If rs.RecordCount > 0 Then
'''''    Hist_EnganchaComprometido.Para_GES_GESTION = rs!GES_GESTION
'''''    Hist_EnganchaComprometido.Para_Codigo_Unidad = rs!codigo_unidad
'''''    Hist_EnganchaComprometido.Para_codigo_solicitud = rs!codigo_solicitud
'''''End If
'''''rs.Close
'''''Hist_EnganchaComprometido.Para_deDonde = 0
'''''Hist_EnganchaComprometido.Para_tipocomprobante = "COM"
'''''Hist_EnganchaComprometido.Show vbModal
''''''Hist_RelacionaAlComprometido.Show vbModal
End Sub

Private Sub cmdRelACYD_Click()
''''''''relaciona a la perosna con un registro de comprobante de pago existente
'''''''Dim rs As New ADODB.Recordset
'''''''rs.Open "select p.ges_Gestion, p.codigo_unidad, p.codigo_solicitud from pagos p, ac_ben_comprdeven b " & _
'''''''         "where   p.ges_Gestion   = b.ges_Gestion and " & _
'''''''         "p.org_codigo    = b.org_codigo  and " & _
'''''''         "p.codigo_pago   = b.codigo_pago and " & _
'''''''         "b.tipocomprobante= 'CYD' and " & _
'''''''         "b.idfuncionario = " & Me.adoBeneficiarios.Recordset!idfuncionario, db, adOpenStatic, adLockReadOnly
'''''''If rs.RecordCount > 0 Then
'''''''    Hist_EnganchaComprometido.Para_GES_GESTION = rs!GES_GESTION
'''''''    Hist_EnganchaComprometido.Para_Codigo_Unidad = rs!codigo_unidad
'''''''    Hist_EnganchaComprometido.Para_codigo_solicitud = rs!codigo_solicitud
'''''''End If
'''''''rs.Close
'''''''Hist_EnganchaComprometido.Para_deDonde = 0
'''''''Hist_EnganchaComprometido.Para_tipocomprobante = "CYD"
'''''''Hist_EnganchaComprometido.Para_MontoADevengar = Me.adoBeneficiarios.Recordset!monto_dol_ext + Me.adoBeneficiarios.Recordset!monto_dol_nal
'''''''Hist_EnganchaComprometido.Show vbModal
''''''''Hist_RelacionaAlComprometido.Show vbModal
End Sub

Private Sub cmdRelADevengado_Click()
'''''''relaciona a la perosna con un registro de comprobante de pago existente
''''''Dim rs As New ADODB.Recordset
''''''rs.Open "select p.ges_Gestion, p.codigo_unidad, p.codigo_solicitud from pagos p, ac_ben_comprdeven b " & _
''''''         "where   p.ges_Gestion   = b.ges_Gestion and " & _
''''''         "p.org_codigo    = b.org_codigo  and " & _
''''''         "p.codigo_pago   = b.codigo_pago and " & _
''''''         "b.tipocomprobante= 'DEV' and " & _
''''''         "b.idfuncionario = " & Me.adoBeneficiarios.Recordset!idfuncionario, db, adOpenStatic, adLockReadOnly
''''''If rs.RecordCount > 0 Then
''''''    Hist_EnganchaComprometido.Para_GES_GESTION = rs!GES_GESTION
''''''    Hist_EnganchaComprometido.Para_Codigo_Unidad = rs!codigo_unidad
''''''    Hist_EnganchaComprometido.Para_codigo_solicitud = rs!codigo_solicitud
''''''End If
''''''rs.Close
''''''Hist_EnganchaComprometido.Para_deDonde = 0
''''''Hist_EnganchaComprometido.Para_tipocomprobante = "DEV"
''''''Hist_EnganchaComprometido.Para_MontoADevengar = Me.adoBeneficiarios.Recordset!monto_dol_ext + Me.adoBeneficiarios.Recordset!monto_dol_nal
''''''Hist_EnganchaComprometido.Show vbModal
'''''''Hist_RelacionaAlComprometido.Show vbModal
End Sub

Private Sub cmdRelCronoCompro_Click()
If Me.adoGrupos.Recordset.RecordCount > 0 Then
    If Not (Me.adoGrupos.Recordset.EOF Or Me.adoGrupos.Recordset.BOF) Then
        HistRelCronoCompro.xGes_Gestion = Me.adoGrupos.Recordset!ges_gestion
        HistRelCronoCompro.xCodigo_Unidad = Me.adoGrupos.Recordset!codigo_unidad
        HistRelCronoCompro.xCodigo_Grupo = Me.adoGrupos.Recordset!codigo_grupo
    End If
End If
If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
    If Not (Me.adoBeneficiarios.Recordset.EOF Or Me.adoBeneficiarios.Recordset.BOF) Then
        HistRelCronoCompro.xNumero_Pago = Me.adoBeneficiarios.Recordset!numero_pago
        HistRelCronoCompro.XIdFuncionario = Me.adoBeneficiarios.Recordset!idfuncionario
    End If
End If
HistRelCronoCompro.Show vbModal
End Sub

Private Sub cmdUp_Click()
'''''mueve hacia arriba el registr de la lista
''''Dim a1%, b1%
''''With Me.adoPagos.Recordset
''''If .RecordCount > 1 Then
''''    .MovePrevious
''''    If .BOF Then
''''        .MoveFirst
''''    Else
''''        a1 = !numero_pago
''''        .MoveNext
''''        b1 = !numero_pago
''''        DE.dbo_edPagosSwap a1, b1
''''        Call RefrescaPagos
''''    End If
''''End If
''''End With
End Sub

Private Sub cmdDown_Click()
'''''mueve hacia abajo el registro
''''Dim a1%, b1%
''''With Me.adoPagos.Recordset
''''If .RecordCount > 1 Then
''''    .MoveNext
''''    If .EOF Then
''''        .MoveLast
''''    Else
''''        a1 = !numero_pago
''''        .MovePrevious
''''        b1 = !numero_pago
''''        DE.dbo_edPagosSwap a1, b1
''''        Call RefrescaPagos
''''    End If
''''End If
''''End With
End Sub

Private Sub dgGrupos_HeadClick(ByVal ColIndex As Integer)
XCOLINDEX = ColIndex
Call RefrescaListaPRincipal
End Sub

Private Sub dtcCodigo_unidad_Click(Area As Integer)
Call RefrescaListaPRincipal
End Sub

Private Sub Form_Load()
If glProceso = "CONSULTORIA" Then
    Me.Caption = "Consultoría - Pagos Principal"
     Me.cmdPrintOP.Visible = True
'    Me.cmdPrintOP.Visible = False
    Me.cmdPrDetHonorarios.Visible = False
    'JQA
    'Me.cmdGenF02.Visible = False
    Me.cmdGenF02.Visible = True
Else
    Me.Caption = "Recursos Humanos - Pagos Principal"
    Me.cmdPrintOP.Visible = True
    Me.cmdPrDetHonorarios.Visible = True
    Me.cmdGenF02.Visible = True
End If
XCOLINDEX = 0
Me.cmdAdicionarBeneficiario.Enabled = False
Me.cmdAdicionarPago.Enabled = False
'Me.cmdAprobarPago.Enabled = False
Me.cmdBorraGrupo.Enabled = False
Me.cmdCopiaCrono.Enabled = False
Me.cmdCapMontos.Enabled = False
Me.cmdEmiteFactura.Enabled = False
Me.cmdConformidad.Enabled = False
Me.cmdEliminarBeneficiario.Enabled = False
Me.cmdRelAComprometido.Enabled = False
Me.cmdRelADevengado.Enabled = False
Me.cmdRelACYD.Enabled = False
Me.cmdEliminarPago.Enabled = False
Me.cmdAprobadoParaEnvio.Enabled = False
Me.cmdJalaMismasPersonas.Enabled = False
Me.cmdModificarGrupo.Enabled = False
Me.cmdModificarPago.Enabled = False
'JQA
'Me.cmdGenF02.Enabled = False
Me.cmdGenF02.Enabled = True
'fadeform Me, -1, -1, -1
Call RefrescaListaPRincipal
End Sub

Private Sub RefrescaListaPRincipal()
'refresca el adoGrupos (plainllas)
Dim orden As String
Dim esparaRH As String
If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
If XCOLINDEX < 3 Then
    orden = " order by ges_Gestion, codigo_unidad, codigo_grupo "
Else
    orden = " order by " & Me.dgGrupos.Columns(XCOLINDEX).DataField
End If
If glProceso = "CONSULTORIA" Then
    esparaRH = "N"
Else
    esparaRH = "S"
End If
If Me.chkFiltroMain.Value = 0 Then
    DE.dbo_apGeneralSearching "select * from ac_pagos_grupos where esparaRH='" & esparaRH & "' " & orden
Else
    DE.dbo_apGeneralSearching "select * from ac_pagos_grupos where esparaRH='" & esparaRH & "' and codigo_unidad='" & Me.dtcCodigo_unidad & "' " & orden
End If
With DE.rsdbo_apGeneralSearching
    Set Me.adoGrupos.Recordset = .Clone
    If DE.rsdbo_apGeneralSearching.State = 1 Then DE.rsdbo_apGeneralSearching.Close
    Call CuidaBotonesPrincipal
End With
End Sub

Sub RefrescaPagos()
'refresca la lista de los pagos determinados para la planilla
Dim rsp As New ADODB.Recordset
Dim fnp%
If Me.adoGrupos.Recordset.RecordCount > 0 Then
    If Me.adoPagos.Caption = "OK" Then
        If Me.adoPagos.Recordset.RecordCount > 0 Then
            fnp = Me.adoPagos.Recordset!numero_pago
        End If
    End If
    If DE.rsdbo_edGeneralSearching.State = 1 Then DE.rsdbo_edGeneralSearching.Close
    DE.dbo_edGeneralSearching "select c.* from ac_pagos_cronograma c where ges_gestion='" & Me.adoGrupos.Recordset!ges_gestion & "' and codigo_unidad='" & Me.adoGrupos.Recordset!codigo_unidad & "' and codigo_grupo=" & Me.adoGrupos.Recordset!codigo_grupo & " order by numero_pago"
    With DE.rsdbo_edGeneralSearching
        Me.adoPagos.Caption = "OK"
        Set rsp = .Clone
        Set Me.adoPagos.Recordset = .Clone
        If DE.rsdbo_edGeneralSearching.State = 1 Then DE.rsdbo_edGeneralSearching.Close
        If rsp.RecordCount > 0 Then
            rsp.Find "numero_pago=" & fnp
            If Not rsp.EOF Then
                Me.adoPagos.Recordset.Bookmark = rsp.Bookmark
            End If
        End If
    End With
    Call CuidaBotonesPagos
End If
End Sub

Sub RefrescaBeneficiarios()
'refresca la lista de los beneficiarios del pago indicado
Dim rsf As New ADODB.Recordset
Dim fif%
Me.cmdGenDevengado.Enabled = False
'JQA
'Me.cmdGenF02.Enabled = False
Me.cmdGenF02.Enabled = True
Me.cmdPrintOP.Enabled = False
Me.cmdPrDetHonorarios.Enabled = False
If Me.adoPagos.Recordset.RecordCount > 0 Then
        Me.cmdGenDevengado.Enabled = True
        Me.cmdPrintOP.Enabled = True
        Me.cmdPrDetHonorarios.Enabled = True
    If Me.adoPagos.Recordset!AprobadoParaEnvioAUnidad = "S" Then
        If Me.adoPagos.Recordset!enviadoalaunidad = "S" Then
            cmdAprobadoParaEnvio.Caption = "Enviado..."
        Else
            cmdAprobadoParaEnvio.Caption = "Anular envío"
        End If
    Else
        cmdAprobadoParaEnvio.Caption = "Aprobar Envío"
    End If
    If Me.adoBeneficiarios.Caption = "OK" Then
        If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
            fif = Me.adoBeneficiarios.Recordset!idfuncionario
        End If
    End If
    If DE.rsdbo_edGeneralSearching.State = 1 Then DE.rsdbo_edGeneralSearching.Close
           DE.dbo_edGeneralSearching "select cd.ges_Gestion ges_gestion1, cd.codigo_unidad codigo_unidad1, cd.codigo_grupo, cd.numero_pago, cd.tipo_moneda, cd.monto_bs_ext, cd.monto_dol_ext,cd.monto_bs_nal,cd.monto_dol_nal, cd.conformidad, cd.ncite_conformidad, cd.fcite_conformidad, cd.org_codigo_ext, cd.emitefactura, cd.devengado, cd.numero_consultoriahist as nchist, cd.fte_financiamientohist as ftehist, cd.codigo_solicitudf02, p.* from ac_pagos_cronograma_detalle_1 cd, rc_personal p where cd.ges_gestion='" & Me.adoPagos.Recordset!ges_gestion & "' and cd.codigo_unidad='" & Me.adoPagos.Recordset!codigo_unidad & "' and cd.codigo_grupo='" & Me.adoPagos.Recordset!codigo_grupo & "' and cd.numero_pago=" & Me.adoPagos.Recordset!numero_pago & " and cd.idfuncionario=p.idfuncionario order by p.paterno, p.materno, p.nombres"
    With DE.rsdbo_edGeneralSearching
        Me.adoBeneficiarios.Caption = "OK"
        Set Me.adoBeneficiarios.Recordset = .Clone
        Set rsf = .Clone
        If DE.rsdbo_edGeneralSearching.State = 1 Then DE.rsdbo_edGeneralSearching.Close
        If rsf.RecordCount > 0 Then
            rsf.Find "idfuncionario=" & fif
            If Not rsf.EOF Then
                Me.adoBeneficiarios.Recordset.Bookmark = rsf.Bookmark
            End If
        End If
        
    End With
    Call CuidaBotonesBeneficiario
Else
    Dim rsn As New ADODB.Recordset
    rsn.Open "select * from ac_pagos_grupos where codigo_grupo=99999", db, adOpenStatic, adLockReadOnly
    Set Me.adoBeneficiarios.Recordset = rsn
End If
End Sub

Sub CuidaBotonesPrincipal()
'habilita o deshabilita los botones del grupo de planillas
Dim rsNada As New ADODB.Recordset
Me.cmdBorraGrupo.Enabled = False
Me.cmdCopiaCrono.Enabled = False
Me.cmdModificarGrupo.Enabled = False
Me.cmdAdicionarPago.Enabled = False
    If Me.adoGrupos.Recordset.RecordCount > 0 Then
        Me.cmdBorraGrupo.Enabled = True
        Me.cmdCopiaCrono.Enabled = True
        Me.cmdModificarGrupo.Enabled = True
        Me.cmdAdicionarPago.Enabled = True
        Me.cmdAdicionarPago.Enabled = True
    Else
        rsNada.Open "select * from ac_pagos_grupos where codigo_grupo=999999", db, adOpenStatic, adLockReadOnly
        Set Me.adoBeneficiarios.Recordset = rsNada
        Set Me.adoPagos.Recordset = rsNada
    End If
End Sub

Sub CuidaBotonesPagos()
'habilita o deshabilita los botones del grupo de pagos
Me.cmdModificarPago.Enabled = False
Me.cmdEliminarPago.Enabled = False
Me.cmdAprobadoParaEnvio.Enabled = False
Me.cmdAdicionarBeneficiario.Enabled = False
Me.cmdPrintPago.Enabled = False
Me.cmdGenDevengado.Enabled = False
Me.cmdPrintOP.Enabled = False
Me.cmdPrDetHonorarios.Enabled = False
Me.cmdModificarPago.Enabled = True
'JQA
'Me.cmdGenF02.Enabled = False
Me.cmdGenF02.Enabled = True
If Me.adoPagos.Recordset.RecordCount > 0 Then
    If Me.adoPagos.Recordset!DEVENGADOGENERADO = "N" Then
        Me.cmdGenF02.Enabled = True
    End If
'    Me.cmdGenF02.Enabled = True
    Me.cmdEliminarPago.Enabled = True
    Me.cmdAprobadoParaEnvio.Enabled = True
    Me.cmdAdicionarBeneficiario.Enabled = True
    Me.cmdPrintPago.Enabled = True
    Me.cmdGenDevengado.Enabled = True
    Me.cmdPrintOP.Enabled = True
    Me.cmdPrDetHonorarios.Enabled = True
End If
End Sub

Sub CuidaBotonesBeneficiario()
'habilita o deshabilita los botones del grupo beneficiarios
On Error GoTo final
Me.cmdJalaMismasPersonas.Enabled = False
Me.cmdEliminarBeneficiario.Enabled = False
Me.cmdRelAComprometido.Enabled = False
Me.cmdRelADevengado.Enabled = False
Me.cmdRelACYD.Enabled = False
Me.cmdCapMontos.Enabled = False
Me.cmdEmiteFactura.Enabled = False
Me.cmdConformidad.Enabled = False
'JQA
'Me.cmdGenF02.Enabled = False
Me.cmdGenF02.Enabled = True
If Me.adoPagos.Recordset!DEVENGADOGENERADO <> "S" Then
    Me.cmdGenF02.Enabled = True
    If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
        Me.cmdEliminarBeneficiario.Enabled = True
        Me.cmdRelAComprometido.Enabled = True
        Me.cmdRelADevengado.Enabled = True
        Me.cmdRelACYD.Enabled = True
        Me.cmdCapMontos.Enabled = True
        Me.cmdEmiteFactura.Enabled = True
        Me.cmdConformidad.Enabled = True
    Else
        If Me.adoPagos.Recordset!numero_pago > 1 Then
            Me.cmdJalaMismasPersonas.Enabled = True
        End If
    End If
Else
    If Me.adoBeneficiarios.Recordset.RecordCount > 0 Then
        If Me.adoBeneficiarios.Recordset!devengado <> "S" Then
            Me.cmdEliminarBeneficiario.Enabled = True
            Me.cmdRelAComprometido.Enabled = True
            Me.cmdRelADevengado.Enabled = True
            Me.cmdRelACYD.Enabled = True
            Me.cmdCapMontos.Enabled = True
            Me.cmdEmiteFactura.Enabled = True
            Me.cmdConformidad.Enabled = True
        End If
    End If
End If
final:
End Sub

Private Sub labCerrar_Click()
Unload Me
End Sub

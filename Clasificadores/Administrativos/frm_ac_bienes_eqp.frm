VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_ac_bienes_eqp 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - Administrativos -  Equipos"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   120
   ClientWidth     =   11145
   Icon            =   "frm_ac_bienes_eqp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11145
   WindowState     =   2  'Maximized
   Begin VB.Frame FraInsumo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "INSUMOS PARA EL CRONOGRAMA POR CADA EQUIPO"
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
      ForeColor       =   &H000040C0&
      Height          =   4575
      Left            =   6120
      TabIndex        =   105
      Top             =   960
      Visible         =   0   'False
      Width           =   10695
      Begin VB.PictureBox FraGrabarDet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   10470
         TabIndex        =   131
         Top             =   3720
         Width           =   10500
         Begin VB.PictureBox CmdGrabaDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3720
            Picture         =   "frm_ac_bienes_eqp.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   133
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox CmdCancelaDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5160
            Picture         =   "frm_ac_bienes_eqp.frx":11D8
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   132
            Top             =   0
            Width           =   1400
         End
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8280
         TabIndex        =   115
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox Txt_cant1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad1"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   114
         Text            =   "0"
         Top             =   885
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8280
         TabIndex        =   113
         Top             =   1380
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8280
         TabIndex        =   112
         Top             =   1860
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8280
         TabIndex        =   111
         Top             =   2340
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   8280
         TabIndex        =   110
         Top             =   2820
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txt_cant2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad2"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   109
         Text            =   "0"
         Top             =   1365
         Width           =   855
      End
      Begin VB.TextBox Txt_cant3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad3"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   108
         Text            =   "0"
         Top             =   1845
         Width           =   855
      End
      Begin VB.TextBox Txt_cant4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad4"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   107
         Text            =   "0"
         Top             =   2325
         Width           =   855
      End
      Begin VB.TextBox Txt_cant5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "cantidad5"
         DataSource      =   "Ado_datos2"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8880
         TabIndex        =   106
         Text            =   "0"
         Top             =   2805
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtc_codigo6Z 
         Bindings        =   "frm_ac_bienes_eqp.frx":1AC4
         DataField       =   "bien_codigo1"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6600
         TabIndex        =   116
         Top             =   885
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6A 
         Bindings        =   "frm_ac_bienes_eqp.frx":1ADD
         DataField       =   "bien_codigo2"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6600
         TabIndex        =   117
         Top             =   1365
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6B 
         Bindings        =   "frm_ac_bienes_eqp.frx":1AF6
         DataField       =   "bien_codigo3"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6600
         TabIndex        =   118
         Top             =   1845
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6C 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B0F
         DataField       =   "bien_codigo4"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6600
         TabIndex        =   119
         Top             =   2325
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6D 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B28
         DataField       =   "bien_codigo5"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   6600
         TabIndex        =   120
         Top             =   2805
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "bien_codigo"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6Z 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B41
         DataField       =   "bien_codigo1"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   1200
         TabIndex        =   121
         Top             =   885
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6A 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B5A
         DataField       =   "bien_codigo2"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   1200
         TabIndex        =   122
         Top             =   1365
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6B 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B73
         DataField       =   "bien_codigo3"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   1200
         TabIndex        =   123
         Top             =   1845
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6C 
         Bindings        =   "frm_ac_bienes_eqp.frx":1B8C
         DataField       =   "bien_codigo4"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   1200
         TabIndex        =   124
         Top             =   2325
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc6D 
         Bindings        =   "frm_ac_bienes_eqp.frx":1BA5
         DataField       =   "bien_codigo5"
         DataSource      =   "Ado_datos2"
         Height          =   315
         Left            =   1200
         TabIndex        =   125
         Top             =   2805
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "bien_descripcion"
         BoundColumn     =   "bien_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frm_ac_bienes_eqp.frx":1BBE
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1200
         TabIndex        =   134
         Top             =   480
         Width           =   8775
      End
      Begin VB.Label lbl_insumo5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   130
         Top             =   2820
         Width           =   780
      End
      Begin VB.Label lbl_insumo2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   129
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lbl_insumo4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   128
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label lbl_insumo1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   127
         Top             =   900
         Width           =   780
      End
      Begin VB.Label lbl_insumo3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Insumo 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   126
         Top             =   1860
         Width           =   780
      End
   End
   Begin VB.Frame FraBco 
      BackColor       =   &H00E0E0E0&
      Caption         =   "INSUMOS BASICOS PARA CRONOGRAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   850
      Left            =   6120
      TabIndex        =   99
      Top             =   5520
      Width           =   10740
      Begin VB.CommandButton BtnModificar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insumos"
         Height          =   560
         Left            =   10020
         Picture         =   "frm_ac_bienes_eqp.frx":1C50
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Editar INSUMOS del equipo"
         Top             =   240
         Width           =   700
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grabar"
         Height          =   560
         Left            =   10020
         Picture         =   "frm_ac_bienes_eqp.frx":2652
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Carga Foto de la Persona"
         Top             =   240
         Visible         =   0   'False
         Width           =   700
      End
      Begin MSDataGridLib.DataGrid dg_datos2 
         Bindings        =   "frm_ac_bienes_eqp.frx":3054
         Height          =   525
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   926
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
            DataField       =   "bien_codigo1"
            Caption         =   "1.Trapo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "cantidad1"
            Caption         =   "1.Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "bien_codigo2"
            Caption         =   "2.Gasolina"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "cantidad2"
            Caption         =   "2.Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "bien_codigo3"
            Caption         =   "3.Aceite.680"
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
            DataField       =   "cantidad3"
            Caption         =   "3.Cantidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "bien_codigo4"
            Caption         =   "4.Aceite.20/50"
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
            DataField       =   "cantidad4"
            Caption         =   "4.Cantidad"
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
            DataField       =   "bien_codigo5"
            Caption         =   "5.Grasa.Rodam"
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
            DataField       =   "cantidad5"
            Caption         =   "5.Cantidad"
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
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   86
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnVer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6840
         Picture         =   "frm_ac_bienes_eqp.frx":306D
         ScaleHeight     =   615
         ScaleWidth      =   1515
         TabIndex        =   97
         ToolTipText     =   "Inventario Fisico"
         Top             =   0
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   15480
         Picture         =   "frm_ac_bienes_eqp.frx":3C4A
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17880
         Picture         =   "frm_ac_bienes_eqp.frx":3E54
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   94
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
         Picture         =   "frm_ac_bienes_eqp.frx":4616
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   93
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "frm_ac_bienes_eqp.frx":4EE3
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   92
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8400
         Picture         =   "frm_ac_bienes_eqp.frx":5698
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   91
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2760
         Picture         =   "frm_ac_bienes_eqp.frx":5ECB
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   90
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "frm_ac_bienes_eqp.frx":6617
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   89
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "frm_ac_bienes_eqp.frx":6F2C
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   88
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox Imprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   16440
         Picture         =   "frm_ac_bienes_eqp.frx":76EB
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   87
         ToolTipText     =   "Inventario Fisico"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO"
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
         Left            =   13320
         TabIndex        =   96
         Top             =   195
         Width           =   885
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
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.CommandButton BtnImprimirA 
         Caption         =   "Inventario Valorado"
         Height          =   720
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "frm_ac_bienes_eqp.frx":7FB8
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   83
         Top             =   0
         Width           =   1455
      End
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4680
         Picture         =   "frm_ac_bienes_eqp.frx":88A4
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   82
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label lbl_titulo2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTITULO"
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
         Left            =   13005
         TabIndex        =   84
         Top             =   195
         Width           =   1425
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00C00000&
      Height          =   8070
      Left            =   60
      TabIndex        =   50
      Top             =   825
      Width           =   6015
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
         TabIndex        =   20
         Top             =   7665
         Visible         =   0   'False
         Width           =   915
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
         TabIndex        =   19
         Top             =   7665
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "frm_ac_bienes_eqp.frx":907A
         Height          =   7290
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   12859
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "bien_codigo"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            DataField       =   "bien_descripcion"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "grupo_codigo"
            Caption         =   "Grupo"
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
            DataField       =   "subgrupo_codigo"
            Caption         =   "SubGrupo"
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
            DataField       =   "par_codigo"
            Caption         =   "Partida"
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
            DataField       =   "marca_codigo"
            Caption         =   "Marca"
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
            DataField       =   "edif_codigo"
            Caption         =   "Edificio"
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
            DataField       =   "observaciones"
            Caption         =   "Nombre.de.Edificio"
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
               Alignment       =   2
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1950.236
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   7605
         Width           =   5865
         _ExtentX        =   10345
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
   Begin VB.Frame FraArticulos 
      BackColor       =   &H00C0C0C0&
      Height          =   7995
      Left            =   6120
      TabIndex        =   26
      Top             =   840
      Width           =   10725
      Begin VB.TextBox TxtDescripcionSIN 
         DataField       =   "descripcion_pSIN"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         ScrollBars      =   1  'Horizontal
         TabIndex        =   138
         Top             =   7200
         Width           =   5775
      End
      Begin VB.TextBox TxtCodigo_pSIN 
         DataField       =   "codigo_pSIN"
         DataSource      =   "Ado_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   136
         Top             =   7200
         Width           =   1695
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   2  'Center
         DataField       =   "bien_cantidad_por_empaque"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   9645
         MaxLength       =   10
         TabIndex        =   104
         Text            =   "0"
         Top             =   6360
         Width           =   765
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7875
         TabIndex        =   79
         Top             =   1860
         Width           =   280
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7875
         TabIndex        =   78
         Top             =   1170
         Width           =   280
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   7875
         TabIndex        =   77
         Top             =   490
         Width           =   280
      End
      Begin VB.TextBox TxtPrecEstD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_final_dol"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   76
         Text            =   "0.00"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox TxtPrecVentaD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_base_dol"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox TxtPrecCompD 
         Alignment       =   2  'Center
         DataField       =   "bien_precio_compra_dol"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox txtStockIni 
         Alignment       =   2  'Center
         DataField       =   "bien_stock_inicial"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   5220
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cmd_rotacion 
         DataField       =   "bien_rotacion"
         DataSource      =   "Ado_datos"
         Height          =   315
         ItemData        =   "frm_ac_bienes_eqp.frx":9092
         Left            =   9120
         List            =   "frm_ac_bienes_eqp.frx":909F
         TabIndex        =   12
         Text            =   "ALTA"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6360
         TabIndex        =   64
         Top             =   4335
         Width           =   290
      End
      Begin MSDataListLib.DataCombo dtc_partida 
         Bindings        =   "frm_ac_bienes_eqp.frx":90B9
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   57
         Top             =   1470
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "par_codigo"
         BoundColumn     =   "par_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "frm_ac_bienes_eqp.frx":90D2
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2445
         TabIndex        =   18
         Top             =   6720
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin VB.TextBox TxtPrecVenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_base"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox TxtDescripcion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   2445
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2475
         Width           =   8025
      End
      Begin VB.PictureBox Img_Foto 
         Height          =   2175
         Left            =   8265
         ScaleHeight     =   2115
         ScaleWidth      =   2235
         TabIndex        =   47
         Top             =   240
         Width           =   2295
         Begin VB.Image Image2 
            Height          =   2115
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2235
         End
      End
      Begin VB.TextBox TxtDescripcion2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_descripcion_anterior"
         DataSource      =   "Ado_datos"
         Height          =   405
         Left            =   1845
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2895
         Width           =   8625
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "Fecha_Vencimiento"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   5400
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   450
         _Version        =   393216
         Format          =   117571585
         CurrentDate     =   44993
      End
      Begin VB.TextBox TxtPrecEst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "bien_precio_venta_final"
         DataSource      =   "Ado_datos"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4920
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   5860
         Width           =   1215
      End
      Begin VB.TextBox txtStockMin 
         Alignment       =   2  'Center
         DataField       =   "bien_stock_minimo"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   160
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   5220
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox TxtPrecComp 
         Alignment       =   2  'Center
         DataField       =   "bien_precio_compra"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   5860
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_sub_cod 
         Bindings        =   "frm_ac_bienes_eqp.frx":90EC
         DataField       =   "subgrupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "subgrupo_codigo"
         BoundColumn     =   "subgrupo_codigo"
         Text            =   "Elige Marca..."
      End
      Begin VB.TextBox TxtInicial 
         Alignment       =   2  'Center
         DataField       =   "bien_codigo_anterior"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   7005
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "0"
         Top             =   4320
         Width           =   1605
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         DataField       =   "bien_codigo"
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   160
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "12345678901234567890"
         Top             =   2475
         Width           =   2295
      End
      Begin VB.CheckBox chkEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   8205
         TabIndex        =   23
         Top             =   5640
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo TDBC_marcas 
         Bindings        =   "frm_ac_bienes_eqp.frx":9106
         DataField       =   "marca_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6720
         TabIndex        =   9
         Top             =   3675
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "marca_descripcion"
         BoundColumn     =   "marca_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo marcas 
         Bindings        =   "frm_ac_bienes_eqp.frx":911D
         DataField       =   "marca_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9240
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "marca_codigo"
         BoundColumn     =   "marca_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_sub_des 
         Bindings        =   "frm_ac_bienes_eqp.frx":9134
         DataField       =   "subgrupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   1160
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   741
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   8388608
         ListField       =   "subgrupo_descripcion"
         BoundColumn     =   "subgrupo_codigo"
         Text            =   "Elige Marca..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo TDBC_Unidad 
         Bindings        =   "frm_ac_bienes_eqp.frx":914E
         DataField       =   "unimed_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   165
         TabIndex        =   7
         Top             =   3675
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "unimed_descripcion"
         BoundColumn     =   "unimed_codigo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo Unidad 
         Bindings        =   "frm_ac_bienes_eqp.frx":9166
         DataField       =   "unimed_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1845
         TabIndex        =   35
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "unimed_codigo"
         BoundColumn     =   "unimed_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoCod 
         Bindings        =   "frm_ac_bienes_eqp.frx":917E
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7080
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "grupo_codigo"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoDes 
         Bindings        =   "frm_ac_bienes_eqp.frx":9195
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1515
         TabIndex        =   1
         Top             =   480
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   741
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   8388608
         ListField       =   "grupo_descripcion"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "Fecha_Alerta"
         Height          =   255
         Left            =   9240
         TabIndex        =   25
         Top             =   5640
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   450
         _Version        =   393216
         Format          =   117571585
         CurrentDate     =   44993
      End
      Begin MSDataListLib.DataCombo DtcPaisD 
         Bindings        =   "frm_ac_bienes_eqp.frx":91AC
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3015
         TabIndex        =   8
         Top             =   3675
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "pais_descripcion"
         BoundColumn     =   "pais_codigo"
         Text            =   "Elige Medida ..."
      End
      Begin MSDataListLib.DataCombo DtcPais 
         Bindings        =   "frm_ac_bienes_eqp.frx":91C2
         DataField       =   "pais_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5280
         TabIndex        =   45
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "pais_codigo"
         BoundColumn     =   "pais_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo DtcGrupoUni 
         Bindings        =   "frm_ac_bienes_eqp.frx":91D8
         DataField       =   "grupo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6000
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "unidad_codigo"
         BoundColumn     =   "grupo_codigo"
         Text            =   "Elige Grupo ..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "frm_ac_bienes_eqp.frx":91EF
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   885
         TabIndex        =   56
         Top             =   6720
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_partida_des 
         Bindings        =   "frm_ac_bienes_eqp.frx":9209
         DataField       =   "par_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         Top             =   1840
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   741
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   8388608
         ListField       =   "par_descripcion"
         BoundColumn     =   "par_codigo"
         Text            =   "Elige Marca..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "frm_ac_bienes_eqp.frx":9222
         DataField       =   "modelo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   2565
         TabIndex        =   59
         Top             =   4320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ForeColor       =   12582912
         ListField       =   "modelo_descripcion"
         BoundColumn     =   "modelo_codigo"
         Text            =   "Elige Marca ..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "frm_ac_bienes_eqp.frx":923B
         DataField       =   "modelo_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   165
         TabIndex        =   10
         Top             =   4320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "modelo_codigo"
         BoundColumn     =   "modelo_codigo"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "frm_ac_bienes_eqp.frx":9254
         DataField       =   "bien_codigo_universal"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9000
         TabIndex        =   66
         Top             =   6120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "tipo_eqp"
         BoundColumn     =   "tipo_eqp"
         Text            =   "Elige Marca..."
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "frm_ac_bienes_eqp.frx":926D
         DataField       =   "bien_codigo_universal"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6600
         TabIndex        =   17
         Top             =   5880
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "tipo_eqp_descripcion"
         BoundColumn     =   "tipo_eqp"
         Text            =   "Elige Tipo Equipo ..."
      End
      Begin MSDataListLib.DataCombo Dtc_descripcionSIN 
         Bindings        =   "frm_ac_bienes_eqp.frx":9286
         DataField       =   "correlativo_pSIN"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   140
         Top             =   7560
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "correlativo_pSIN"
         Text            =   "Elige producto de SIN..."
      End
      Begin MSDataListLib.DataCombo Dtc_codigoSIN 
         Bindings        =   "frm_ac_bienes_eqp.frx":92A3
         DataField       =   "correlativo_pSIN"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9960
         TabIndex        =   141
         Top             =   7560
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ListField       =   "correlativo_pSIN"
         BoundColumn     =   "correlativo_pSIN"
         Text            =   "Elige Marca..."
      End
      Begin VB.Label LblProductoSIN 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto SIN"
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label LblDescripcionSIN 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion del Producto para SIN"
         Height          =   375
         Left            =   3240
         TabIndex        =   137
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label LblCodigoPSin 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo para SIN"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas de Servicio (Cronograma):"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6600
         TabIndex        =   103
         Top             =   6360
         Width           =   2955
      End
      Begin VB.Label txtCantVendida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_salida"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6720
         TabIndex        =   80
         Top             =   5220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4440
         TabIndex        =   74
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2400
         TabIndex        =   73
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4440
         TabIndex        =   72
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2400
         TabIndex        =   71
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   70
         Top             =   6240
         Width           =   435
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bs."
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   150
         TabIndex        =   69
         Top             =   5880
         Width           =   285
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Iinicial"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   68
         Top             =   4980
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rotación"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9000
         TabIndex        =   67
         Top             =   4080
         Width           =   1125
      End
      Begin VB.Label lbl_eqp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Equipo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6600
         TabIndex        =   65
         Top             =   5625
         Width           =   1395
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB GRUPO 2"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   63
         Top             =   1580
         Width           =   1350
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB GRUPO 1"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   62
         Top             =   900
         Width           =   1350
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   160
         TabIndex        =   61
         Top             =   200
         Width           =   735
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   60
         Top             =   4080
         Width           =   690
      End
      Begin VB.Label txt_par 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "par_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   165
         TabIndex        =   58
         Top             =   1840
         Width           =   1320
      End
      Begin VB.Label lbl_edif 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   55
         Top             =   6720
         Width           =   660
      End
      Begin VB.Label TxtActual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_actual"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8880
         TabIndex        =   54
         Top             =   5220
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label txtCantComprada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "bien_stock_ingreso"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   53
         Top             =   6300
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label TxtSub 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "subgrupo_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   165
         TabIndex        =   52
         Top             =   1160
         Width           =   1320
      End
      Begin VB.Label TxtGrupo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "grupo_codigo"
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   160
         TabIndex        =   51
         Top             =   480
         Width           =   1320
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Mínimo"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   165
         TabIndex        =   49
         Top             =   4980
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Características Complementarias"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   165
         TabIndex        =   46
         Top             =   2835
         Width           =   1740
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Industria (Pais Origen)"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3015
         TabIndex        =   44
         Top             =   3420
         Width           =   1965
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Total Vendida"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6525
         TabIndex        =   43
         Top             =   4980
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Total Comprada"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4155
         TabIndex        =   42
         Top             =   4980
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Venta.Cliente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   5595
         Width           =   1845
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Referencia"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7005
         TabIndex        =   39
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label TDBFrame3D6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Actual"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8880
         TabIndex        =   38
         Top             =   4980
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label TDBFrame3D7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Compra.FOB"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   37
         Top             =   5595
         Width           =   1815
      End
      Begin VB.Label TDBFrame3D8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio.Venta.Base"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   5595
         Width           =   1725
      End
      Begin VB.Label TDBFrame3D5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad de Medida"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   32
         Top             =   3420
         Width           =   1695
      End
      Begin VB.Label TDBFrame3D9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marca"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6765
         TabIndex        =   31
         Top             =   3420
         Width           =   570
      End
      Begin VB.Label TDBFrame3D1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "  CODIGO EQUIPO                                DESCRIPCION DEL EQUIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   165
         TabIndex        =   30
         Top             =   2235
         Width           =   8010
      End
   End
   Begin MSAdodcLib.Adodc AdoSubGrupo 
      Height          =   375
      Left            =   6960
      Top             =   9360
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   661
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
      Caption         =   "AdoSubGrp"
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
   Begin MSAdodcLib.Adodc AdoMedida 
      Height          =   375
      Left            =   9360
      Top             =   9360
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   661
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
      Caption         =   "medida"
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
   Begin MSAdodcLib.Adodc AdoMarca 
      Height          =   375
      Left            =   11760
      Top             =   9360
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "marca"
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
   Begin VB.PictureBox picFondo 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11145
      TabIndex        =   27
      Top             =   7860
      Width           =   11145
      Begin VB.Frame Frame4 
         Height          =   60
         Left            =   15
         TabIndex        =   28
         Top             =   255
         Width           =   12570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   2
         Left            =   12840
         TabIndex        =   29
         Top             =   75
         Width           =   1845
      End
   End
   Begin VB.PictureBox imlMaterial 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   4200
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   98
      Top             =   4200
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc AdoPais 
      Height          =   375
      Left            =   4680
      Top             =   9360
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "AdoPais"
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
   Begin MSAdodcLib.Adodc AdoGrupo 
      Height          =   375
      Left            =   2400
      Top             =   9360
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   661
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
      Caption         =   "AdoGrupo"
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
   Begin Crystal.CrystalReport CryLista 
      Left            =   120
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryBBSS 
      Left            =   600
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CryFis 
      Left            =   1080
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Ado_datos10 
      Height          =   330
      Left            =   120
      Top             =   9360
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   120
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
      Left            =   2400
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
      Left            =   4680
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   6960
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
      Left            =   9240
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
   Begin MSAdodcLib.Adodc AdoProductoSin 
      Height          =   330
      Left            =   11640
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "AdoProductoSin"
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
Attribute VB_Name = "frm_ac_bienes_eqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim rsMarcas As ADODB.Recordset
Dim rsUnidad As ADODB.Recordset
Dim rsSubGrupo As ADODB.Recordset
Dim rsProductosSIN As ADODB.Recordset

Dim rsgrupo As ADODB.Recordset
Dim RsArt, rsPais As ADODB.Recordset
Dim rsNada As ADODB.Recordset
Dim rs_datos2 As ADODB.Recordset
Dim rs_datos10 As ADODB.Recordset
Dim rs_datos6 As ADODB.Recordset
Dim rs_datos7 As ADODB.Recordset
Dim rs_datos8 As ADODB.Recordset
Dim rs_aux1, rs_aux2, rs_aux3 As ADODB.Recordset
Dim rs_aux4, rs_aux6, rs_aux7 As ADODB.Recordset
'--------
Dim estado, VAR_CONT As Integer ' 0 navegar, 1 Agregar, 2 Editar
Dim var_cod As Integer

Dim swnuevo As Boolean

Dim sino As String
Dim NombreCarpeta, e As String
Dim CodBien, COD_EDIF, COD_MOD As String
Dim VAR_OA, VAR_NEW As String
Dim VAR_SW2 As String
Dim marca1 As BookmarkEnum

Dim VAR_Dol As Double
Dim C_FIJO As Double
Dim C_MANOBR As Double
Dim C_GTOADM As Double
Dim C_UTILID As Double
Dim C_ROTALT As Double
Dim C_ROTBAJ As Double
Dim C_IMPSTO As Double
Dim C_IMPSTO2 As Double
'--
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim PosibleApliqueFiltro As Boolean
'Dim queryinicial As String

Public Sub ALPrincipal(QEstado As Integer)
    '
'    Screen.MousePointer = vbHourglass
'    estado = QEstado
'    '
'    Select Case estado
'        Case 0
'            Set RsArt = New ADODB.Recordset
'            'JQA 04/2008
'            'GlSqlAux = "SELECT * FROM ac_bienes WHERE CAST(grupo_codigo AS INT)< 50  AND bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY CAST (grupo_codigo AS INT)"
'            'GlSqlAux = "SELECT * FROM ac_bienes WHERE bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY grupo_codigo, subgrupo_codigo, bien_codigo "
'            queryinicial = "SELECT * FROM ac_bienes WHERE bien_codigo = ISNULL(bien_codigo, NULL) ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion "
'            RsArt.Open queryinicial, db, adOpenDynamic, adLockOptimistic
'            If RsArt.RecordCount > 0 Then
'               GlHayRegs = True  'Variable global
'            Else
'               GlHayRegs = False
'            End If
'            BotonesNavegar Me
'            FraArticulos.Enabled = False
'            Set Ado_datos.Recordset = RsArt
'        Case 1
'
'        Case 2
'
'    End Select
'    '
'    Screen.MousePointer = vbDefault
'    Me.Show
End Sub

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'Dim Marca As String
'Dim a As Integer
'Dim COD_MARCAx, cod_UMedida As String
If Ado_datos.Recordset.BOF Or Ado_datos.Recordset.EOF Then
        If Ado_datos.Recordset.BOF And Ado_datos.Recordset.EOF Then
            TxtGrupo.Caption = ""
            TxtDetalle.Text = ""
            txtDescripcion.Text = ""
            TxtActual.Caption = ""
            chkEstado.Value = vbUnchecked
'            Ado_datos.Caption = "Registro: 0 de 0"
'            BuscaNodo "Grupo"
        Else
            Exit Sub
        End If
Else
        If Ado_datos.Recordset!bien_stock_minimo < Ado_datos.Recordset!bien_stock_actual Then
            TxtActual.backColor = &HE0E0E0
        Else
            TxtActual.backColor = &HFF&
        End If

        'dtc_sub_des
    Set Img_Foto = Leer_Imagen(db, "Select Foto From ac_bienes_foto Where bien_codigo = '" & Ado_datos.Recordset("bien_codigo") & "' ", "Foto")
    Image2 = Img_Foto
    If Ado_datos.Recordset!estado_codigo = "APR" Then
        'chkEstado.Value = vbChecked
        BtnVer.Visible = True
    Else
        BtnVer.Visible = False
        'chkEstado.Value = vbUnchecked
    End If
    If Ado_datos.Recordset!subgrupo_codigo = "43000" Then
        dtc_codigo10.Visible = True
        dtc_desc10.Visible = True
        lbl_edif.Visible = True
        'dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        lbl_eqp.Visible = True
    Else
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        lbl_edif.Visible = False
        'dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        lbl_eqp.Visible = False
    End If
        'chkEstado.Value =IIf(CBool(Ado_datos.Recordset!estado), vbChecked, vbUnchecked)
'        BuscaNodo Ado_datos.Recordset!grupo_codigo
    'INSUMOS POR CADA EQUIPO
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "SELECT * FROM ac_bienes_insumos_kit where bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
    'rs_aux4.Sort = "grupo_codigo, subgrupo_codigo, bien_codigo"
    If rs_aux4.RecordCount > 0 Then
       dg_datos2.Visible = True
       Set Ado_datos2.Recordset = rs_aux4
       Set dg_datos2.DataSource = Ado_datos2.Recordset
    Else
       dg_datos2.Visible = False
    End If
    
    
End If
End Sub

Private Sub BtnAñadir_Click()
    If glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Or glusuario = "JMAMANI" Or glusuario = "VMEJIA" Then
        swnuevo = True
    '    Set dg_datos.DataSource = rsNada
        Ado_datos.Recordset.AddNew
        estado = 1
    '    BotonesEditar Me
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        FraArticulos.Enabled = True
        TxtDetalle.Enabled = True
    '    TxtGrupo.Enabled = False
    '    DtcGrupoDes.Enabled = True
    '    TxtSub.Enabled = False
    '    dtc_sub_des.Enabled = False
    '    trv.SetFocus
    '    BuscaNodo "grupo"
        txtStockMin.Text = 0
    Else
        MsgBox "El usuario NO tiene acceso, consulte con el Administrador del Sistema ...", vbExclamation, "Validación de Registro"
    End If
End Sub

Private Sub BtnAprobar_Click()
   If Valida Then
       sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
       If Ado_datos.Recordset("estado_codigo") = "REG" Then
          If sino = vbYes Then
            CodBien = Ado_datos.Recordset!bien_codigo
            COD_EDIF = Ado_datos.Recordset!edif_codigo
            COD_MOD = Ado_datos.Recordset!modelo_codigo
'            If Ado_datos.Recordset!grupo_codigo = "40000" Then
'                Call ACTUALIZA_ID
'                Call ACTUALIZA_VENTA
'            End If
    '        Dim RUTA1, RUTA2 As String
    '        RUTA1 = "BIENES" + "\" + Trim(adoLista.Recordset("iniciales")) + "-" + Trim(adoLista.Recordset("codigo_beneficiario"))
    '        MsgBox RUTA1
    '        MkDir RUTA1
    '        MkDir RUTA1 + "\CONTRATOS"
    '        MkDir RUTA1 + "\FINIQUITO"
    '        MkDir RUTA1 + "\MEMORANDUMS"
    '        MkDir RUTA1 + "\RESPALDOS"
    '        MkDir RUTA1 + "\HOJA_VIDA"
    '        MkDir RUTA1 + "\OTROS"
    '        MkDir RUTA1 + "\EVALUACIONES"
    '        MkDir RUTA1 + "\LICENCIAS"
    '        MkDir RUTA1 + "\VACACIONES"
            'Ado_datos.Recordset("estado") = 1
            db.Execute "update ac_bienes set estado_codigo = 'APR' WHERE bien_codigo = '" & CodBien & "'  "
'            Ado_datos.Recordset("estado_codigo") = "APR"
'            Ado_datos.Recordset("fecha_registro") = Date
'            Ado_datos.Recordset("usr_codigo") = glusuario
'            Ado_datos.Recordset.Update
            
'            If OptFilGral1.Value = True Then
'               Call OptFilGral1_Click        'Pendientes
'            Else
'               Call OptFilGral2_Click        'TODOS
'            End If
            Call OptFilGral2_Click
            If (dg_datos.SelBookmarks.Count <> 0) Then
               dg_datos.SelBookmarks.Remove 0
            End If
            If Ado_datos.Recordset.RecordCount > 0 Then
               RsArt.Find "bien_codigo = '" & CodBien & "'   ", , , 1
               dg_datos.SelBookmarks.Add (RsArt.Bookmark)
            Else
               RsArt.MoveLast
            End If

          End If
       Else
            MsgBox "No se puede APROBAR un registro Anulado o Aprobado anteriormente ...", vbExclamation, "Validación de Registro"
       End If
   Else
        MsgBox "Existe un error en los datos registrados, Verifique y vuelva a intentar...", vbExclamation + vbOKOnly, "Atención"
   End If
End Sub

Private Sub ACTUALIZA_ID()
    'wwwwwwwwwwwwwwwwwwwwwwwwwwww
    'ACTUALIZA EQUIPOS
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "select * from ao_solicitud where edif_codigo = '" & COD_EDIF & "'   ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux1.RecordCount > 0 Then
        Set rs_aux3 = New ADODB.Recordset
        If rs_aux3.State = 1 Then rs_aux3.Close
        'Id. CLIENTE "36NO" EXISTENTE
        rs_aux3.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND bien_codigo = '" & CodBien & "' ", db, adOpenStatic
        If rs_aux1.RecordCount > 0 Then
            db.Execute "update ao_solicitud_bienes set modelo_codigo = '" & COD_MOD & "' WHERE bien_codigo = '" & CodBien & "' AND unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & " "
        Else
            VAR_CONT = 1
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            'Id. CLIENTE "36NO" NUEVO
            rs_aux2.Open "Select * from ao_solicitud_bienes where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND grupo_codigo = '90000' ", db, adOpenStatic
            db.Execute "INSERT INTO ao_solicitud_bienes (ges_gestion, unidad_codigo, solicitud_codigo, bien_codigo, grupo_codigo, subgrupo_codigo, par_codigo, marca_codigo, modelo_codigo, bien_cantidad, bien_precio_compra, bien_total_compra, bien_precio_venta_base, bien_total_venta, tipo_moneda, unimed_codigo, unimed_codigo_empaque, bien_cantidad_por_empaque, venta_o_compra, fosa_dimension_frente, fosa_dimension_fondo, estado_codigo, usr_codigo, fecha_registro ) VALUES ('" & glGestion & "', '" & rs_aux1!unidad_codigo & "',  " & rs_aux1!solicitud_codigo & ", '" & CodBien & "', '40000', '43000', '43340', '" & Ado_datos.Recordset!marca_codigo & "', '" & COD_MOD & "', " & rs_aux2!bien_cantidad & ", 0, 0, " & rs_aux2!bien_precio_venta_base & ", " & rs_aux2!bien_total_venta & ", 'BOB', '" & rs_aux2!unimed_codigo & "', '" & rs_aux2!unimed_codigo & "', " & rs_aux2!bien_cantidad & ", 'V', 0, 0, 'APR', '" & glusuario & "', '" & Date & "')"
    
            If rs_aux2!bien_codigo = "NA1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              If rs_aux2.RecordCount > 1 Then
                  db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                  If rs_aux2.RecordCount > 2 Then
                      db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                      If rs_aux2.RecordCount > 3 Then
                          db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA4' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                          If rs_aux2.RecordCount > 4 Then
                              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA5' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                              If rs_aux2.RecordCount > 5 Then
                                  db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA6' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                                  If rs_aux2.RecordCount > 6 Then
                                      db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NA7' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
                                  End If
                              End If
                          End If
                      End If
                  End If
              End If
            End If
            If rs_aux2!bien_codigo = "NE1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NE3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
            End If
            If rs_aux2!bien_codigo = "NP1" Then
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP1' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP2' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
              db.Execute "update ao_solicitud_bienes set estado_codigo = 'ANL' WHERE bien_codigo = 'NP3' AND unidad_codigo = '" & rs_aux2!unidad_codigo & "' and solicitud_codigo = " & rs_aux2!solicitud_codigo & " "
            End If
        End If
    End If
    'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub ACTUALIZA_VENTA()
   'VENTAS
   Set rs_aux1 = New ADODB.Recordset
   If rs_aux1.State = 1 Then rs_aux1.Close
   rs_aux1.Open "select * from ao_ventas_cabecera where edif_codigo = '" & COD_EDIF & "'   ", db, adOpenKeyset, adLockBatchOptimistic
   If rs_aux1.RecordCount > 0 Then
      Set rs_aux3 = New ADODB.Recordset
      If rs_aux3.State = 1 Then rs_aux3.Close
        'Id. detalle "36NO" EXISTENTE
      rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & rs_aux1!unidad_codigo & "' and solicitud_codigo = " & rs_aux1!solicitud_codigo & "  AND bien_codigo = '" & CodBien & "' ", db, adOpenStatic
      If rs_aux1.RecordCount > 0 Then
            db.Execute "update ao_ventas_detalle set modelo_codigo = '" & COD_MOD & "' WHERE bien_codigo = '" & CodBien & "' AND venta_codigo = " & rs_aux1!venta_codigo & "   "
      Else
       'VENTAS DETALLE
       Set rs_aux2 = New ADODB.Recordset
       If rs_aux2.State = 1 Then rs_aux2.Close
       rs_aux2.Open "Select * from ao_ventas_detalle where venta_codigo = " & rs_aux1!venta_codigo & "  AND grupo_codigo = '90000' ", db, adOpenStatic
       If rs_aux2.RecordCount > 0 Then
           VAR_CONT = rs_aux2.RecordCount + 1
           If Ado_datos.Recordset!bien_precio_venta_final > 0 Then
                VAR_Dol = Round(Ado_datos.Recordset!bien_precio_venta_final / GlTipoCambioOficial, 2)
           Else
                VAR_Dol = 0
           End If
           'VENTAS
           db.Execute "INSERT INTO ao_ventas_detalle (ges_gestion, venta_codigo, bien_codigo, venta_codigo_det, venta_det_cantidad, venta_precio_unitario_bs, venta_descuento_bs, venta_precio_total_bs, venta_precio_unitario_dol, venta_descuento_dol, venta_precio_total_dol, concepto_venta, grupo_codigo, subgrupo_codigo, par_codigo, tipo_descuento, almacen_codigo, modelo_codigo, modelo_codigo1, modelo_codigo_h, modelo_codigo_x, modelo_elegido,   estado_codigo, usr_codigo, fecha_registro) " & _
           " VALUES ('" & glGestion & "', " & rs_aux1!venta_codigo & ", '" & CodBien & "', " & VAR_CONT & ", " & rs_aux2!venta_det_cantidad & ", " & rs_aux2!venta_precio_unitario_bs & ", 0, " & rs_aux2!venta_precio_total_bs & ", " & Round(rs_aux2!venta_precio_unitario_dol, 2) & ", 0, " & Round(rs_aux2!venta_precio_total_dol, 2) & ", '" & Ado_datos.Recordset!bien_descripcion & "', '40000', '43000', '43340', 0, 0, '" & COD_MOD & "', '" & COD_MOD & "', 'S/M', 'S/M', 'S',  'APR', '" & glusuario & "', '" & Date & "') "

        If rs_aux2!bien_codigo = "NA1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          If rs_aux2.RecordCount > 1 Then
              db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
              If rs_aux2.RecordCount > 2 Then
                  db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                  If rs_aux2.RecordCount > 3 Then
                      db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA4' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                      If rs_aux2.RecordCount > 4 Then
                          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA5' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                          If rs_aux2.RecordCount > 5 Then
                              db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA6' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                              If rs_aux2.RecordCount > 6 Then
                                  db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NA7' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
                              End If
                          End If
                      End If
                  End If
              End If
          End If
        End If
        If rs_aux2!bien_codigo = "NE1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NE3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
        End If
        If rs_aux2!bien_codigo = "NP1" Then
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP1' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP2' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
          db.Execute "update ao_ventas_detalle set estado_codigo = 'ANL' WHERE bien_codigo = 'NP3' AND venta_codigo = " & rs_aux2!venta_codigo & "  "
        End If
      End If
     End If
   End If
End Sub

Private Sub BtnBuscar_Click()
'  Set ClBuscaGrid = New ClBuscaEnGridExterno
'  Set ClBuscaGrid.Conexión = db
'  ClBuscaGrid.QueryUtilizado = GlSqlAux
'  ClBuscaGrid.Título = "Elija un Detalle"
'  ClBuscaGrid.EsTdbGrid = True
'  Set ClBuscaGrid.GridTrabajo = dg_datos
'  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
'  ClBuscaGrid.Ejecutar
''  If ClBuscaGrid.ElegidoCol1 <> "" Then
''    Ado_datos.Recordset.Filter = adFilterNone
''    Ado_datos.Recordset.MoveFirst
''    Ado_datos.Recordset.Find "grupo_codigo + '-' + bien_codigo   = " & ClBuscaGrid.ElegidoCol1 & " - " & ClBuscaGrid.ElegidoCol2 & ""
'  End If

    buscados = 1
'    OptFilGral2.Visible = False
'    OptFilGral1.Visible = False
    Call OptFilGral2_Click
  PosibleApliqueFiltro = False
  Set ClBuscaGrid = New ClBuscaEnGridExterno
  Set ClBuscaGrid.Conexión = db
  ClBuscaGrid.EsTdbGrid = False
  Set ClBuscaGrid.GridTrabajo = dg_datos
  ClBuscaGrid.QueryUtilizado = queryinicial
  Set ClBuscaGrid.RecordsetTrabajo = Ado_datos.Recordset
  ClBuscaGrid.CamposVisibles = "110"
  ClBuscaGrid.Ejecutar
  PosibleApliqueFiltro = True

End Sub

Private Sub BtnCancelar_Click()

On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    If swnuevo = False Then
        CodBien = Ado_datos.Recordset!bien_codigo
        If Ado_datos.Recordset.EditMode <> adEditNone Then Ado_datos.Recordset.CancelUpdate
    Else
        CodBien = "0"
    End If
    
    Ado_datos.Recordset.Cancel
    
    Call OptFilGral2_Click
    If (dg_datos.SelBookmarks.Count <> 0) Then
       dg_datos.SelBookmarks.Remove 0
    End If
    If Ado_datos.Recordset.RecordCount > 0 And CodBien <> "0" Then
       RsArt.Find "bien_codigo = '" & CodBien & "'   ", , , 1
       dg_datos.SelBookmarks.Add (RsArt.Bookmark)
    Else
       RsArt.MoveLast
    End If
'    Call OptFilGral2_Click
'    Call CARGA
'    Ado_datos.Caption = "Registro: " & CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & Ado_datos.Recordset.RecordCount
    'BotonesNavegar Me
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraArticulos.Enabled = False
'    TxtGrupo.Enabled = True
'    DtcGrupoDes.Enabled = True
'    TxtSub.Enabled = True
'    dtc_sub_des.Enabled = True
'    Set dg_datos.DataSource = Ado_datos
    Screen.MousePointer = vbDefault
    estado = 0
    swnuevo = False
    dg_datos.Enabled = True
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnGrabar2_Click()
    If glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Then
        dg_datos2.AllowUpdate = False
        dg_datos2.Enabled = False
    Else
        dg_datos2.AllowUpdate = False
        MsgBox "El Ususario NO tiene acceso, consulte con el Administrador del Sistema ... ", vbExclamation, "Validación de Registro"
    End If
    BtnGrabar2.Visible = False
    BtnModificar2.Visible = True
End Sub

Private Sub BtnModificar_Click()
On Error GoTo Que_Error
    If Ado_datos.Recordset!estado_codigo = "REG" Or (glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Or glusuario = "JMAMANI" Or glusuario = "VMEJIA") Then
        If Ado_datos.Recordset!estado_codigo = "REG" Then
            TxtDetalle.Locked = True
        Else
            TxtDetalle.Locked = False
        End If
        swnuevo = False
        Screen.MousePointer = vbHourglass
        'BotonesEditar Me
        estado = 2
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        FraArticulos.Enabled = True
'        Ado_datos.Caption = "Editando Registro..."
        Screen.MousePointer = vbDefault
        dg_datos.Enabled = False
        If Ado_datos.Recordset!subgrupo_codigo = "43000" Then
            dtc_codigo10.Visible = True
            dtc_desc10.Visible = True
            lbl_edif.Visible = True
    '        dtc_codigo8.Visible = True
            dtc_desc8.Visible = True
            lbl_eqp.Visible = True
        Else
            dtc_codigo10.Visible = False
            dtc_desc10.Visible = False
            lbl_edif.Visible = False
    '        dtc_codigo8.Visible = False
            dtc_desc8.Visible = False
            lbl_eqp.Visible = False
        End If
    Else
        MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
    End If
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub btnEliminar_Click()
On Error GoTo Que_Error
    'ao_adjudica_detalle_D
    If glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Then
       If Ado_datos.Recordset.RecordCount > 0 Then
          If ExisteDetalle(Ado_datos.Recordset!bien_codigo) Then MsgBox "No se puede eliminar un BIEN o SERVICIO que ya tiene Registros en COMPRAS o ALMACEN.", vbInformation + vbOKOnly, "Atención": Exit Sub
          sino = MsgBox("Está Seguro de ANULAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
          If sino = vbYes Then
             'Ado_datos.Recordset.Delete
             Ado_datos.Recordset!estado_codigo = "ANL"
             Ado_datos.Recordset.Update
             Ado_datos.Recordset.Requery
          End If
       Else
            MsgBox "No existen registros para Anular.", vbExclamation, "Atención"
       End If
    Else
       MsgBox "No se puede MODIFICAR un registro Aprobado (APR) o Errado (ERR) ...", vbExclamation, "Validación de Registro"
    End If
   Exit Sub
    
'    If Not GlHayRegs Then
'        MsgBox "No existen registro para Anular", vbExclamation + vbOKOnly, "Atención"
'        Exit Sub
'    End If
'    If ExisteDetalle(Ado_datos.Recordset!grupo_codigo & "-" & Ado_datos.Recordset!bien_codigo) Then MsgBox "No se puede eliminar el Detalle seleccionado ya que se tiene registro de Movimientos en Almacen.", vbInformation + vbOKOnly, "Atención": Exit Sub
'    If MsgBox("¿ Está seguro que se va a Anular el registro visualizado ?", vbExclamation + vbOKCancel, "Atención") = vbOK Then
'        Screen.MousePointer = vbHourglass
'        'Ado_datos.Recordset.Delete
'        Ado_datos.Recordset!estado = 2
'        Ado_datos.Recordset.MoveNext
'        If Ado_datos.Recordset.EOF Then
'          If Ado_datos.Recordset.RecordCount > 0 Then
'            Ado_datos.Recordset.MoveLast
'          Else
'            GlHayRegs = False
'            Ado_datos.Refresh
'          End If
'        End If
'        Screen.MousePointer = vbDefault
'    End If
'    BotonesNavegar Me
'    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnModificar2_Click()
    If glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "ADMIN" Or glusuario = "JSAAVEDRA" Or glusuario = "GSOLIZ" Or glusuario = "CSALINAS" Or glusuario = "VMEJIA" Or glusuario = "JMAMANI" Then
        'dg_datos2.Enabled = True
        'dg_datos2.AllowUpdate = True
        '45. ADD. Kit de INSUMOS al modificar bienes...
        '
        Set rs_datos2 = New ADODB.Recordset
            If rs_datos2.State = 1 Then rs_datos2.Close
            rs_datos2.Open "SELECT * FROM ac_bienes where KIT = 'I01' ", db, adOpenKeyset, adLockOptimistic
            Set Ado_datos3.Recordset = rs_datos2
            
        If rs_aux4.RecordCount = 0 Then
            VAR_SW2 = "ADD"
            rs_aux4.AddNew
            FraInsumo.Visible = True
            FraInsumo.Enabled = True
            Txt_cant1.Text = IIf(Txt_cant1.Text = "", "0.2", Txt_cant1.Text)
            Txt_cant2.Text = IIf(Txt_cant2.Text = "", "0.25", Txt_cant2.Text)
            Txt_cant3.Text = IIf(Txt_cant3.Text = "", "0", Txt_cant3.Text)
            Txt_cant4.Text = IIf(Txt_cant4.Text = "", "0", Txt_cant4.Text)
            Txt_cant5.Text = IIf(Txt_cant5.Text = "", "0", Txt_cant5.Text)
            If dtc_codigo6Z.Text = "" Or dtc_codigo6Z.Text <> "4211" Then
                dtc_codigo6Z.Text = "4211"
                dtc_desc6Z.BoundText = dtc_codigo6Z.BoundText
            End If
            If dtc_codigo6A.Text = "" Or dtc_codigo6A.Text <> "479" Then
                dtc_codigo6A.Text = "479"
                dtc_desc6A.BoundText = dtc_codigo6A.BoundText
            End If
            If dtc_codigo6B.Text = "" Or dtc_codigo6B.Text <> "500" Then
                dtc_codigo6B.Text = "500"
                dtc_desc6B.BoundText = dtc_codigo6B.BoundText
            End If
            If dtc_codigo6C.Text = "" Or dtc_codigo6C.Text <> "4529" Then
                dtc_codigo6C.Text = "4529"
                dtc_desc6C.BoundText = dtc_codigo6C.BoundText
            End If
            If dtc_codigo6D.Text = "" Then
                dtc_codigo6D.Text = "3113"
                dtc_desc6D.BoundText = dtc_codigo6D.BoundText
            End If
        Else
            VAR_SW2 = "MOD"
            FraInsumo.Visible = True
            FraInsumo.Enabled = True
            Txt_cant1.Text = IIf(Txt_cant1.Text = "", "0.2", Txt_cant1.Text)
            Txt_cant2.Text = IIf(Txt_cant2.Text = "", "0.25", Txt_cant2.Text)
            Txt_cant3.Text = IIf(Txt_cant3.Text = "", "0", Txt_cant3.Text)
            Txt_cant4.Text = IIf(Txt_cant4.Text = "", "0", Txt_cant4.Text)
            Txt_cant5.Text = IIf(Txt_cant5.Text = "", "0", Txt_cant5.Text)
            If dtc_codigo6Z.Text = "" Or dtc_codigo6Z.Text <> "4211" Then
                dtc_codigo6Z.Text = "4211"
                dtc_desc6Z.BoundText = dtc_codigo6Z.BoundText
            End If
            If dtc_codigo6A.Text = "" Or dtc_codigo6A.Text <> "479" Then
                dtc_codigo6A.Text = "479"
                dtc_desc6A.BoundText = dtc_codigo6A.BoundText
            End If
            If dtc_codigo6B.Text = "" Or dtc_codigo6B.Text <> "500" Then
                dtc_codigo6B.Text = "500"
                dtc_desc6B.BoundText = dtc_codigo6B.BoundText
            End If
            If dtc_codigo6C.Text = "" Or dtc_codigo6C.Text <> "4529" Then
                dtc_codigo6C.Text = "4529"
                dtc_desc6C.BoundText = dtc_codigo6C.BoundText
            End If
            If dtc_codigo6D.Text = "" Then
                dtc_codigo6D.Text = "3113"
                dtc_desc6D.BoundText = dtc_codigo6D.BoundText
            End If
        End If
    Else
        dg_datos2.AllowUpdate = False
        MsgBox "El Ususario NO tiene permiso ... ", vbExclamation, "Validación de Registro"
    End If
    'BtnGrabar2.Visible = True
    'BtnModificar2.Visible = False
    'BtnModificar2.Enabled = False
End Sub

Private Sub BtnVer_Click()
  On Error GoTo QError
    If Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo" Then
      NombreCarpeta = App.Path & "\BIENES\" & Trim(Ado_datos.Recordset!grupo_codigo) & "\"
      Frmexporta.DirDestino.Path = NombreCarpeta
      GlArch = "FOTB"
'      If GlServidor = "SRVPRO" Then
'         e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'      Else
         e = NombreCarpeta
'      End If
      Frmexporta.DirDestino2.Path = e
      Frmexporta.Show vbModal
    Else
      'MsgBox ""
      sino = MsgBox("El archivo ya existe, desea Volver a Cargarlo ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
          NombreCarpeta = App.Path & "\BIENES\" & Trim(Ado_datos.Recordset!grupo_codigo) & "\"
          Frmexporta.DirDestino.Path = NombreCarpeta
          GlArch = "FOTB"
'          If GlServidor = "SRVPRO" Then
'            e = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" & Trim(Ado_datos.Recordset!iniciales) & "-" & Trim(Ado_datos.Recordset!codigo_beneficiario) & "\"
'          Else
            e = NombreCarpeta
'          End If
          Frmexporta.DirDestino2.Path = e
          Frmexporta.Show vbModal
      End If
    End If

    Dim ARCH_FOTO As String
'    If GlServidor = "SRVPRO" Then
'        ARCH_FOTO = "\\" & Trim(GlServidor) & "\SIGPER\PERSONAL\" + Trim(Ado_datos.Recordset!iniciales) + "-" + Trim(Ado_datos.Recordset("codigo_beneficiario")) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_FOTO)
'    Else
        ARCH_FOTO = App.Path + "\BIENES\" + Trim(Ado_datos.Recordset!grupo_codigo) + "\" + Trim(Ado_datos.Recordset!ARCHIVO_Foto)
'    End If
    'ARCH_FOTO = App.Path + "\" + "PERSONAL" + "\" + Ado_datos.Recordset!codigo_beneficiario + "\" + Ado_datos.Recordset("codigo_beneficiario") + "-FOTO.JPG"
    CodBien = Ado_datos.Recordset!bien_codigo
    If Guardar_Imagen(db, "Select Foto From ac_bienes_foto Where bien_codigo= '" & CodBien & "' ", "Foto", ARCH_FOTO) Then
        MsgBox "Se cargo la Imagen Correctamente !!"
    Else
        MsgBox "ERROR No existe la Imagen, Verifique por Favor..."
    End If
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
'    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub BtnGrabar_Click()
On Error GoTo QError
   If Valida Then
      Screen.MousePointer = vbHourglass
        ' Empezar a grabar
        '*********************************
    VAR_NEW = "N"
    If Left(TxtDetalle, 2) = "NA" And estado = 1 Then
      sino = MsgBox("Desea crear un nuevo código de Equipo ? ", vbYesNo + vbQuestion, "Atención ...")
      If sino = vbYes Then
         Set rs_aux6 = New ADODB.Recordset
         If rs_aux6.State = 1 Then rs_aux6.Close
         rs_aux6.Open "select * from fc_partida_gasto where par_codigo = '43340' ", db, adOpenKeyset, adLockReadOnly
         If rs_aux6.RecordCount > 0 Then
            VAR_OA = "OA36" + LTrim(Str(rs_aux6!correlativo36 + 1))
            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "select * from ac_bienes where bien_codigo = '" & VAR_OA & "' ", db, adOpenKeyset, adLockReadOnly
            If rs_aux7.RecordCount > 0 Then
                MsgBox "El Código de Equipo " + VAR_OA + " YA Existe, Vuelva a Intentar !! ", vbExclamation, "Atención!"
                'db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
                VAR_NEW = "N"
                Exit Sub
            Else
                'ado_datos14.Recordset!bien_codigo = Trim(VAR_OA)
                db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
                VAR_NEW = "S"
            End If
         Else
            VAR_NEW = "N"
         End If
      Else
            VAR_OA = Trim(TxtDetalle.Text)
            VAR_NEW = "N"
      End If
    Else
        If Left(TxtDetalle, 3) = "EQP" And estado = 1 Then
          sino = MsgBox("Desea crear un nuevo código de Equipo ? ", vbYesNo + vbQuestion, "Atención ...")
          If sino = vbYes Then
             Set rs_aux6 = New ADODB.Recordset
             If rs_aux6.State = 1 Then rs_aux6.Close
             rs_aux6.Open "select * from fc_partida_gasto where par_codigo = '43340' ", db, adOpenKeyset, adLockReadOnly
             If rs_aux6.RecordCount > 0 Then
                var_cod = rs_aux6!correlativo_eqp + 1
                If var_cod < 10 Then
                   VAR_OA = "EQP" + "000" + LTrim(Str(var_cod))
                End If
                If var_cod > 9 And var_cod < 100 Then
                   VAR_OA = "EQP" + "00" + LTrim(Str(var_cod))
                End If
                If var_cod > 99 And var_cod < 1000 Then
                   VAR_OA = "EQP" + "0" + LTrim(Str(var_cod))
                End If
                If var_cod > 999 Then
                   VAR_OA = "EQP" + LTrim(Str(var_cod))
                End If
                'VAR_OA = "EQP" + LTrim(Str(rs_aux6!correlativo_eqp + 1))
                Set rs_aux7 = New ADODB.Recordset
                If rs_aux7.State = 1 Then rs_aux7.Close
                rs_aux7.Open "select * from ac_bienes where bien_codigo = '" & VAR_OA & "' ", db, adOpenKeyset, adLockReadOnly
                If rs_aux7.RecordCount > 0 Then
                    MsgBox "El Código de Equipo " + VAR_OA + " YA Existe, Vuelva a Intentar !! ", vbExclamation, "Atención!"
                    'db.Execute "update fc_partida_gasto set correlativo36 = correlativo36 + 1 where par_codigo = '43340' "
                    VAR_NEW = "N"
                    Exit Sub
                Else
                    'ado_datos14.Recordset!bien_codigo = Trim(VAR_OA)
                    db.Execute "update fc_partida_gasto set correlativo_eqp = correlativo_eqp + 1 where par_codigo = '43340' "
                    VAR_NEW = "S"
                End If
             Else
                VAR_NEW = "N"
             End If
          Else
                VAR_OA = Trim(TxtDetalle.Text)
                VAR_NEW = "N"
          End If
        End If
    End If
      db.BeginTrans
        'JQA 04/2008
      If swnuevo = True Then
        Ado_datos.Recordset!grupo_codigo = IIf(TxtGrupo.Caption = "", "40000", Trim(TxtGrupo.Caption))
        Ado_datos.Recordset!subgrupo_codigo = IIf(TxtSub.Caption = "", "43000", Trim(TxtSub.Caption))
        Ado_datos.Recordset!par_codigo = IIf(txt_par.Caption = "", "43340", Trim(txt_par.Caption))
        If VAR_NEW = "N" Then
            'Ado_datos.Recordset!bien_codigo = Trim(TxtDetalle.Text)
            CodBien = Trim(TxtDetalle.Text)
        Else
            'Ado_datos.Recordset!bien_codigo = Trim(VAR_OA)
            CodBien = Trim(VAR_OA)
        End If
        Ado_datos.Recordset!bien_codigo = CodBien
        
        Ado_datos.Recordset!ARCHIVO_Foto = "Cargar_Archivo"
        Ado_datos.Recordset!bien_descripcion = txtDescripcion.Text + " - " + TxtInicial
      End If
      If swnuevo = False Then
        Ado_datos.Recordset!bien_descripcion = txtDescripcion.Text
        CodBien = Ado_datos.Recordset!bien_codigo
      End If
        Ado_datos.Recordset!bien_descripcion_anterior = TxtDescripcion2.Text
        Ado_datos.Recordset!unimed_codigo = IIf(Unidad.Text = "", "EQP", Unidad.Text)
        Ado_datos.Recordset!marca_codigo = IIf(marcas.Text = "", "S/M", marcas.Text)
        Ado_datos.Recordset!modelo_codigo = IIf(dtc_codigo6.Text = "", "S/M", dtc_codigo6.Text)
        Ado_datos.Recordset!bien_cantidad_por_empaque = IIf(TxtCantidad.Text = "", 2, TxtCantidad.Text)
        ' Campos no liga
        'Ado_datos.Recordset!estado = IIf(chkEstado.Value = vbChecked, 1, 0)
'        Ado_datos.Recordset!StockInicial = IIf(TxtInicial.Text = "", 0, Val(TxtInicial.Text))      'Val(TxtInicial.Text)
        Ado_datos.Recordset!bien_codigo_anterior = TxtInicial.Text
        Ado_datos.Recordset!bien_codigo_universal = IIf(dtc_codigo8.Text = "", "X", dtc_codigo8.Text) 'TxtInicial.Text
        Ado_datos.Recordset!bien_precio_compra = IIf(TxtPrecComp.Text = "", 0, CDbl(TxtPrecComp.Text))      'CDbl(TxtPrecComp.Text)
        Ado_datos.Recordset!bien_precio_venta_base = IIf(TxtPrecVenta.Text = "", 0, CDbl(TxtPrecVenta.Text))      'CDbl(txtStockMin)
        Ado_datos.Recordset!bien_precio_venta_final = IIf(TxtPrecEst.Text = "", 0, CDbl(TxtPrecEst.Text))      'CDbl(TxtPrecEst)
        Ado_datos.Recordset!bien_precio_compra_dol = IIf(TxtPrecCompD.Text = "", 0, CDbl(TxtPrecCompD.Text))            'DOLARES
        Ado_datos.Recordset!bien_precio_venta_base_dol = IIf(TxtPrecVentaD.Text = "", 0, CDbl(TxtPrecVentaD.Text))      'DOLARES
        Ado_datos.Recordset!bien_precio_venta_final_dol = IIf(TxtPrecEstD.Text = "", 0, CDbl(TxtPrecEstD.Text))         'DOLARES
        Ado_datos.Recordset!bien_stock_inicial = IIf(txtStockIni.Text = "", 0, CDbl(txtStockIni.Text))      'CDbl(txtStockMin)
        Ado_datos.Recordset!bien_stock_minimo = IIf(txtStockMin.Text = "", 0, CDbl(txtStockMin.Text))
        If txtCantComprada.Caption = "" Then
            txtCantComprada.Caption = "0"
        End If
        
        If TxtActual.Caption = "" Then
            TxtActual.Caption = "0"
        End If
        
        If txtCantVendida = "" Then
            txtCantVendida.Caption = "0"
        End If
        
        Ado_datos.Recordset!bien_stock_ingreso = IIf(txtCantComprada.Caption = "", 0, CDbl(txtCantComprada.Caption)) 'CDbl(txtStockMin)
        Ado_datos.Recordset!bien_stock_salida = IIf(txtCantVendida = "", 0, CDbl(txtCantVendida))
        Ado_datos.Recordset!bien_stock_actual = IIf(TxtActual.Caption = "", 0, CDbl(TxtActual.Caption))

        Ado_datos.Recordset!observaciones = IIf(dtc_desc10.Text = "", "NO ASIGNADO", dtc_desc10.Text)
        
        Ado_datos.Recordset!bien_rotacion = IIf(cmd_rotacion.Text = "", "PROMEDIO", cmd_rotacion.Text)      'CDbl(txtStockMin)
        Ado_datos.Recordset!edif_codigo = IIf(dtc_codigo10.Text = "", "20101-0", dtc_codigo10.Text)      'CDbl(txtStockMin)
        'Ado_datos.Recordset!tipo_eqp = IIf(dtc_codigo8.Text = "", "X", dtc_codigo8.Text)
        Ado_datos.Recordset!pais_codigo = DtcPais.Text
        'Ado_datos.Recordset!ARCHIVO_F = Trim(Ado_datos.Recordset!subgrupo_codigo) + "-" + Trim(Ado_datos.Recordset!bien_codigo) + ".JPG"
        Ado_datos.Recordset!archivo_foto2 = Trim(Ado_datos.Recordset!bien_codigo) + ".JPG"
        Ado_datos.Recordset!estado_codigo = "REG"  'chkEstado
        Ado_datos.Recordset!usr_codigo = glusuario
        Ado_datos.Recordset!Fecha_Registro = Date
        Ado_datos.Recordset!hora_registro = Format(Time, "hh:mm:ss")
        '*********************************
        Ado_datos.Recordset!codigo_pSIN = IIf(IsNull(TxtCodigo_pSIN.Text), "", TxtCodigo_pSIN.Text)
        Ado_datos.Recordset!descripcion_pSIN = IIf(IsNull(TxtDescripcionSIN.Text), "", TxtDescripcionSIN.Text)
        Ado_datos.Recordset!correlativo_pSIN = IIf(IsNull(Dtc_codigoSIN.Text), 9, Dtc_codigoSIN.Text)
        '*********************************
        ' Grabar
        Ado_datos.Recordset.Update
        db.CommitTrans
    '*********************************
'        Ado_datos.Caption = "Registro: " & CStr(Ado_datos.Recordset.AbsolutePosition) & " de " & Ado_datos.Recordset.RecordCount
        ' Colocar los botones en modo navegar
        GlHayRegs = True
        'BotonesNavegar Me
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        FraArticulos.Enabled = False
'        TxtGrupo.Enabled = True
'        DtcGrupoDes.Enabled = True
'        TxtSub.Enabled = True
'        dtc_sub_des.Enabled = True
        Screen.MousePointer = vbDefault
        marca1 = Ado_datos.Recordset.Bookmark
        If swnuevo = True Then
            MsgBox "El Código de Equipo " + VAR_OA + " fue generado satisfactoriamente !! ", vbExclamation, "Atención!"
        End If
'        Call CARGA
'        Ado_datos.Recordset.Move marca1 - 1
        'Ado_datos.Recordset.MoveLast
        'Set dg_datos.DataSource = Ado_datos
        'Refrescar Grid
'        If OptFilGral1.Value = True Then
'           Call OptFilGral1_Click        'Pendientes
'        Else
'           Call OptFilGral2_Click        'TODOS
'        End If
        Call OptFilGral2_Click
        If (dg_datos.SelBookmarks.Count <> 0) Then
           dg_datos.SelBookmarks.Remove 0
        End If
        If Ado_datos.Recordset.RecordCount > 0 And estado = 2 Then
           RsArt.Find "bien_codigo = '" & CodBien & "'   ", , , 1
           dg_datos.SelBookmarks.Add (RsArt.Bookmark)
        Else
           RsArt.MoveLast
        End If
             
        estado = 0
        'CARGA
        swnuevo = False
        dg_datos.Enabled = True
   Else
        MsgBox "Existe un error en los datos registrados, Verifique y vuelva a intentar...", vbExclamation + vbOKOnly, "Atención"
   
   End If
   swnuevo = False
   Exit Sub
QError:
    ' Manejo de errores
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
    db.RollbackTrans
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRefrescar_Click()
On Error GoTo Que_Error
    Screen.MousePointer = vbHourglass
    Ado_datos.Recordset.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
Que_Error:
    ' Manejo de errores
    Screen.MousePointer = vbDefault
    MsgBox Err.Number & " : " & Err.Description, vbExclamation + vbOKOnly, "Atención"
End Sub

Private Sub BtnImprimirA_Click()
  Dim iResult As Integer
'     LiteralCry = Str(Round(AdoRegularizacion.Recordset!monto_Bolivianos, 2))
'  Literal2 = Literal(LiteralCry) + "  Bolivianos"
'  org2 = AdoRegularizacion.Recordset!org_codigo
'  cocmCod_Comp = AdoRegularizacion.Recordset!codigo_pago
  With CryBBSS
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
'    .StoredProcParam(0) = org2
'    .StoredProcParam(1) = cocmCod_Comp
'    .StoredProcParam(2) = Literal2
        .ReportFileName = App.Path & "\Reportes\Almacen\productos.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub

Private Sub BtnImprimir_Click()
  db.Execute "UPDATE AC_BIENES SET AC_BIENES.observaciones = gc_edificaciones.edif_descripcion FROM AC_BIENES inner join gc_edificaciones on AC_BIENES.edif_codigo  = gc_edificaciones.edif_codigo where par_codigo = '43340' "
  
  db.Execute "UPDATE AC_BIENES SET AC_BIENES.estado_vigente  = 'NO' where AC_BIENES.par_codigo = '43340' "

  db.Execute "UPDATE AC_BIENES SET AC_BIENES.estado_vigente  = 'SI' FROM AC_BIENES inner join ao_ventas_detalle on AC_BIENES.bien_codigo   = ao_ventas_detalle.bien_codigo where AC_BIENES.par_codigo = '43340' AND ao_ventas_detalle.par_codigo = '43340' "
  
  Dim iResult As Integer
  With CryLista
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
        '.ReportFileName = App.Path & "\Reportes\Almacen\Alm_Listado_Gral_Productos.rpt"
        .ReportFileName = App.Path & "\Reportes\Clasificadores\ar_bienes_equipos.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub CmdCancelaDet_Click()
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "SELECT * FROM ac_bienes_insumos_kit where bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux4.RecordCount > 0 Then
       dg_datos2.Visible = True
       Set Ado_datos2.Recordset = rs_aux4
       Set dg_datos2.DataSource = Ado_datos2.Recordset
    Else
       dg_datos2.Visible = False
    End If
    FraInsumo.Visible = False
End Sub

Private Sub CmdGrabaDet_Click()
  If dtc_codigo6Z = "" Then
        MsgBox "Debe Elejir: " + lbl_insumo1.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
  End If
  If dtc_codigo6A = "" Then
        MsgBox "Debe Elejir : " + lbl_insumo2.Caption + ", !! Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
  End If
  If dtc_codigo6B = "" Then
        MsgBox "Debe Registrar: " + lbl_insumo3.Caption + ", !! en el Proyecto de Edificación, Vuelva a Intentar ...", vbExclamation, "Atención"
        Exit Sub
  End If
    If VAR_SW2 = "ADD" Then
        db.Execute "INSERT INTO ac_bienes_insumos_kit (bien_codigo, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, estado_codigo, usr_codigo, fecha_registro) " & _
        " VALUES ('" & Ado_datos.Recordset!bien_codigo & "', '" & dtc_codigo6Z.Text & "', '" & dtc_codigo6A.Text & "', '" & dtc_codigo6B.Text & "', '" & dtc_codigo6C.Text & "', '" & dtc_codigo6D.Text & "', " & _
        " " & IIf(Txt_cant1.Text = "", 0, Txt_cant1.Text) & ", " & IIf(Txt_cant2.Text = "", 0, Txt_cant2.Text) & ", " & IIf(Txt_cant3.Text = "", 0, Txt_cant3.Text) & " , " & IIf(Txt_cant4.Text = "", 0, Txt_cant4.Text) & ", " & IIf(Txt_cant5.Text = "", 0, Txt_cant5.Text) & ",  " & _
        " 'APR', '" & glusuario & "', '" & Date & "') "
    Else
        'bien_codigo, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, estado_codigo, usr_codigo, fecha_registro
        db.Execute "update ac_bienes_insumos_kit set bien_codigo1 = '" & dtc_codigo6Z.Text & "', bien_codigo2 = '" & dtc_codigo6A.Text & "', bien_codigo3 = '" & dtc_codigo6B.Text & "', bien_codigo4 = '" & dtc_codigo6C.Text & "', bien_codigo5 = '" & dtc_codigo6D.Text & "', " & _
        " cantidad1 = " & IIf(Txt_cant1.Text = "", 0, Txt_cant1.Text) & ", cantidad2 = " & IIf(Txt_cant2.Text = "", 0, Txt_cant2.Text) & ", cantidad3 = " & IIf(Txt_cant3.Text = "", 0, Txt_cant3.Text) & " , cantidad4 = " & IIf(Txt_cant4.Text = "", 0, Txt_cant4.Text) & " , cantidad5 = " & IIf(Txt_cant5.Text = "", 0, Txt_cant5.Text) & "  " & _
        " where bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'  "
    End If
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "SELECT * FROM ac_bienes_insumos_kit where bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "' ", db, adOpenKeyset, adLockOptimistic
    If rs_aux4.RecordCount > 0 Then
       dg_datos2.Visible = True
       Set Ado_datos2.Recordset = rs_aux4
       Set dg_datos2.DataSource = Ado_datos2.Recordset
    Else
       dg_datos2.Visible = False
    End If
    FraInsumo.Visible = False
End Sub

Private Sub dg_datos_Click()
'    If buscados = 0 Then
'        'OptFilGral2.Visible = True
'        'OptFilGral1.Visible = True
'
'    End If
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo6A_Click(Area As Integer)
    dtc_desc6A.BoundText = dtc_codigo6A.BoundText
End Sub

Private Sub dtc_codigo6B_Click(Area As Integer)
    dtc_desc6B.BoundText = dtc_codigo6B.BoundText
End Sub

Private Sub dtc_codigo6C_Click(Area As Integer)
    dtc_desc6C.BoundText = dtc_codigo6C.BoundText
End Sub

Private Sub dtc_codigo6D_Click(Area As Integer)
    dtc_desc6D.BoundText = dtc_codigo6D.BoundText
End Sub

Private Sub dtc_codigo6Z_Click(Area As Integer)
    dtc_desc6Z.BoundText = dtc_codigo6Z.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc6A_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6B_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6C_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6D_Click(Area As Integer)
    dtc_codigo6A.BoundText = dtc_desc6A.BoundText
End Sub

Private Sub dtc_desc6Z_Click(Area As Integer)
    dtc_codigo6Z.BoundText = dtc_desc6Z.BoundText
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_partida_Click(Area As Integer)
    dtc_partida_des.BoundText = dtc_partida.BoundText
End Sub

Private Sub dtc_partida_des_Click(Area As Integer)
    dtc_partida.BoundText = dtc_partida_des.BoundText
End Sub

Private Sub DtcGrupoCod_Click(Area As Integer)
    DtcGrupoDes.BoundText = DtcGrupoCod.BoundText
    DtcGrupoUni.BoundText = DtcGrupoCod.BoundText
End Sub

Private Sub DtcGrupoDes_Click(Area As Integer)
   DtcGrupoCod.BoundText = DtcGrupoDes.BoundText
   DtcGrupoUni.BoundText = DtcGrupoDes.BoundText
'   Call pOrganismo(DtcGrupoCod.BoundText)
'   dtc_sub_des.Enabled = True
End Sub

Private Sub pOrganismo(CodFuente As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from ac_bienes_subgrupo where grupo_codigo='" & CodFuente & "'"
   
   Set dtc_sub_cod.RowSource = Nothing
   Set dtc_sub_cod.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_sub_cod.ReFill
   dtc_sub_cod.BoundText = Empty
   
   Set dtc_sub_des.RowSource = Nothing
   Set dtc_sub_des.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_sub_des.ReFill
   dtc_sub_des.BoundText = Empty

End Sub

Private Sub DtcGrupoUni_Click(Area As Integer)
    DtcGrupoDes.BoundText = DtcGrupoUni.BoundText
    DtcGrupoCod.BoundText = DtcGrupoUni.BoundText
End Sub

Private Sub DtcPais_Click(Area As Integer)
    DtcPaisD.BoundText = DtcPais.BoundText
End Sub

Private Sub DtcPaisD_Click(Area As Integer)
    DtcPais.BoundText = DtcPaisD.BoundText
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys vbTab
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Dim Nodo As Node
    Me.Top = 0
    Me.Left = 0
    Screen.MousePointer = vbHourglass
    estado = 0
    
    ' Cargamos el Arbol
    ' Nodo Principal
'    Set Nodo = trv.Nodes.Add(, , "Grupo", "Grupos", "Grupos")
'    Nodo.Expanded = True
'    Nodo.Bold = True
    dtc_codigo10.Visible = False
    dtc_desc10.Visible = False
    lbl_edif.Visible = False
'    dtc_codigo8.Visible = False
    dtc_desc8.Visible = False
    lbl_eqp.Visible = False
        
'    OptFilGral1 = True
    'Call OptFilGral1_Click
    Call OptFilGral2_Click
    Call CARGA
'    Set rsgrupo = New ADODB.Recordset
'    rsgrupo.Open "SELECT * FROM ALClGrupo ORDER BY CAST (grupo_codigo AS INT) ", db, adOpenStatic
'    Set AdoGrupo.Recordset = rsgrupo
'    If rsgrupo.RecordCount > 0 Then
'      rsgrupo.MoveFirst
'      While Not rsgrupo.EOF
'        Set Nodo = trv.Nodes.Add("Grupo", tvwChild, "D" & Trim(rsgrupo!grupo_codigo), rsgrupo!descgrupo, "NoElegido", "Elegido")
'        rsgrupo.MoveNext
'      Wend
'    Else
'        trv.Nodes(1).Text = "No Existen Grupos Creados..."
'    End If
    '
    'BotonesNavegar Me
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = False
    FraArticulos.Enabled = False
    Screen.MousePointer = vbDefault
    C_FIJO = 0      '1.92
    C_MANOBR = 0.01
    C_GTOADM = 0.6094
    C_UTILID = 0.2
    C_ROTALT = 0.01
    C_ROTBAJ = 0.02
    C_FIJO = 0.0636
    C_IMPSTO2 = 0.1494
        Call SeguridadSet(Me)
End Sub

Private Sub OptFilGral2_Click()
    Set RsArt = New ADODB.Recordset
    'JQA 04/2008
    If RsArt.State = 1 Then RsArt.Close
    'queryinicial = "SELECT * FROM ac_bienes WHERE Estado <> 2 "   'ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion
    queryinicial = "SELECT * FROM ac_bienes  where par_codigo = '43340'  "       'where estado_codigo <> 'ER' "   'ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion
    RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    RsArt.Sort = "grupo_codigo, subgrupo_codigo, bien_codigo"
    If RsArt.RecordCount > 0 Then
       GlHayRegs = True  'Variable global
    Else
       GlHayRegs = False
    End If
    Set Ado_datos.Recordset = RsArt
    'Set dg_datos.DataSource = Ado_datos.Recordset
'    Ado_datos.Recordset.Requery
'    Ado_datos.Refresh
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral1_Click()
    Set RsArt = New ADODB.Recordset
    'JQA 04/2008
    If RsArt.State = 1 Then RsArt.Close
    'queryinicial = "SELECT * FROM ac_bienes WHERE Estado <> 2 "   'ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion
    queryinicial = "SELECT * FROM ac_bienes WHERE estado_codigo = 'REG'  and par_codigo = '43340'  "   'ORDER BY grupo_codigo, subgrupo_codigo, bien_descripcion
    RsArt.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    'RsArt.Sort = "grupo_codigo, subgrupo_codigo"
    RsArt.Sort = "grupo_codigo, subgrupo_codigo, bien_codigo"
    If RsArt.RecordCount > 0 Then
       GlHayRegs = True  'Variable global
    Else
       GlHayRegs = False
    End If
    Set Ado_datos.Recordset = RsArt
    'Set dg_datos.DataSource = Ado_datos.Recordset
'    Ado_datos.Recordset.Requery
'    Ado_datos.Refresh
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Function Valida() As Boolean
    Valida = False
    'If swnuevo <> True Then
'    If Trim(TxtGrupo.Caption) = "" Then
'        MsgBox "Elija el Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            DtcGrupoDes.SetFocus
'        End If
'        Exit Function
'    End If
'    If Trim(TxtSub.Caption) = "" Then
'        MsgBox "Elija el Sub-Grupo al Cual pertenece el Detalle.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            DtcGrupoDes.SetFocus
'        End If
'        Exit Function
'    End If
    If Trim(TxtDetalle.Text) = "" Then
        MsgBox "Ingrese el Codigo del Detalle.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            TxtDetalle.SetFocus
        End If
        Exit Function
    End If
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la Descripción del Detalle.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            txtDescripcion.SetFocus
        End If
        Exit Function
    End If
    If Trim(Unidad.Text) = "" Then
        MsgBox "Ingrese la Unidad de Medida.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            Unidad.SetFocus
'        End If
        Exit Function
    End If
    If Trim(DtcPais.Text) = "" Then
        MsgBox "Ingrese la Industria (Pais Origen).", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            Unidad.SetFocus
'        End If
        Exit Function
    End If
    
     If Trim(marcas.Text) = "" Then
        MsgBox "Ingrese la Marca.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            Unidad.SetFocus
'        End If
        Exit Function
    End If
    
    If Trim(TxtPrecComp.Text) = "" Then
        MsgBox "Ingrese EL Precio de Compra del Detalle.", vbExclamation + vbOKOnly, "Atención"
'        If estado <> 0 Then
'            TxtPrecComp.SetFocus
'        End If
  Exit Function
    End If
    If Trim(txtStockMin.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Salon del Detalle.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            txtStockMin.SetFocus
        End If
        Exit Function
    End If
    If Trim(TxtPrecEst.Text) = "" Then
        MsgBox "Ingrese el Precio de Venta Cliente del Detalle.", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            TxtPrecEst.SetFocus
        End If
        Exit Function
    End If
    If dtc_codigo6.Text = "" Or dtc_desc6.Text = "" Then
    MsgBox "Ingrese el modelo", vbExclamation + vbOKOnly, "Atención"
'        MsgBox "El MODELO Registrado es incorrecto, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
        If estado <> 0 Then
            dtc_codigo6.SetFocus
        End If
        Exit Function
    End If
    'If TxtGrupo.Caption = "40000" Then
'    If txt_par.Caption = "43340" Then
'        If dtc_codigo8.Text = "" Then
'            MsgBox "El TIPO de EQUIPO Registrado es incorrecto, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
'            If estado <> 0 Then
'                dtc_codigo8.SetFocus
'            End If
'            Exit Function
'        End If
'    End If
    If txtStockIni.Text = "" Then
            MsgBox "Ingrese el Stock inicial, verifique y vuelva a intentar ... ", vbExclamation + vbOKOnly, "Atención"
            If estado <> 0 Then
                txtStockIni.SetFocus
            End If
            Exit Function
    End If
    Valida = True
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set ClBuscaGrid = Nothing
End Sub

Private Sub Imprimir_Click()
  Dim iResult As Integer
'     LiteralCry = Str(Round(AdoRegularizacion.Recordset!monto_Bolivianos, 2))
'  Literal2 = Literal(LiteralCry) + "  Bolivianos"
'  org2 = AdoRegularizacion.Recordset!org_codigo
'  cocmCod_Comp = AdoRegularizacion.Recordset!codigo_pago
  With CryFis
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .WindowShowPrintSetupBtn = True
    .WindowShowGroupTree = True
    .WindowShowExportBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowShowSearchBtn = True
'    .StoredProcParam(0) = org2
'    .StoredProcParam(1) = cocmCod_Comp
'    .StoredProcParam(2) = Literal2
        .ReportFileName = App.Path & "\Reportes\Almacen\productos_inventario.rpt"
    iResult = .PrintReport
    If iResult <> 0 Then
        MsgBox .LastErrorNumber & " : " & .LastErrorString, vbCritical + vbOKOnly, "Error..."
    End If
  End With

End Sub

Private Sub marcas_Click(Area As Integer)
    TDBC_marcas.BoundText = marcas.BoundText
End Sub

Private Sub TDBC_marcas_Click(Area As Integer)
    marcas.BoundText = TDBC_marcas.BoundText
End Sub

Private Sub dtc_sub_cod_Click(Area As Integer)
    dtc_sub_des.BoundText = dtc_sub_cod.BoundText
End Sub

Private Sub dtc_sub_des_Click(Area As Integer)
    dtc_sub_cod.BoundText = dtc_sub_des.BoundText
'    Call pPartida(dtc_sub_cod.BoundText)
'    dtc_partida_des.Enabled = True
End Sub

Private Sub pPartida(CodPar As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from fc_partida_gasto where subgrupo_codigo='" & CodPar & "' AND estado_codigo = 'APR' "
   
   Set dtc_partida.RowSource = Nothing
   Set dtc_partida.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_partida.ReFill
   dtc_partida.BoundText = Empty
   
   Set dtc_partida_des.RowSource = Nothing
   Set dtc_partida_des.RowSource = db.Execute(strConsultaF, , adCmdText)
   dtc_partida_des.ReFill
   dtc_partida_des.BoundText = Empty

End Sub

Private Sub dtc_sub_des_LostFocus()
    If TxtSub.Caption = "43000" Then
        dtc_codigo10.Visible = True
        dtc_desc10.Visible = True
        lbl_edif.Visible = True
'        dtc_codigo8.Visible = True
        dtc_desc8.Visible = True
        lbl_eqp.Visible = True
    Else
        dtc_codigo10.Visible = False
        dtc_desc10.Visible = False
        lbl_edif.Visible = False
'        dtc_codigo8.Visible = False
        dtc_desc8.Visible = False
        lbl_eqp.Visible = False
    End If
End Sub
    
Private Sub TDBC_Unidad_Click(Area As Integer)
    Unidad.BoundText = TDBC_Unidad.BoundText
End Sub

Private Sub TDBC_Unidad_LostFocus()
    If Unidad.Text = "EQP" Then
        dtc_desc8.Visible = True
        lbl_eqp.Visible = True
    Else
        dtc_desc8.Visible = False
        lbl_eqp.Visible = False
    End If
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtPrecComp_LostFocus()
    If TxtPrecComp = "" Then
        TxtPrecCompD = 0
        TxtPrecComp = 0
        TxtPrecVenta = 0
        TxtPrecVentaD = 0
        TxtPrecEst = 0
        TxtPrecEstD = 0
    Else
        TxtPrecCompD = Round(CDbl(TxtPrecComp) / GlTipoCambioOficial, 2)
        'TxtPrecVenta = Round(CDbl(TxtPrecComp) * C_FIJO, 2)
        TxtPrecVenta = Round(CDbl(TxtPrecComp) + (CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTALT) + (CDbl(TxtPrecComp) * C_IMPSTO), 2)
        TxtPrecVentaD = Round(CDbl(TxtPrecVenta) / GlTipoCambioOficial, 2)
        If cmd_rotacion.Text = "ALTA" Then
            C_FIJO = Round((CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTALT), 2)
            TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTALT) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
            'TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(TxtPrecVenta) * C_MANOBR) + (CDbl(TxtPrecVenta) * C_GTOADM) + (CDbl(TxtPrecVenta) * C_UTILID) + (CDbl(TxtPrecVenta) * C_ROTALT) + (CDbl(TxtPrecVenta) * C_IMPSTO), 2)
        Else
            C_FIJO = Round((CDbl(TxtPrecComp) * C_MANOBR) + (CDbl(TxtPrecComp) * C_GTOADM) + (CDbl(TxtPrecComp) * C_UTILID) + (CDbl(TxtPrecComp) * C_ROTBAJ), 2)
            TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTBAJ) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
            'TxtPrecEst = Round(CDbl(TxtPrecVenta) + (CDbl(TxtPrecVenta) * C_MANOBR) + (CDbl(TxtPrecVenta) * C_GTOADM) + (CDbl(TxtPrecVenta) * C_UTILID) + (CDbl(TxtPrecVenta) * C_ROTBAJ) + (CDbl(TxtPrecVenta) * C_IMPSTO), 2)
        End If
        TxtPrecEstD = Round(CDbl(TxtPrecEst) / GlTipoCambioOficial, 2)
    End If
End Sub

Private Sub TxtPrecCompD_LostFocus()
    If TxtPrecCompD = "" Then
        TxtPrecCompD = 0
        TxtPrecComp = 0
        TxtPrecVenta = 0
        TxtPrecVentaD = 0
        TxtPrecEst = 0
        TxtPrecEstD = 0
    Else
        TxtPrecComp = Round(CDbl(TxtPrecCompD) * GlTipoCambioOficial, 2)
        'TxtPrecVentaD = Round(CDbl(TxtPrecCompD) * C_FIJO, 2)
        TxtPrecVentaD = Round(CDbl(TxtPrecCompD) + (CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_IMPSTO), 2)
        TxtPrecVenta = Round(CDbl(TxtPrecVentaD) * GlTipoCambioOficial, 2)
        If cmd_rotacion.Text = "ALTA" Then
            C_FIJO = Round((CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_UTILID) + (CDbl(TxtPrecCompD) * C_ROTALT), 2)
            'TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(TxtPrecVentaD) * C_MANOBR) + (CDbl(TxtPrecVentaD) * C_GTOADM) + (CDbl(TxtPrecVentaD) * C_UTILID) + (CDbl(TxtPrecVentaD) * C_ROTALT) + (CDbl(TxtPrecVentaD) * C_IMPSTO), 2)
            TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTALT) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
        Else
            C_FIJO = Round((CDbl(TxtPrecCompD) * C_MANOBR) + (CDbl(TxtPrecCompD) * C_GTOADM) + (CDbl(TxtPrecCompD) * C_UTILID) + (CDbl(TxtPrecCompD) * C_ROTBAJ), 2)
            'TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(TxtPrecVentaD) * C_MANOBR) + (CDbl(TxtPrecVentaD) * C_GTOADM) + (CDbl(TxtPrecVentaD) * C_UTILID) + (CDbl(TxtPrecVentaD) * C_ROTBAJ) + (CDbl(TxtPrecVentaD) * C_IMPSTO), 2)
            TxtPrecEstD = Round(CDbl(TxtPrecVentaD) + (CDbl(C_FIJO) * C_UTILID) + (CDbl(C_FIJO) * C_ROTBAJ) + (CDbl(C_FIJO) * C_IMPSTO2), 2)
        End If
        TxtPrecEst = Round(CDbl(TxtPrecEstD) * GlTipoCambioOficial, 2)
    End If
End Sub



Private Sub Unidad_Click(Area As Integer)
    TDBC_Unidad.BoundText = Unidad.BoundText
End Sub

'Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
'    If InStr(Node.Key, "G") = 0 Then
'        TxtGrupo.caption = Mid(Node.Key, 2)
'    Else
'        TxtGrupo.caption = ""
'    End If
'End Sub

'Private Sub BuscaNodo(QNodo As String)
'Dim Nodo As Node
'Dim Indice As Integer
'    For Indice = 1 To trv.Nodes.Count
'        If Mid(trv.Nodes(Indice).Key, 2) = QNodo Then
'            trv.Nodes(Indice).Selected = True
'            Exit For
'        End If
'    Next
'End Sub

'Private Sub txtStockMin_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
'End Sub
'Private Sub txtUnidadCaja_KeyPress(KeyAscii As Integer)
'    KeyAscii = IIf(Chr(KeyAscii) Like "[0-9]", KeyAscii, 0)
'End Sub

Private Function ExisteDetalle(bien_codigo As String) As Boolean
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_no_objecion_detalle_D WHERE bien_codigo = '" & bien_codigo & "'"
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud_bienes WHERE bien_codigo = '" & bien_codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteDetalle = rs!Cuantos > 0
End Function

Private Sub CARGA()
    Set rsProductosSIN = New ADODB.Recordset
    If rsProductosSIN.State = 1 Then rsProductosSIN.Close
    rsProductosSIN.Open "SELECT [correlativo] AS correlativo_pSIN, CONCAT([CodProducto], ' ', [Producto], ' || ', [CodActividad], ' ', [ActividadEconomica]) AS descripcion FROM [dbo].[gc_productos_sin] WHERE [EsBien] = CAST(1 AS BIT)", db, adOpenStatic
    Set AdoProductoSin.Recordset = rsProductosSIN
    Dtc_descripcionSIN.BoundText = Dtc_codigoSIN.BoundText
    
    Set rsMarcas = New ADODB.Recordset
    If rsMarcas.State = 1 Then rsMarcas.Close
    rsMarcas.Open "SELECT * FROM ac_bienes_marcas ORDER BY marca_descripcion", db, adOpenStatic
    Set AdoMarca.Recordset = rsMarcas
    
    Set rsUnidad = New ADODB.Recordset
    If rsUnidad.State = 1 Then rsUnidad.Close
    rsUnidad.Open "Select * from ac_bienes_unidad_medida order by unimed_descripcion", db, adOpenStatic
    Set AdoMedida.Recordset = rsUnidad
    
    Set rsSubGrupo = New ADODB.Recordset
    If rsSubGrupo.State = 1 Then rsSubGrupo.Close
    rsSubGrupo.Open "select * from ac_bienes_subgrupo order by subgrupo_descripcion", db, adOpenStatic
    Set AdoSubGrupo.Recordset = rsSubGrupo
    
    Set rsgrupo = New ADODB.Recordset
    If rsgrupo.State = 1 Then rsgrupo.Close
    rsgrupo.Open "SELECT * FROM ac_bienes_grupo WHERE estado_codigo='APR' ", db, adOpenStatic
    Set AdoGrupo.Recordset = rsgrupo
    
    Set rsPais = New ADODB.Recordset
    If rsPais.State = 1 Then rsPais.Close
    rsPais.Open "SELECT * FROM gc_pais WHERE estado_codigo='APR' order by pais_descripcion", db, adOpenStatic
    Set AdoPais.Recordset = rsPais
    
    'gc_edificaciones
    Set rs_datos10 = New ADODB.Recordset
    If rs_datos10.State = 1 Then rs_datos10.Close
    rs_datos10.Open "Select * from gc_edificaciones order by edif_descripcion", db, adOpenStatic
    Set Ado_datos10.Recordset = rs_datos10
    dtc_desc10.BoundText = dtc_codigo10.BoundText
    
    'ac_bienes_modelos
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "Select * from ac_bienes_modelos ", db, adOpenStatic     'order by modelo_descripcion
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    'fc_partidas_gasto
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "Select * from fc_partida_gasto order by par_descripcion", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_partida_des.BoundText = dtc_partida.BoundText
    
    'ac_bienes_equipo_tipos
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "Select * from ac_bienes_equipo_tipos order by tipo_eqp_descripcion", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
End Sub


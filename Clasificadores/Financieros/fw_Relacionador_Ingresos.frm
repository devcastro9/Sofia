VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form fw_Relacionador_Ingresos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clasificadores - Financieros - Relacionador Ingresos "
   ClientHeight    =   8040
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   13080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   27
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   11760
         Picture         =   "fw_Relacionador_Ingresos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10800
         Picture         =   "fw_Relacionador_Ingresos.frx":020A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "fw_Relacionador_Ingresos.frx":064C
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   34
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
         Picture         =   "fw_Relacionador_Ingresos.frx":0E0B
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   33
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
         Picture         =   "fw_Relacionador_Ingresos.frx":1720
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   32
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
         Picture         =   "fw_Relacionador_Ingresos.frx":1E6C
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   31
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
         Picture         =   "fw_Relacionador_Ingresos.frx":269F
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   30
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
         Picture         =   "fw_Relacionador_Ingresos.frx":2E54
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   29
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
         Picture         =   "fw_Relacionador_Ingresos.frx":3721
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   28
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RELACIONADOR INGRESOS"
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
         Left            =   12945
         TabIndex        =   37
         Top             =   195
         Width           =   3315
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
      TabIndex        =   23
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
         Picture         =   "fw_Relacionador_Ingresos.frx":3EE3
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   25
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
         Picture         =   "fw_Relacionador_Ingresos.frx":46B9
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   24
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
         TabIndex        =   26
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   6375
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   12975
      Begin MSDataGridLib.DataGrid dg_datos 
         Bindings        =   "fw_Relacionador_Ingresos.frx":4FA5
         Height          =   5655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   12795
         _ExtentX        =   22569
         _ExtentY        =   9975
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "correlativo"
            Caption         =   "Correl."
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
            DataField       =   "rubro_codigo_I"
            Caption         =   "Rubro Ini"
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
            DataField       =   "rubro_codigo_F"
            Caption         =   "Rubro Fin"
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
            DataField       =   "Descripcion"
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
            DataField       =   "Cta_Deb"
            Caption         =   "Cta Debe"
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
            DataField       =   "SubCta_Deb1"
            Caption         =   "Cta Debe1"
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
            DataField       =   "SubCta_Deb2"
            Caption         =   "Cta Debe2"
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
            DataField       =   "Cta_Cred"
            Caption         =   "Cta Haber"
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
            DataField       =   "Subcta_Cred1"
            Caption         =   "Cta Haber1"
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
            DataField       =   "SubCta_Cred2"
            Caption         =   "Cta Haber2"
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
         BeginProperty Column11 
            DataField       =   "Codigo_Tipo"
            Caption         =   "Tipo Comp"
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
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   5925
         Width           =   12795
         _ExtentX        =   22569
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
         Caption         =   " <-- Inicio                                           Gerencia General                    Fin -->"
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
      BackColor       =   &H00C0C0C0&
      Height          =   6375
      Left            =   13200
      TabIndex        =   9
      Top             =   720
      Width           =   6135
      Begin VB.TextBox txt_porcentaje 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DataField       =   "porcentaje"
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0"
         Top             =   5880
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtc_Correl5 
         Bindings        =   "fw_Relacionador_Ingresos.frx":4FBD
         DataField       =   "Correl_Cred"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   50
         Top             =   5040
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "correl"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_Correl4 
         Bindings        =   "fw_Relacionador_Ingresos.frx":4FD6
         DataField       =   "Correl_Deb"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3120
         TabIndex        =   49
         Top             =   4080
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "correl"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_SubCta2_4 
         Bindings        =   "fw_Relacionador_Ingresos.frx":4FEF
         DataField       =   "Correl_Deb"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   41
         Top             =   4080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "SubCta2"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_SubCta2_5 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5008
         DataField       =   "Correl_Cred"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1800
         TabIndex        =   47
         Top             =   5040
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "SubCta2"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_SubCta1_5 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5021
         DataField       =   "Correl_Cred"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   960
         TabIndex        =   46
         Top             =   5040
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "SubCta1"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin VB.TextBox Txt_descripcion 
         DataField       =   "Descripcion"
         DataSource      =   "Ado_datos"
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Text            =   "-"
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txt_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "correlativo"
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "fw_Relacionador_Ingresos.frx":503A
         DataField       =   "Codigo_Tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   14737632
         ListField       =   "Codigo_Tipo"
         BoundColumn     =   "Codigo_Tipo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5053
         DataField       =   "Codigo_Tipo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Denominacion_Tipo"
         BoundColumn     =   "Codigo_Tipo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "fw_Relacionador_Ingresos.frx":506C
         DataField       =   "Correl_Deb"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5085
         DataField       =   "rubro_codigo_I"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   21
         Top             =   2280
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "rubro_codigo"
         BoundColumn     =   "rubro_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "fw_Relacionador_Ingresos.frx":509E
         DataField       =   "rubro_codigo_I"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "rubro_descripcion"
         BoundColumn     =   "rubro_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "fw_Relacionador_Ingresos.frx":50B7
         DataField       =   "rubro_codigo_F"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "rubro_descripcion"
         BoundColumn     =   "rubro_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_SubCta1_4 
         Bindings        =   "fw_Relacionador_Ingresos.frx":50D0
         DataField       =   "Correl_Deb"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   960
         TabIndex        =   39
         Top             =   4080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "SubCta1"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "fw_Relacionador_Ingresos.frx":50E9
         DataField       =   "rubro_codigo_F"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   42
         Top             =   3000
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   14737632
         ForeColor       =   0
         ListField       =   "rubro_codigo"
         BoundColumn     =   "rubro_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_Cta_Deb 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5102
         DataField       =   "Correl_Deb"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   4080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_Cta_Cred 
         Bindings        =   "fw_Relacionador_Ingresos.frx":511B
         DataField       =   "Correl_Cred"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   5040
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "Cuenta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "fw_Relacionador_Ingresos.frx":5134
         DataField       =   "Correl_Cred"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   5400
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "NombreCta"
         BoundColumn     =   "correl"
         Text            =   "Todos"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje respecto del valor base"
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
         Left            =   120
         TabIndex        =   52
         Top             =   5880
         Width           =   3150
      End
      Begin VB.Label lbl_Haber 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta Haber"
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
         Left            =   120
         TabIndex        =   44
         Top             =   4800
         Width           =   1245
      End
      Begin VB.Label lbl_enlace3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rubro de Ingresos Final"
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
         Left            =   120
         TabIndex        =   40
         Top             =   3120
         Width           =   2145
      End
      Begin VB.Label lbl_Debe 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta Debe"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   1185
      End
      Begin VB.Label lbl_enlace2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rubro de Ingresos Inicial"
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
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   2220
      End
      Begin VB.Label lbl_enlace1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Comprobante"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label lbl_descripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripcion del Asiento"
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2130
      End
      Begin VB.Label lbl_codigo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Correlativo"
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
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   4485
         TabIndex        =   10
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
      ScaleWidth      =   13080
      TabIndex        =   3
      Top             =   8040
      Width           =   13080
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   690
         TabIndex        =   8
         Top             =   0
         Width           =   3360
      End
   End
   Begin Crystal.CrystalReport cr01 
      Left            =   11520
      Top             =   7320
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
      Top             =   7320
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
      Left            =   2400
      Top             =   7320
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
      Left            =   4680
      Top             =   7320
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
      Left            =   6960
      Top             =   7320
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
      Left            =   9240
      Top             =   7320
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
End
Attribute VB_Name = "fw_Relacionador_Ingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents Ado_datos As Recordset
Dim rs_datos As New ADODB.Recordset
Attribute rs_datos.VB_VarHelpID = -1
Dim rs_datos1 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim var_cod, VAR_COD2 As String
Dim VAR_VAL As String
Dim VAR_SW As String

Dim Var_ing As Integer

Dim mvBookMark As Variant
Dim mbDataChanged As Boolean

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
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
        rs_datos.CancelBatch
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
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!correlativo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
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

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
'    If VAR_SW = "ADD" Then
'        Set rs_aux1 = New ADODB.Recordset
'        If rs_aux1.State = 1 Then rs_aux1.Close
'        SQL_FOR = "select * from fc_relacionador_ingresos where correlativo = '" & txt_codigo.Text & "'  "
'        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux1.RecordCount > 0 Then
'            var_cod = rs_aux1!correl_etapa + 1
'            rs_aux1!correl_etapa = Val(var_cod)
'            rs_aux1.Update
'        Else
'            var_cod = 1
'        End If
'        If var_cod < 10 Then
'            VAR_COD2 = Trim(dtc_codigo2.Text) + "-0" + Trim(var_cod)
'        Else
'            VAR_COD2 = Trim(dtc_codigo2.Text) + "-" + Trim(var_cod)
'        End If
'        rs_datos!correlativo = txt_codigo.Text
        rs_datos!descripcion = Txt_descripcion.Text
'       rs_datos!org_sigla = Txt_sigla.Text
        rs_datos!estado_codigo = "REG"  ' no cambia
        rs_datos!Codigo_Tipo = dtc_codigo1.Text
        rs_datos!rubro_codigo_I = dtc_codigo2.Text
        rs_datos!rubro_codigo_F = dtc_codigo3.Text
        rs_datos!Correl_Deb = dtc_Correl4.Text
        rs_datos!cta_deb = dtc_Cta_Deb.Text
        rs_datos!Subcta_deb1 = dtc_SubCta1_4.Text
        rs_datos!Subcta_deb2 = dtc_SubCta2_4.Text
        rs_datos!Correl_Cred = dtc_Correl5.Text
        rs_datos!cta_cred = dtc_Cta_Cred.Text
        rs_datos!Subcta_cred1 = dtc_SubCta1_5.Text
        rs_datos!Subcta_cred2 = dtc_SubCta2_5.Text
        
        'rs_datos!CORREL = 0
        'Guarda en el Padre, en el campo ctrl de correlativos
        'db.Execute "Update gc_direccion_administrativa Set correl_unidad = CAST('" & var_cod & "' AS INT) + 1 Where da_codigo = '" & dtc_codigo3.Text & "' "
  
        rs_datos!porcentaje = txt_porcentaje.Text
        
     rs_datos!fecha_registro = Date     ' no cambia
     rs_datos!usr_codigo = glusuario    ' no cambia
     rs_datos.UpdateBatch adAffectAll
    
     Call ABRIR_TABLA
     rs_datos.MoveLast
     mbDataChanged = False
      
      Fra_ABM.Enabled = False
      fraOpciones.Visible = True
      FraGrabarCancelar.Visible = False
      dg_datos.Enabled = True
      'txt_codigo.Enabled = True
      dtc_desc1.Enabled = True
      dtc_desc2.Enabled = True
      dtc_desc3.Enabled = True
      dtc_desc4.Enabled = True
      dtc_desc5.Enabled = True
   End If
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub valida_campos()
'  If txt_codigo.Text = "" Then
'    MsgBox "Debe registrar el " + lbl_codigo.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  If Txt_descripcion.Text = "" Then
    MsgBox "Debe registrar la " + lbl_descripcion.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_Correl4.Text = "" Then
    MsgBox "Debe registrar la " + lbl_Debe.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo1.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo2.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_codigo3.Text = "" Then
    MsgBox "Debe registrar: " + lbl_enlace3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If dtc_Correl5.Text = "" Then
    MsgBox "Debe registrar la " + lbl_Haber.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub BtnImprimir_Click()
  Dim iResult As Integer
  CR01.WindowShowPrintSetupBtn = True
  CR01.WindowShowRefreshBtn = True
  CR01.ReportFileName = App.Path & "\REPORTES\clasificadores\fr_Relacionador_Ingresos.rpt"
  iResult = CR01.PrintReport
  If iResult <> 0 Then
      MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbExclamation + vbOKOnly, "Error"
  End If
  CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "MOD"
'    txt_codigo.Enabled = False
    dtc_desc1.Enabled = True
    dtc_desc2.Enabled = True
    dtc_desc3.Enabled = True
    dtc_desc4.Enabled = True
    dtc_desc5.Enabled = True

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

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_Correl4_Click(Area As Integer)
dtc_desc4.BoundText = dtc_Correl4.BoundText
dtc_Cta_Deb.BoundText = dtc_Correl4.BoundText
dtc_SubCta1_4.BoundText = dtc_Correl4.BoundText
dtc_SubCta2_4.BoundText = dtc_Correl4.BoundText
End Sub

Private Sub dtc_Correl5_Click(Area As Integer)
dtc_desc5.BoundText = dtc_Correl5.BoundText
dtc_SubCta1_5.BoundText = dtc_Correl5.BoundText
dtc_SubCta2_5.BoundText = dtc_Correl5.BoundText
dtc_Cta_Cred.BoundText = dtc_Correl5.BoundText
End Sub

Private Sub dtc_Cta_Cred_Click(Area As Integer)
dtc_desc5.BoundText = dtc_Cta_Cred.BoundText
dtc_SubCta1_5.BoundText = dtc_Cta_Cred.BoundText
dtc_SubCta2_5.BoundText = dtc_Cta_Cred.BoundText
dtc_Correl5.BoundText = dtc_Cta_Cred.BoundText
End Sub

Private Sub dtc_Cta_Deb_Click(Area As Integer)
dtc_desc4.BoundText = dtc_Cta_Deb.BoundText
dtc_Correl4.BoundText = dtc_Cta_Deb.BoundText
dtc_SubCta1_4.BoundText = dtc_Cta_Deb.BoundText
dtc_SubCta2_4.BoundText = dtc_Cta_Deb.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText

    dtc_desc2.Enabled = True
End Sub
'Private Sub pnivel2(codigo2 As String)
'   Dim strConsultaF As String
'   strConsultaF = "select * from gc_proceso_nivel2 where proceso_codigo = '" & codigo2 & "'"
'
'   Set dtc_codigo2.RowSource = Nothing
'   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_codigo2.ReFill
'   dtc_codigo2.BoundText = Empty
'
'   Set dtc_desc2.RowSource = Nothing
'   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   dtc_desc2.ReFill
'   dtc_desc2.BoundText = Empty
'End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    'dtc_desc4.Enabled = True
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
'    'Var_ing = dtc_Correl4.Text
    dtc_SubCta2_4.BoundText = dtc_desc4.BoundText
    dtc_Cta_Deb.BoundText = dtc_desc4.BoundText
    dtc_SubCta1_4.BoundText = dtc_desc4.BoundText
    dtc_Correl4.BoundText = dtc_desc4.BoundText
'    Var_ing = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
dtc_SubCta2_5.BoundText = dtc_desc5.BoundText
dtc_Cta_Cred.BoundText = dtc_desc5.BoundText
dtc_SubCta1_5.BoundText = dtc_desc5.BoundText
dtc_Correl5.BoundText = dtc_desc5.BoundText

End Sub

Private Sub dtc_SubCta1_4_Click(Area As Integer)
dtc_desc4.BoundText = dtc_SubCta1_4.BoundText
dtc_Cta_Deb.BoundText = dtc_SubCta1_4.BoundText
dtc_Correl4.BoundText = dtc_SubCta1_4.BoundText
dtc_SubCta2_4.BoundText = dtc_SubCta1_4.BoundText
End Sub

Private Sub dtc_SubCta1_5_Click(Area As Integer)
dtc_desc5.BoundText = dtc_SubCta1_5.BoundText
dtc_Cta_Cred.BoundText = dtc_SubCta1_5.BoundText
dtc_SubCta2_5.BoundText = dtc_SubCta1_5.BoundText
dtc_Correl5.BoundText = dtc_SubCta1_5.BoundText

End Sub

Private Sub dtc_SubCta2_4_Click(Area As Integer)
dtc_desc4.BoundText = dtc_SubCta2_4.BoundText
dtc_Cta_Deb.BoundText = dtc_SubCta2_4.BoundText
dtc_SubCta1_4.BoundText = dtc_SubCta2_4.BoundText
dtc_Correl4.BoundText = dtc_SubCta2_4.BoundText
End Sub

Private Sub dtc_SubCta2_5_Click(Area As Integer)
dtc_desc5.BoundText = dtc_SubCta2_5.BoundText
dtc_SubCta1_5.BoundText = dtc_SubCta2_5.BoundText
dtc_Cta_Cred.BoundText = dtc_SubCta2_5.BoundText
dtc_Correl5.BoundText = dtc_SubCta2_5.BoundText
End Sub

Private Sub Form_Load()
    Call ABRIR_TABLAS_AUX
    Call ABRIR_TABLA
'   txt_codigo.Enabled = True
    mbDataChanged = False
    Fra_ABM.Enabled = False
    dg_datos.Enabled = True
    FraNavega.Caption = lbl_titulo.Caption
    lbl_titulo2.Caption = lbl_titulo.Caption
End Sub

Private Sub ABRIR_TABLA()
  Set rs_datos = New Recordset
  If rs_datos.State = 1 Then rs_datos.Close
  queryinicial = "select  * from fc_relacionador_ingresos  "
  rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
  Set Ado_datos.Recordset = rs_datos.DataSource
  Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLAS_AUX()
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_tipo_comprobante order by Denominacion_Tipo", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
    
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from fc_ingresos_rubro order by rubro_descripcion", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText
    
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from fc_ingresos_rubro order by rubro_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
     Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "Select * from cc_plan_cuentas where MOV = 'D' order by NombreCta", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_Correl4.BoundText
    dtc_Cta_Deb.BoundText = dtc_Correl4.BoundText
    dtc_SubCta1_4.BoundText = dtc_Correl4.BoundText
    dtc_SubCta2_4.BoundText = dtc_Correl4.BoundText
    
     Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from cc_plan_cuentas where MOV = 'D' order by NombreCta", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_Cta_Cred.BoundText
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
    Call ABRIR_TABLA
    If rs_datos.RecordCount > 0 Then rs_datos.MoveLast
    rs_datos.AddNew
    'lblStatus.Caption = "Agregar registro"
    Fra_ABM.Enabled = True
    fraOpciones.Visible = False
    FraGrabarCancelar.Visible = True
    dg_datos.Enabled = False
    VAR_SW = "ADD"
'    txt_codigo.Enabled = True
'    txt_codigo.SetFocus
    Txt_descripcion.SetFocus
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
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM ao_solicitud WHERE etapa_codigo = '" & Unidad & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Sub Txt_descripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

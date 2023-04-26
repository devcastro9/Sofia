VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_calendario_zonas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tecnico - Calendario por Zonas"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   14100
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   44
      Top             =   0
      Width           =   20280
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H0080FFFF&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   15120
         Picture         =   "tw_calendario_zonas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Picture         =   "tw_calendario_zonas.frx":0442
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   52
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1320
         Picture         =   "tw_calendario_zonas.frx":0C01
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   51
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
         Picture         =   "tw_calendario_zonas.frx":1516
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   50
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
         Picture         =   "tw_calendario_zonas.frx":1C62
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   49
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "tw_calendario_zonas.frx":2495
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   48
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
         Picture         =   "tw_calendario_zonas.frx":2C4A
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   47
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
         Picture         =   "tw_calendario_zonas.frx":3517
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   46
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H0080FFFF&
         Height          =   600
         Left            =   15960
         Picture         =   "tw_calendario_zonas.frx":3CD9
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
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
         TabIndex        =   54
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
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   20280
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6480
         Picture         =   "tw_calendario_zonas.frx":3EE3
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   42
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
         Picture         =   "tw_calendario_zonas.frx":47CF
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   41
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
         TabIndex        =   43
         Top             =   195
         Width           =   1005
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   16560
      TabIndex        =   38
      Top             =   1800
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483637
      Appearance      =   1
      StartOfWeek     =   119668738
      TitleBackColor  =   -2147483635
      TitleForeColor  =   -2147483637
      CurrentDate     =   42828
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CALENDARIO POR ZONA PILOTO - Detalle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7215
      Left            =   7080
      TabIndex        =   14
      Top             =   720
      Width           =   9405
      Begin VB.PictureBox fra_opciones_det 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   9225
         TabIndex        =   55
         Top             =   240
         Width           =   9225
         Begin VB.PictureBox BtnCancelarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5880
            Picture         =   "tw_calendario_zonas.frx":4FA5
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnGrabarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4440
            Picture         =   "tw_calendario_zonas.frx":5891
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   59
            Top             =   0
            Visible         =   0   'False
            Width           =   1280
         End
         Begin VB.PictureBox BtnAddDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            Picture         =   "tw_calendario_zonas.frx":6067
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1545
            Picture         =   "tw_calendario_zonas.frx":6826
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   57
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3120
            Picture         =   "tw_calendario_zonas.frx":713B
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   56
            Top             =   0
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Height          =   6135
         Left            =   75
         TabIndex        =   15
         Top             =   960
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   10821
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "fmes_plan"
            Caption         =   "Mes"
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
            DataField       =   "dia_correl"
            Caption         =   "#.Dia"
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
            DataField       =   "dia_fecha"
            Caption         =   "Fecha"
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
            DataField       =   "dia_nombre"
            Caption         =   "Nombre.Dia"
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
            DataField       =   "horario_codigo"
            Caption         =   "Horario"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Hora.Inicio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "hora_salida"
            Caption         =   "Hora.Fin"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nro_total_horas"
            Caption         =   "Nro.Horas"
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
            DataField       =   "bien_codigo"
            Caption         =   "Codigo.Equipo"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tec.Mantenimiento"
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
            DataField       =   "beneficiario_codigo_resp2"
            Caption         =   "Tec.Emergencias"
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
         BeginProperty Column11 
            DataField       =   "estado_activo"
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
         BeginProperty Column12 
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2880
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Registro Cabecera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3735
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Width           =   7020
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   37
         Top             =   1270
         Width           =   270
      End
      Begin VB.ComboBox Cmb_Mes_ant 
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Text            =   "ENERO"
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_gestion_Ant 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Text            =   "2016"
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Txt_campo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "observaciones"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3360
         Width           =   6645
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Top             =   2005
         Width           =   270
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DataField       =   "fmes_nro_hrs_habiles"
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
         ForeColor       =   &H00FFFF80&
         Height          =   290
         Left            =   285
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   8
         Top             =   3675
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox Txt_estado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5535
         TabIndex        =   6
         Top             =   2700
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_calendario_zonas.frx":7887
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   9
         Top             =   880
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_calendario_zonas.frx":78A0
         DataField       =   "unidad_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "unidad_codigo_tec"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_calendario_zonas.frx":78B9
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   2000
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
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
         DataField       =   "fmes_fecha_registro"
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
         Left            =   5280
         TabIndex        =   0
         Top             =   3645
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   119668739
         CurrentDate     =   41678
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_calendario_zonas.frx":78D2
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1260
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ForeColor       =   0
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
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
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_calendario_zonas.frx":78EB
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   2685
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_calendario_zonas.frx":7904
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5880
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "0"
      End
      Begin VB.Label Cmb_Mes 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1320
         TabIndex        =   36
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label Cmb_gestion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "ges_gestion"
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
         Height          =   300
         Left            =   240
         TabIndex        =   35
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3135
         Width           =   1275
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsable Zona"
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
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   2460
         Width           =   1605
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad Ejecutora"
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
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Horas X Dia"
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
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   29
         Top             =   3720
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Height          =   195
         Index           =   6
         Left            =   6135
         TabIndex        =   28
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Elaboracion"
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
         Height          =   195
         Left            =   3600
         TabIndex        =   24
         Top             =   3765
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_nro_dias"
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
         Height          =   300
         Left            =   5535
         TabIndex        =   23
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total.Dias X Mes"
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
         Height          =   195
         Index           =   3
         Left            =   5475
         TabIndex        =   22
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Correlativo Crono."
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
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   21
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
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
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   20
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion"
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
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   280
         Width           =   660
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto (Ruta)"
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
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label lbl_texto1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "0"
         DataField       =   "fmes_correl"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_plan"
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
         Height          =   300
         Left            =   3960
         TabIndex        =   16
         Top             =   525
         Width           =   1335
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   7020
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   4920
         TabIndex        =   4
         Top             =   3075
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   3
         Top             =   3075
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3000
         Width           =   6825
         _ExtentX        =   12039
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
         BackColor       =   12632256
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
         Height          =   2730
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   4815
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
         ColumnCount     =   7
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
            DataField       =   "fmes_correl"
            Caption         =   "Mes"
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
            DataField       =   "observaciones"
            Caption         =   "Zona.Piloto"
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
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Responsable"
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
            DataField       =   "fmes_nro_dias"
            Caption         =   "Total.Dias"
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
               Alignment       =   2
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   4320
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
      Left            =   10800
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Calendario solo para Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   16520
      TabIndex        =   39
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "tw_calendario_zonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
'Dim swnuevo As String
Dim imag2 As Long

Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES As String
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

Dim mes2 As String

Dim mvBookMark, marca1 As Variant
Dim mbDataChanged As Boolean

Private Sub Ado_datos_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
     '<-- Inicio                Identificación del Cliente                Fin -->
   If (Not Ado_datos.Recordset.BOF) And (Not Ado_datos.Recordset.EOF) Then
     If VAR_SW = "NO" Then
'        Select Case dtc_codigo2.Text
'            Case "1"
'            Case "2"
'            Case "3"
'                Call ABRIR_TABLA_DET3
'            Case "4"
'
'        End Select
        If Ado_datos.Recordset.RecordCount > 0 Then
            Call ABRIR_TABLA_DET
            If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                'Cmb_Mes = UCase(MonthName(Val(lbl_texto1.Caption)))
                Cmb_Mes = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
            End If
            'mes2 = MonthName(Month(DTPfecha1.Value))
        End If
        VAR_SW = ""
    Else
        VAR_SW = ""
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
    End If
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle1.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle1.Recordset!estado_activo = "ANL"
        Ado_detalle1.Recordset!observaciones = "HORARIO NO LABORABLE"
        Ado_detalle1.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue Aprobado (Estado=APR) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAñadir_Click()
    On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    'If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
'        FrmABMDet.Visible = False
        dg_det1.Enabled = False
        swnuevo = 1
    '    BtnVer.Visible = True
    'Else
    '  MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    'End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
'  On Error GoTo UpdateErr
'   Set rs_aux2 = New ADODB.Recordset
'   rs_aux2.Open "Select * from ao_solicitud_costos where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "'  and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'   If rs_aux2.RecordCount > 0 Then
'        VAR_CONT2 = rs_aux2.RecordCount
'   End If
'   'If rs_datos!estado_codigo = "REG" And Ado_datos.Recordset!correl_edificacion > 0 Then
'   If rs_datos!estado_codigo = "REG" And VAR_CONT2 > 0 Then
'      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'
''        Select Case dtc_codigo2.Text
''            Case "1"
''            Case "2"
''            Case "3"
'                Set rs_aux1 = New ADODB.Recordset
'                'SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  and edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "'  "
'                SQL_FOR = "select * from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "    "
'                rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'                If rs_aux1.RecordCount > 0 Then
'                    MsgBox "Una Cotización anterior ya fue Aprobada, el Registro Actual se adicionará al que fue aprobado anteriormente ..."
'                    '    var_cod = 0
'                    '    Exit Sub
'                    rs_aux1!venta_monto_total_bs = rs_aux1!venta_monto_total_bs + Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = rs_aux1!venta_monto_total_dol + Ado_datos.Recordset!cotiza_precio_total_dol
'                Else
'                    'CREA VENTA CABECERA
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    'rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'                    rs_aux2.Open "Select max(venta_codigo) as Codigo from ao_ventas_cabecera    ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'                    End If
'                    Set rs_aux2 = New ADODB.Recordset
'                    If rs_aux2.State = 1 Then rs_aux2.Close
'                    rs_aux2.Open "Select beneficiario_codigo as Codigo from ao_solicitud where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'                    If Not rs_aux2.EOF Then
'                        VAR_AUX = rs_aux2!Codigo
'                    End If
'                    rs_aux1.AddNew
'                    'var_cod = rs_aux1.RecordCount + 1
'                    rs_aux1!ges_gestion = Year(Date)
'                    rs_aux1!unidad_codigo = Ado_datos.Recordset!unidad_codigo
'                    rs_aux1!solicitud_codigo = Ado_datos.Recordset!solicitud_codigo
'                    rs_aux1!edif_codigo = Ado_datos.Recordset!edif_codigo
'                    rs_aux1!venta_codigo = var_cod
'                    rs_aux1!beneficiario_codigo = VAR_AUX
'                    rs_aux1!venta_monto_total_bs = Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_monto_total_dol = Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!venta_monto_cobrado_bs = 0
'                    rs_aux1!venta_monto_cobrado_dol = 0
'                    rs_aux1!venta_saldo_p_cobrar_bs = Ado_datos.Recordset!cotiza_precio_total_bs
'                    rs_aux1!venta_saldo_p_cobrar_dol = Ado_datos.Recordset!cotiza_precio_total_dol
'                    rs_aux1!unidad_codigo_ant = Ado_datos.Recordset!unidad_codigo_ant
'                    rs_aux1!estado_codigo = "REG"
'                    rs_aux1!fecha_registro = Date
'                    rs_aux1!usr_codigo = glusuario
'                    rs_aux1.Update
''                    db.Execute "Update ao_solicitud Set correl_calculo = " & var_cod & " Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
'                End If
'                'db.Execute "Update ao_solicitud_calculo_trafico Set estado_codigo = 'APR' Where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "  "
''            Case "4"
''        End Select
'        'GRABA VENTA DETALLE
'        If var_cod = "" Then
'            var_cod = rs_aux1!venta_codigo
'        End If
'        Set rs_aux3 = New ADODB.Recordset
'        If rs_aux3.State = 1 Then rs_aux3.Close
'        'rs_aux3.Open "Select * from ao_ventas_detalle where unidad_codigo = '" & Ado_datos.Recordset!unidad_codigo & "' and solicitud_codigo = " & Ado_datos.Recordset!solicitud_codigo & "   ", db, adOpenStatic
'        rs_aux3.Open "Select * from ao_ventas_detalle where venta_codigo = " & var_cod & " and ges_gestion = '" & Year(Date) & "'   ", db, adOpenKeyset, adLockOptimistic
'        'If rs_aux3.RecordCount > 0 Then
'            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
'        'Else
'            VAR_AUX = rs_aux3.RecordCount + 1
'            rs_aux3.AddNew
'            rs_aux3!ges_gestion = Year(Date)
'            rs_aux3!venta_codigo = var_cod
'            rs_aux3!venta_codigo_det = VAR_AUX
'            rs_aux3!bien_codigo = Ado_datos.Recordset!bien_codigo
'            rs_aux3!venta_det_cantidad = Ado_datos.Recordset!cotiza_cantidad
'            rs_aux3!venta_precio_unitario_bs = 0
'            rs_aux3!venta_descuento_bs = 0
'            rs_aux3!venta_precio_total_bs = 0
'            rs_aux3!venta_precio_unitario_dol = 0
'            rs_aux3!venta_descuento_dol = 0
'            rs_aux3!venta_precio_total_dol = 0
''            rs_aux3!concepto_venta = dtc_desc21.Text + " - " + Ado_datos.Recordset!bien_codigo
'            'ok
'            rs_aux3!grupo_codigo = "40000"
'            rs_aux3!subgrupo_codigo = "43000"
'            rs_aux3!par_codigo = "43340"
'            'ok
'            rs_aux3!tipo_descuento = 0
'            rs_aux3!almacen_codigo = 0
'            rs_aux3!modelo_codigo1 = Ado_datos.Recordset!modelo_codigo
'            rs_aux3!modelo_codigo_h = Ado_datos.Recordset!modelo_codigo_h
'            rs_aux3!modelo_codigo_x = Ado_datos.Recordset!modelo_codigo_x
'            rs_aux3!modelo_elegido = "N"
'            rs_aux3!modelo_elegido_h = "N"
'            rs_aux3!modelo_elegido_x = "N"
'            'rs_aux3!estado_codigo = "REG"
'            rs_aux3!fecha_registro = Date
'            rs_aux3!usr_codigo = glusuario
'            rs_aux3.Update
'        'End If
'        'INI GRABA ALMACEN DETALLE (EN LA ENTREGA EN OBRA)
''        Set rs_aux4 = New ADODB.Recordset
''        If rs_aux4.State = 1 Then rs_aux4.Close
''        rs_aux4.Open "Select * from ao_almacen_detalle where almacen_codigo = 0 and bien_codigo = '" & Ado_datos.Recordset!bien_codigo & "'   ", db, adOpenKeyset, adLockOptimistic
''        If rs_aux4.RecordCount = 0 Then
''            'var_cod = IIf(IsNull(rs_aux2!Codigo), 1, rs_aux2!Codigo + 1)
''            rs_aux4.AddNew
''            rs_aux4!almacen_codigo = 0
''            rs_aux4!bien_codigo = Ado_datos.Recordset!bien_codigo
''            rs_aux4!grupo_codigo = "40000"
''            rs_aux4!subgrupo_codigo = "43000"
''            rs_aux4!par_codigo = "43340"
''            rs_aux4!stock_ingreso = 1
''            rs_aux4!stock_salida = 0
''            rs_aux4!stock_actual = 1
''            rs_aux4!estado_codigo = "REG"
''            rs_aux4!usr_codigo = GlUsuario
''            rs_aux4!fecha_registro = Date
''            rs_aux4.Update
''        End If
'        'R-222 "COTIZACION DE EQUIPOS PARA EL CLIENTE"
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            rs_datos!doc_numero = rs_aux2!correl_doc
'            'Txt_campo1.Caption = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        'rs_datos!doc_numero = Txt_campo1.Caption
'        'REVISAR !!! JQA 2014_07_08
'        'VAR_ARCH = RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
'        VAR_ARCH = "COM_" + RTrim(RTrim(rs_datos!doc_codigo) + "-") + LTrim(Str(rs_datos!doc_numero))
'        rs_datos!archivo_respaldo = VAR_ARCH + ".PDF"
'        rs_datos!archivo_respaldo_cargado = "N"
'        'R-224 "PROPUESTA DE COTIZACION DE EQUIPOS PARA EL CLIENTE"
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & Ado_datos.Recordset!doc_codigo2 & "'  "
'        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'            rs_aux2!correl_doc = rs_aux2!correl_doc + 1
'            rs_datos!doc_numero2 = rs_aux2!correl_doc
'            rs_aux2.Update
'        End If
'        VAR_ARCH2 = "COM_" + RTrim(RTrim(rs_datos!doc_codigo2) + "-") + LTrim(Str(rs_datos!doc_numero2))
'        rs_datos!archivo_respaldo2 = VAR_ARCH2 + ".PDF"
'        rs_datos!archivo_respaldo_cargado2 = "N"
'
'        rs_datos!estado_codigo = "APR"
'        rs_datos!fecha_registro = Date
'        rs_datos!usr_codigo = glusuario
'        rs_datos.UpdateBatch adAffectAll
'      End If
'   Else
'       MsgBox "No se puede APROBAR un registro Anulado o Aprobado o que no tiene detalle ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description

End Sub

Private Sub BtnBuscar_Click()
    If Ado_datos.Recordset.RecordCount > 0 Then
        OptFilGral1.Visible = True
        OptFilGral2.Visible = True
'        If Ado_datos.Recordset!estado_codigo = "REG" Then
'            Call OptFilGral1_Click
'        Else
'            Call OptFilGral2_Click
'        End If
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_datos
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_datos
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existen registros. ", vbExclamation, "Atención!"
      OptFilGral1.Visible = True
      OptFilGral2.Visible = True
    End If
End Sub

Private Sub BtnCancelar_Click()
  On Error Resume Next
   sino = MsgBox("Está Seguro de CANCELAR la operación ? ", vbYesNo + vbQuestion, "Atención")
   If sino = vbYes Then
'        rs_datos.CancelUpdate
'        Call ABRIR_TABLA
        rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Enabled = False
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
'        FrmABMDet.Visible = True
        dg_det1.Enabled = True
        swnuevo = 0
    End If

End Sub



Private Sub btnEliminar_Click()
'  On Error GoTo UpdateErr
'   If ExisteReg(Ado_datos.Recordset!zpiloto_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
'   If rs_datos!estado_codigo = "APR" Then
'      sino = MsgBox("Está Seguro de ANULAR el Registro? Este ya no podrá ser utilizado...", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'         rs_datos!estado_codigo = "ANL"
'         rs_datos!fecha_registro = Date
'         rs_datos!usr_codigo = glusuario
'         rs_datos.UpdateBatch adAffectAll
'      End If
'   Else
'      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Anulado (ANL) ...", vbExclamation, "Validación de Registro"
'   End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo AddErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
    If swnuevo = 1 Then
'        db.Execute "Update to_cronograma_mensual Set estado_codigo = 'REG' Where ges_gestion = '" & Cmb_gestion & "' AND fmes_correl = " & mes2 & " AND zpiloto_codigo =" & dtc_codigo3.Text & "  "
'
'
'
'''------EMPIEZA EL CODIOGO PARA CREAR EL CALENDARIO
'
'Dim dia2 As String
'Dim fechita As Date
''Dim mes2 As String
''dia = WeekdayName(Weekday(Date))
''MsgBox dia
'Dim i As String
'i = Year(Date)
'Dim rsCalendar As New ADODB.Recordset
'   Set rsCalendar = New ADODB.Recordset
'   rsCalendar.Open "select * from to_cronograma_mensual WHERE ges_gestion = '" & i & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
''   Set AdoPais.Recordset = rsCalendar
'  If rsCalendar.RecordCount > 0 Then
'    MsgBox "Ya fue creado el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
'  Else
'    fechita = CDate("01/01/" + i)
'    While fechita >= CDate("01/01/" + i) And fechita <= CDate("31/12/" + i)
'       rsCalendar.AddNew
'       dia2 = WeekdayName(Weekday(fechita))
'       rsCalendar!fmes_fecha_registro = fechita   'Format(fechita, "dd/mm/aaaa")
'       rsCalendar!ges_gestion = Year(fechita)
'       mes2 = MonthName(Month(fechita))
'       rsCalendar!mes = UCase(mes2)
''       Select Case mes2
''            Case 2:                     Dias_Del_Mes = IIf(saltarYear(fecha), 29, 28)
''            Case 1, 3, 5, 7, 8, 10, 12: Dias_Del_Mes = 31
''            Case 4, 6, 9, 11:           Dias_Del_Mes = 30
''       End Select
'
'       rsCalendar!dia = UCase(dia2)
'       If dia2 = "sábado" Or dia2 = "domingo" Then
'          rsCalendar!tipo = "F"
'          rsCalendar!descripcion = "DIA FERIADO - NO LABORAL"
'       Else
'          rsCalendar!tipo = "L"
'          rsCalendar!descripcion = "DIA LABORAL"
'       End If
'       rsCalendar.Update
'       fechita = fechita + 1
'    Wend
'    MsgBox "Se ha creado con éxito el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
'  End If
'  Call ABRIR_TABLA
'
'
'
''    Else
''     rs_datos!fmes_fecha_registro = DTPfecha1.Value
''     rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
''     rs_datos!observaciones = Txt_campo2.Text
''
''     rs_datos!fecha_registro = Date     'no cambia
''     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
''     rs_datos.Update    'Batch 'adAffectAll
'    End If
    
    Dim dia2 As String
Dim fechita As Date
'Dim mes2 As String
'dia = WeekdayName(Weekday(Date))
'MsgBox dia
Dim i As String
i = Year(Date)
Dim rsCalendar As New ADODB.Recordset
   Set rsCalendar = New ADODB.Recordset
   rsCalendar.Open "select * from to_cronograma_mensual WHERE ges_gestion = '" & i & "' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'   Set AdoPais.Recordset = rsCalendar
'rsCalendar.Open "select * from to_cronograma_diario WHERE dia_fecha = '" & fechita & "' " ,and dia_fecha = '" & i & "' "
  If rsCalendar.RecordCount > 0 Then
    MsgBox "Ya fue creado el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
  Else
    fechita = CDate("01/01/" + i)
    While fechita >= CDate("01/01/" + i) And fechita <= CDate("31/12/" + i)
       rsCalendar.AddNew
       dia2 = WeekdayName(Weekday(fechita))
       rsCalendar!dia_fecha = fechita   'Format(fechita, "dd/mm/aaaa")
       rsCalendar!ges_gestion = Year(fechita)
       mes2 = MonthName(Month(fechita))
       rsCalendar!mes = UCase(mes2)
'       Select Case mes2
'            Case 2:                     Dias_Del_Mes = IIf(saltarYear(fecha), 29, 28)
'            Case 1, 3, 5, 7, 8, 10, 12: Dias_Del_Mes = 31
'            Case 4, 6, 9, 11:           Dias_Del_Mes = 30
'       End Select

       rsCalendar!dia_nombre = UCase(dia2)
       If dia2 = "sábado" Or dia2 = "domingo" Then
          rsCalendar!hora_registro = "F"
          rsCalendar!descripcion = "DIA FERIADO - NO LABORAL"
       Else
          rsCalendar!hora_registro = "L"
          rsCalendar!descripcion = "DIA LABORAL"
       End If
       rsCalendar.Update
       fechita = fechita + 1
    Wend
    MsgBox "Se ha creado con éxito el calendario de la Gestión: " + i, vbInformation + vbOKOnly, "Atención"
  End If
    
     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'   "

    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Call OptFilGral1_Click
    Else
        Call OptFilGral2_Click
    End If
    
       'rs_datos.MoveFirst
'     mbDataChanged = False
     Fra_datos.Enabled = False
     fraOpciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
'     FrmABMDet.Visible = True
     dg_det1.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     swnuevo = 0
  End If
  End If
  

  
  
 
  
  
  
  Exit Sub
AddErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'Valida compos para editables
'  If (dtc_codigo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo3.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (dtc_codigo4 = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
'  If (Txt_campo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
  
End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R-302_cronograma_mensual.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
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
      'Cmb_Mes.Text = "ENERO"
      CR01.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
      'CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
      
'    cr01.StoredProcParam(0) = "2015"    'Me.Ado_datos.Recordset!ges_gestion
'    cr01.StoredProcParam(1) = "DNMAN"   'Me.Ado_datos.Recordset!unidad_codigo_tec
'    cr01.StoredProcParam(2) = 0     'Me.Ado_datos.Recordset!zpiloto_codigo
'    cr01.StoredProcParam(3) = 1     'Me.Ado_datos.Recordset!fmes_correl
    
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
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
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
    If Ado_detalle1.Recordset("estado_activo") = "ANL" Then             'And Ado_detalle1.Recordset("estado_codigo") = "REG"
      sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle1.Recordset!estado_activo = "REG"
        Ado_detalle1.Recordset!observaciones = "HORARIO LABORABLE"
        Ado_detalle1.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModificar_Click()
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
         fraOpciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        swnuevo = 2
    '    BtnVer.Visible = True
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
    
End Sub


Private Sub Cmb_Mes_Change()
    Select Case Cmb_Mes
        Case "ENERO"
            mes2 = "1"
        Case "FEBRERO"
            mes2 = "2"
        Case "MARZO"
            mes2 = "3"
        Case "ABRIL"
            mes2 = "4"
        Case "MAYO"
            mes2 = "5"
        Case "JUNIO"
            mes2 = "6"
        Case "JULIO"
            mes2 = "7"
        Case "AGOSTO"
            mes2 = "8"
        Case "SEPTIEMBRE"
            mes2 = "9"
        Case "OCTUBRE"
            mes2 = "10"
        Case "NOVIEMBRE"
            mes2 = "11"
        Case "DICIEMBRE"
            mes2 = "12"

    End Select
End Sub

Private Sub dg_datos_Click()
    VAR_SW = "NO"
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'Fra_Gestion.Visible = True
    'VAR_GES = Cmb_gestion.Text
    'parametro = Aux
    Call ABRIR_TABLAS_AUX
    Call OptFilGral1_Click
    
    Fra_datos.Enabled = False
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
        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_unidad_ejecutora
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    'rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText
        
    'tc_zonas_piloto
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    'rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = 'DNMAN' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
'    Call pnivel1(dtc_codigo1.BoundText)
'    dtc_desc10.Enabled = True
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (registros NO aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close   '
    'queryinicial = "select * From tv_cronograma_mensual_zona WHERE estado_codigo = 'REG' "      'AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '2015'
    queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
    
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "Select * from tv_cronograma_mensual_zona where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
    queryinicial = "Select * from to_cronograma_mensual WHERE estado_codigo <> 'ERR'  "         'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "'
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub ABRIR_TABLA()
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    queryinicial = "Select * from ao_solicitud_cotiza_venta where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
        
'    dtc_desc31.BoundText = dtc_codigo31.BoundText
'    dtc_desc32.BoundText = dtc_codigo31.BoundText
'    dtc_desc33.BoundText = dtc_codigo31.BoundText
'    dtc_desc34.BoundText = dtc_codigo31.BoundText
'
'    dtc_desc41.BoundText = dtc_codigo41.BoundText
'    dtc_desc42.BoundText = dtc_codigo41.BoundText
'    dtc_desc43.BoundText = dtc_codigo41.BoundText
'    dtc_desc44.BoundText = dtc_codigo41.BoundText
'
'    dtc_desc51.BoundText = dtc_codigo51.BoundText
'    dtc_desc52.BoundText = dtc_codigo51.BoundText
'    dtc_desc53.BoundText = dtc_codigo51.BoundText
'    dtc_desc54.BoundText = dtc_codigo51.BoundText
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

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub ABRIR_TABLA_DET()
    Dim rs_datos1 As New ADODB.Recordset
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from to_cronograma_diario where fmes_plan = " & Ado_datos.Recordset!fmes_plan & "  ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
'    If rs_det1.RecordCount > 0 Then
'
'    Else
'        Set Ado_detalle1.Recordset = rsNada.DataSource
'    End If
End Sub

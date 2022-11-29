VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_cronograma_certifica_inst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Comercial - Ejecucion del Servicio Instalaciones"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10620
   Icon            =   "tw_cronograma_certifica_inst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   10620
   WindowState     =   2  'Maximized
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Ejecución del Servicio"
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
      Height          =   3840
      Left            =   4920
      TabIndex        =   52
      Top             =   3120
      Visible         =   0   'False
      Width           =   7140
      Begin VB.TextBox txt_correl_carta 
         DataField       =   "doc_numero_carta"
         DataSource      =   "Ado_detalle2"
         Height          =   285
         Left            =   3960
         TabIndex        =   69
         Text            =   "0"
         Top             =   1320
         Width           =   1440
      End
      Begin VB.ComboBox Cmb_carta 
         DataField       =   "carta"
         DataSource      =   "Ado_detalle2"
         Height          =   315
         ItemData        =   "tw_cronograma_certifica_inst.frx":0A02
         Left            =   4560
         List            =   "tw_cronograma_certifica_inst.frx":0A0C
         TabIndex        =   68
         Text            =   "NO"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt_cm 
         DataField       =   "doc_numero"
         DataSource      =   "Ado_detalle2"
         Height          =   285
         Left            =   5520
         TabIndex        =   64
         Text            =   "0"
         Top             =   600
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txt_hdm 
         DataField       =   "nro_fojas"
         DataSource      =   "Ado_detalle2"
         Height          =   285
         Left            =   5520
         TabIndex        =   63
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txt_obs 
         DataField       =   "observaciones"
         DataSource      =   "Ado_datos"
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Text            =   "tw_cronograma_certifica_inst.frx":0A18
         Top             =   2040
         Width           =   6600
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6600
         TabIndex        =   55
         Top             =   3090
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton BtnGraba3 
         BackColor       =   &H80000015&
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   2040
         Picture         =   "tw_cronograma_certifica_inst.frx":0A1A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Aprueba Registro"
         Top             =   3000
         Width           =   1125
      End
      Begin VB.CommandButton BtnCancelar3 
         BackColor       =   &H80000015&
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   3840
         Picture         =   "tw_cronograma_certifica_inst.frx":0C24
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3000
         Width           =   1125
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Height          =   315
         Left            =   240
         TabIndex        =   56
         Top             =   3075
         Visible         =   0   'False
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Height          =   315
         Left            =   5880
         TabIndex        =   57
         Top             =   3075
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPEjecucion 
         DataField       =   "fecha_conformidad"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Ado_detalle2"
         Height          =   315
         Left            =   3720
         TabIndex        =   59
         Top             =   360
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   128581635
         CurrentDate     =   44797
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Nota E.D.E."
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
         Left            =   1080
         TabIndex        =   67
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota de Entrega Definitiva de Equipo?"
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
         Left            =   1080
         TabIndex        =   66
         Top             =   840
         Width           =   3300
      End
      Begin VB.Label Label6 
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
         TabIndex        =   65
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. NOTA"
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
         Left            =   5640
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. NOTA"
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
         Left            =   5520
         TabIndex        =   60
         Top             =   1080
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ejecución de la Tarea"
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
         Left            =   1080
         TabIndex        =   58
         Top             =   360
         Width           =   2460
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Registro de Certificados de una Zona Piloto"
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
      Height          =   7455
      Left            =   1680
      TabIndex        =   81
      Top             =   1200
      Visible         =   0   'False
      Width           =   12525
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   12420
         TabIndex        =   82
         Top             =   240
         Width           =   12420
         Begin VB.PictureBox BtnCancelarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1320
            Picture         =   "tw_cronograma_certifica_inst.frx":0E2E
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   89
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
            Left            =   0
            Picture         =   "tw_cronograma_certifica_inst.frx":171A
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   1280
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2640
            Picture         =   "tw_cronograma_certifica_inst.frx":1EF0
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   83
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "tw_cronograma_certifica_inst.frx":26A5
            Height          =   315
            Left            =   5880
            TabIndex        =   84
            Top             =   120
            Visible         =   0   'False
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "tw_cronograma_certifica_inst.frx":26C0
            Height          =   315
            Left            =   10920
            TabIndex        =   85
            Top             =   120
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Buscar Edificio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   10
            Left            =   4320
            TabIndex        =   86
            Top             =   165
            Visible         =   0   'False
            Width           =   1770
         End
      End
      Begin MSDataGridLib.DataGrid dg_det3 
         Bindings        =   "tw_cronograma_certifica_inst.frx":26DB
         Height          =   6375
         Left            =   75
         TabIndex        =   87
         Top             =   960
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   11245
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
            Caption         =   "Fecha.Crono"
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
            DataField       =   "edif_descripcion"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "beneficiario_codigo_resp"
            Caption         =   "Tec.Mantenim."
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
            DataField       =   "fecha_conformidad"
            Caption         =   "Fecha.Ejecutado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "nro_fojas"
            Caption         =   "Nro.HDM"
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.C.M."
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
            DataField       =   "carta"
            Caption         =   "Carta?"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "NO"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "doc_numero_carta"
            Caption         =   "#Carta"
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
         BeginProperty Column14 
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
         BeginProperty Column15 
            DataField       =   "nro_total_horas"
            Caption         =   "#.Horas"
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
         BeginProperty Column16 
            DataField       =   "hora_registro"
            Caption         =   "Hora_registro"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   3135.118
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column07 
               Locked          =   -1  'True
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   -1  'True
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column14 
               Alignment       =   2
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column16 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   3120
      Left            =   4920
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   6540
      Begin VB.OptionButton OptFilGral3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2021"
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
         Left            =   3480
         TabIndex        =   74
         Top             =   2835
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral0 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2019"
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
         TabIndex        =   73
         Top             =   2835
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2022"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   2835
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2020"
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
         Left            =   2280
         TabIndex        =   3
         Top             =   2835
         Width           =   855
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2760
         Width           =   6345
         _ExtentX        =   11192
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
         Bindings        =   "tw_cronograma_certifica_inst.frx":26F6
         Height          =   2490
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   4392
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
            Caption         =   "Nro.Dias"
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
            DataField       =   "estado_certifica"
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
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3135.118
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
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
      TabIndex        =   35
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   9360
         Picture         =   "tw_cronograma_certifica_inst.frx":270E
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   80
         ToolTipText     =   "Nota de Entrega Definitiva de Equipo"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnImprimir4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10680
         Picture         =   "tw_cronograma_certifica_inst.frx":2FDB
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   79
         ToolTipText     =   "Ejecución de Cronograma (Certificados)"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnImprimir3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6720
         Picture         =   "tw_cronograma_certifica_inst.frx":38A8
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   78
         ToolTipText     =   "Ejecución Cronograma y Cobrador"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnImprimir1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8040
         Picture         =   "tw_cronograma_certifica_inst.frx":4175
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   77
         ToolTipText     =   "Cronograma con Insumos"
         Top             =   0
         Width           =   1400
      End
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17760
         Picture         =   "tw_cronograma_certifica_inst.frx":4A42
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   76
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnModDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_certifica_inst.frx":5204
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   75
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5400
         Picture         =   "tw_cronograma_certifica_inst.frx":5B19
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4080
         Picture         =   "tw_cronograma_certifica_inst.frx":63E6
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   42
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7320
         Picture         =   "tw_cronograma_certifica_inst.frx":6B9B
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   10440
         Picture         =   "tw_cronograma_certifica_inst.frx":73CE
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   10
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9000
         Picture         =   "tw_cronograma_certifica_inst.frx":7810
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl_titulo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CRONOGRAMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   270
         Left            =   16035
         TabIndex        =   38
         Top             =   195
         Width           =   1695
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
      TabIndex        =   33
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
         Picture         =   "tw_cronograma_certifica_inst.frx":7A1A
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   40
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
         Picture         =   "tw_cronograma_certifica_inst.frx":81F0
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   39
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
         Left            =   12735
         TabIndex        =   34
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   8295
      Left            =   0
      TabIndex        =   31
      Top             =   360
      Width           =   19245
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   60
         ScaleHeight     =   420
         ScaleWidth      =   19140
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   19140
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "tw_cronograma_certifica_inst.frx":8ADC
            Height          =   315
            Left            =   3480
            TabIndex        =   70
            Top             =   0
            Visible         =   0   'False
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "tw_cronograma_certifica_inst.frx":8AF7
            Height          =   315
            Left            =   8520
            TabIndex        =   71
            Top             =   0
            Visible         =   0   'False
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Appearance      =   0
            BackColor       =   12632256
            ListField       =   "edif_codigo"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Elije una o más columnas para Filtrar ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   72
            Top             =   45
            Visible         =   0   'False
            Width           =   3360
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B12
         Height          =   7815
         Left            =   75
         TabIndex        =   32
         Top             =   360
         Width           =   19110
         _ExtentX        =   33708
         _ExtentY        =   13785
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
         Caption         =   "EJECUCION DEL SERVICIO DE INSTALACIONES"
         ColumnCount     =   20
         BeginProperty Column00 
            DataField       =   "fmes_plan"
            Caption         =   "#Crono"
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
            DataField       =   "ges_gestion"
            Caption         =   "Gestion"
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
            DataField       =   "zpiloto_codigo"
            Caption         =   "#Grupo"
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
            DataField       =   "zpiloto_descripcion"
            Caption         =   "Grupo.Descripcion"
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
         BeginProperty Column04 
            DataField       =   "edif_codigo"
            Caption         =   "Codigo.Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Edificio_Descripcion"
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
            DataField       =   "horario_codigo"
            Caption         =   "#Tarea"
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
            DataField       =   "TareaDescripcion"
            Caption         =   "Tarea.Descripcion"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Fecha.Inicio"
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
            DataField       =   "hora_salida"
            Caption         =   "Fecha.Fin"
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
         BeginProperty Column11 
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
         BeginProperty Column12 
            DataField       =   "nro_total_horas"
            Caption         =   "#.Horas"
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
         BeginProperty Column13 
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
         BeginProperty Column14 
            DataField       =   "fecha_conformidad"
            Caption         =   "Fecha.Ejecutado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "estado_activo"
            Caption         =   "Ejecutado"
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
            DataField       =   "doc_numero"
            Caption         =   "Nro.Nota"
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
         BeginProperty Column17 
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
         BeginProperty Column18 
            DataField       =   "beneficiario_denominacion"
            Caption         =   "Tec.Instalacion"
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
            DataField       =   "hora_registro"
            Caption         =   "Hora_registro"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   1980.284
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               ColumnWidth     =   3135.118
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2894.74
            EndProperty
            BeginProperty Column08 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column17 
               Locked          =   -1  'True
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column18 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column19 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   13320
      Top             =   9000
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
      Left            =   11040
      Top             =   9000
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
      Left            =   8760
      Top             =   9000
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
      Left            =   4560
      Top             =   9360
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
      Top             =   9000
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2280
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
   Begin Crystal.CrystalReport CR02 
      Left            =   240
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   10
      WindowTop       =   10
      WindowWidth     =   600
      WindowHeight    =   900
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR03 
      Left            =   840
      Top             =   9720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowLeft      =   600
      WindowTop       =   10
      WindowWidth     =   600
      WindowHeight    =   900
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport CR04 
      Left            =   6360
      Top             =   9360
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
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000040&
      Height          =   5055
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   6540
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6105
         TabIndex        =   46
         Top             =   3660
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B2D
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   25
         Top             =   3645
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "0"
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5655
         TabIndex        =   51
         Top             =   2235
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "fmes_nro_horarios_hab"
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
         Height          =   290
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   47
         Top             =   1680
         Width           =   1090
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "observaciones"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   3240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "tw_cronograma_certifica_inst.frx":8B46
         Top             =   4080
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txt_codigo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ges_gestion"
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   520
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6090
         TabIndex        =   12
         Top             =   3015
         Width           =   270
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   290
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1680
         Width           =   1090
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4440
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B48
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   2220
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "zpiloto_codigo"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B61
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632256
         ListField       =   "unidad_codigo"
         BoundColumn     =   "unidad_codigo"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B7A
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   4965
         _ExtentX        =   8758
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
         DataField       =   "fecha_registro_cert"
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
         Left            =   1560
         TabIndex        =   0
         Top             =   4440
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
         Format          =   128581635
         CurrentDate     =   44235
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_cronograma_certifica_inst.frx":8B93
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   2220
         Width           =   4125
         _ExtentX        =   7276
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
         Bindings        =   "tw_cronograma_certifica_inst.frx":8BAC
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   3645
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horarios Hábiles X Mes"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   49
         Top             =   1680
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "fmes_nro_dias_habiles"
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
         Left            =   5280
         TabIndex        =   45
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dias Hábiles X Mes"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   44
         Top             =   1100
         Width           =   1650
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Responsable de la Zona"
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
         Top             =   3420
         Width           =   2820
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
         TabIndex        =   29
         Top             =   2775
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Horas Hábiles X Mes"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   1770
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
         Left            =   4800
         TabIndex        =   27
         Top             =   4455
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Registro"
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
         TabIndex        =   23
         Top             =   4460
         Width           =   1305
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
         Left            =   2040
         TabIndex        =   22
         Top             =   1080
         Width           =   1090
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Dias X Mes"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1095
         Width           =   1740
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
         Left            =   4680
         TabIndex        =   20
         Top             =   285
         Width           =   1545
      End
      Begin VB.Label lbl_texto2 
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
         Left            =   1800
         TabIndex        =   19
         Top             =   525
         Width           =   2415
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
         Left            =   1800
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   280
         Width           =   660
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto"
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
         TabIndex        =   16
         Top             =   2235
         Width           =   990
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
         TabIndex        =   15
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
         Left            =   4800
         TabIndex        =   14
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro.Horas X Mes"
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
      Index           =   5
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "tw_cronograma_certifica_inst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset

Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset

Dim rs_aux8 As New ADODB.Recordset
'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial0 As String

'OTROS
'Dim swnuevo As String
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2, VAR_EQUIPO As String
Dim SQL_FOR As String
Dim sql As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_TIT, VAR_ANL As String
Dim VAR_DA, VAR_UORIGEN, VAR_DPTOC As String

Dim VAR_AUX, VAR_CONT2 As Double

Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Dim VAR_AUX2, VAR_COD0, CONT3 As Integer
Dim DIAS_HAB, NRO_HRS, NRO_HORARIO As Integer
Dim VAR_REG, VAR_CANT1 As Integer
Dim VAR_BUSCA As Integer

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
        If Ado_datos.Recordset.RecordCount > 0 Then
            buscados = buscados + 1
            If buscados = 1 Then
                Call ABRIR_TABLA_DET(1)
                If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                    lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                End If
                'mes2 = MonthName(Month(DTPFec_Inicio.Value))
                buscados = buscados + 1
            End If
            
            Call ABRIR_TABLA_DET(1)
            If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
            End If
            'mes2 = MonthName(Month(DTPFec_Inicio.Value))
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnBuscar_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
'        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
''        If Ado_datos.Recordset!estado_codigo = "REG" Then
''            Call OptFilGral1_Click
''        Else
''            Call OptFilGral2_Click
''        End If
        VAR_BUSCA = 0
        queryinicial = queryinicial0
        Set ClBuscaGrid = New ClBuscaEnGridExterno
        Set ClBuscaGrid.Conexión = db
        ClBuscaGrid.EsTdbGrid = False
        Set ClBuscaGrid.GridTrabajo = dg_det2
        ClBuscaGrid.QueryUtilizado = queryinicial
        Set ClBuscaGrid.RecordsetTrabajo = rs_det2
        'ClBuscaGrid.CamposVisibles = "11010011"
        ClBuscaGrid.Ejecutar
    Else
      MsgBox "NO se puede Procesar !!. Verifique si existen registros. ", vbExclamation, "Atención!"
      'OptFilGral1.Visible = True
      'OptFilGral2.Visible = True
    End If

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
        fraOpciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
    End If

End Sub

Private Sub BtnCancelar3_Click()
    'fraOpciones.Enabled = True
    ' fraOpciones2.Enabled = True
    ' FrmABMDet.Enabled = True
     FraDet2.Visible = False
     BtnImprimir2.Visible = True
End Sub

Private Sub BtnCancelarDet_Click()
'    BtnModDetalle.Visible = True
'    BtnImprimir2.Visible = True
'    BtnGrabarDet.Visible = False
'    BtnCancelarDet.Visible = False
''    dg_det2.Enabled = False
'    dg_det2.AllowUpdate = False
End Sub

Private Sub BtnGraba3_Click()
    'db.Execute "update to_cronograma_diario_final set fecha_conformidad = '" & DTPEjecucion.Value & "', nro_fojas = " & txt_hdm.Text & ", doc_numero = " & txt_cm.Text & ", observaciones = '" & txt_obs.Text & "', carta = '" & Cmb_carta.Text & "', doc_numero_carta = '" & txt_correl_carta.Text & "' where fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' "
    db.Execute "update to_cronograma_diario_final_INST set fecha_conformidad = '" & DTPEjecucion.Value & "', doc_numero = " & txt_cm.Text & ", observaciones = '" & txt_obs.Text & "', carta = '" & Cmb_carta.Text & "', doc_numero_carta = '" & txt_correl_carta.Text & "'  where fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "' and horario_codigo = " & Ado_detalle2.Recordset!horario_codigo & "  "
    db.Execute "update tc_zona_piloto_edif_inst set estado_activo = 'APR'  where correlativo = " & Ado_detalle2.Recordset!fmes_plan & " "
    FraDet2.Visible = False
    BtnImprimir2.Visible = True
    VAR_EQUIPO = Ado_detalle2.Recordset!bien_codigo
    'db.Execute "tp_certificados_actulizacion"
    Call ABRIR_TABLA_DET(1)
     If (dg_det2.SelBookmarks.Count <> 0) Then
        dg_det2.SelBookmarks.Remove 0
     End If
     If Ado_detalle2.Recordset.RecordCount > 0 Then
        rs_det2.Find "bien_codigo = '" & VAR_EQUIPO & "'   ", , , 1
        dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
     Else
        rs_det2.MoveLast
     End If
    FraDet2.Visible = False
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     '
     Set rs_aux5 = New ADODB.Recordset
     If rs_aux5.State = 1 Then rs_aux5.Close
     rs_aux5.Open "select dia_correl from to_cronograma_diario where fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and estado_activo <> 'ANL' group by dia_correl", db, adOpenStatic
     If rs_aux5.RecordCount > 0 Then
        DIAS_HAB = rs_aux5.RecordCount
     End If
        
     Set rs_aux5 = New ADODB.Recordset
     If rs_aux5.State = 1 Then rs_aux5.Close
     rs_aux5.Open "select COUNT(dia_correl) as nro_horarios, SUM(nro_total_horas) as nro_horas from to_cronograma_diario where fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and estado_activo <> 'ANL' ", db, adOpenStatic
     If rs_aux5.RecordCount > 0 Then
        NRO_HORARIO = rs_aux5!nro_horarios
        NRO_HRS = rs_aux5!nro_horas
     End If
     
     rs_datos!fmes_fecha_registro = DTPfecha1.Value
     rs_datos!beneficiario_codigo_resp = dtc_codigo4.Text
     rs_datos!observaciones = Txt_campo2.Text
     
     rs_datos!fmes_nro_dias_habiles = DIAS_HAB
     rs_datos!fmes_nro_horarios_hab = NRO_HORARIO
     rs_datos!fmes_nro_hrs_habiles = NRO_HRS
     
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & ", beneficiario_codigo_resp2 = " & dtc_codigo4.Text & " Where fmes_plan = " & Ado_datos.Recordset!fmes_plan & "   "
     
     Call OptFilGral2_Click
     rs_datos.MoveFirst
'     mbDataChanged = False

     Fra_datos.Enabled = False
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
  If (dtc_codigo4 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (Txt_campo2.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub BtnGrabarDet_Click()
    BtnModDetalle.Visible = True
    BtnImprimir2.Visible = True
    BtnGrabarDet.Visible = False
    BtnCancelarDet.Visible = False
    dg_det2.AllowUpdate = False
    fraOpciones.Visible = True
    FraGrabarCancelar.Visible = True
    FraNavega.Enabled = True
    db.Execute "tp_certificados_actulizacion"
''    dg_det2.Enabled = False
    Call ABRIR_TABLA_DET(1)
End Sub

Private Sub BtnImprimir_Click()
If Ado_detalle2.Recordset.RecordCount > 0 Then
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual.rpt"
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
      CR01.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR01.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
      
'    cr01.StoredProcParam(0) = "2015"    'Me.Ado_datos.Recordset!ges_gestion
'    cr01.StoredProcParam(1) = "DNMAN"   'Me.Ado_datos.Recordset!unidad_codigo_tec
'    cr01.StoredProcParam(2) = 0     'Me.Ado_datos.Recordset!zpiloto_codigo
'    cr01.StoredProcParam(3) = 1     'Me.Ado_datos.Recordset!fmes_correl
    
    CR01.StoredProcParam(0) = Me.Ado_detalle2.Recordset!ges_gestion
    CR01.StoredProcParam(1) = Me.Ado_detalle2.Recordset!unidad_codigo_tec
    CR01.StoredProcParam(2) = Me.Ado_detalle2.Recordset!zpiloto_codigo
    CR01.StoredProcParam(3) = Me.Ado_detalle2.Recordset!fmes_correl
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir1_Click()
If Ado_detalle2.Recordset.RecordCount > 0 Then
    'to_cronograma_diario_final
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "select distinct bien_codigo  from to_cronograma_diario_final_INST where fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " and bien_codigo <>'' ", db, adOpenStatic
    If rs_datos1.RecordCount > 0 Then
        VAR_REG = rs_datos1.RecordCount
        VAR_CANT1 = rs_datos1.RecordCount
    Else
        VAR_REG = "0"
        VAR_CANT1 = "0"
    End If
    'lbl_texto2.Caption = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    'CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual_eqp.rpt"
    CR04.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_Instalacion_PRUEBA.rpt"
    CR04.WindowShowPrintSetupBtn = True
    CR04.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
'    Select Case Me.Ado_detalle2.Recordset!unidad_codigo_tec
'          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
'          Case "DNAJS"
'              var_titulo = "Módulo Ajustes"
'          Case "DNMAN", "DMANS", "DMANB", "DMANC"
'              var_titulo = "Módulo Mantenimiento"
'          Case "DNREP"
'              var_titulo = "Módulo Reparaciones"
'          Case "DNEME"
'              var_titulo = "Módulo Emergencias"
'          Case "DNMOD"
'              var_titulo = "Módulo Modernización"
'      End Select
      'Cmb_Mes.Text = "ENERO"
      CR04.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR04.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR04.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      CR04.Formulas(3) = "TotalReg = " & VAR_REG & " "
      CR04.Formulas(4) = "CANT1 = " & VAR_CANT1 & " "
      
     CR04.StoredProcParam(0) = Me.Ado_detalle2.Recordset!fmes_plan
     CR04.StoredProcParam(1) = Me.Ado_detalle2.Recordset!zpiloto_codigo
     
    iResult = CR04.PrintReport
    If iResult <> 0 Then MsgBox CR04.LastErrorNumber & " : " & CR04.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR04.WindowState = crptMaximized

End Sub

Private Sub BtnImprimir2_Click()
    If Ado_detalle2.Recordset.RecordCount > 0 Then
        Dim iResult As Integer
        'Dim co As New ADODB.Command
        'CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_acta_entrega_equipos_PRUEBA.rpt"
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_acta_entrega_equipos.rpt"
        CR02.WindowShowPrintSetupBtn = True
        CR02.WindowShowRefreshBtn = True
        CR02.StoredProcParam(0) = Ado_detalle2.Recordset!fmes_plan
        CR02.StoredProcParam(1) = Ado_detalle2.Recordset!horario_codigo
        CR02.StoredProcParam(2) = Ado_detalle2.Recordset!EDIF_CODIGO
        iResult = CR02.PrintReport
        If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
        CR02.WindowState = crptMaximized
        
        
        'CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_acta_entrega_equipos_PRUEBA2.rpt"
        CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_acta_entrega_equipos2.rpt"
        CR03.WindowShowPrintSetupBtn = True
        CR03.WindowShowRefreshBtn = True
        CR03.StoredProcParam(0) = Ado_detalle2.Recordset!fmes_plan
        CR03.StoredProcParam(1) = Ado_detalle2.Recordset!horario_codigo
        CR03.StoredProcParam(2) = Ado_detalle2.Recordset!EDIF_CODIGO
        iResult = CR03.PrintReport
        If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
        CR03.WindowState = crptMaximized
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
End Sub

Private Sub BtnImprimir3_Click()
'If Ado_detalle2.Recordset.RecordCount > 0 Then
'    If glusuario = "JORAQUENI" Then
'        MsgBox "El Usuario No tiene acceso, Consulte con el Administrador del Sistema ...", , "Atención"
'        Exit Sub
'    End If
'    Dim iResult As Integer
'    'Dim co As New ADODB.Command
'    CR03.ReportFileName = App.Path & "\Reportes\tecnico\tr_cronograma_mensual_ejecucion_eqp.rpt"
'    CR03.WindowShowPrintSetupBtn = True
'    CR03.WindowShowRefreshBtn = True
'    'MsgBox rs.RecordCount
'    Select Case Me.Ado_detalle2.Recordset!unidad_codigo_tec
'          Case "DNINS"
'              var_titulo = "Módulo Instalaciones"
'          Case "DNAJS"
'              var_titulo = "Módulo Ajustes"
'          Case "DNMAN", "DMANS", "DMANB", "DMANC"
'              var_titulo = "Módulo Mantenimiento"
'          Case "DNREP"
'              var_titulo = "Módulo Reparaciones"
'          Case "DNEME"
'              var_titulo = "Módulo Emergencias"
'          Case "DNMOD"
'              var_titulo = "Módulo Modernización"
'      End Select
'      'Cmb_Mes.Text = "ENERO"
'      VAR_TIT = "EJECUCION SERVICIO DE MANTENIMIENTO"
'      CR03.Formulas(0) = "titulo = '" & VAR_TIT & "' "
'      CR03.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
'      CR03.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
'
'     CR03.StoredProcParam(0) = Ado_detalle2.Recordset!fmes_plan
'     CR03.StoredProcParam(1) = Ado_detalle2.Recordset!zpiloto_codigo
''    'CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'
'    iResult = CR03.PrintReport
'    If iResult <> 0 Then MsgBox CR03.LastErrorNumber & " : " & CR03.LastErrorString, vbCritical, "Error de impresión"
'Else
'    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
'End If
'    CR03.WindowState = crptMaximized

End Sub

Private Sub BtnImprimir4_Click()
If Ado_detalle2.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cronograma_mensual_ejecucion.rpt"
    '
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_detalle2.Recordset!unidad_codigo_tec
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
      VAR_TIT = "EJECUCION SERVICIO DE MANTENIMIENTO"
      CR02.Formulas(0) = "titulo = '" & VAR_TIT & "' "
      CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR02.Formulas(3) = "TotalReg = 0 "          '" & VAR_REG & " "
      'CR02.Formulas(4) = "CANT1 = 0 "               '" & VAR_CANT1 & " "
      
     CR02.StoredProcParam(0) = Ado_detalle2.Recordset!fmes_plan
     CR02.StoredProcParam(1) = Ado_detalle2.Recordset!zpiloto_codigo
''      CR02.Formulas(0) = "@titulo = '" & var_titulo & "' "
''      CR02.Formulas(1) = "@subtitulo = '" & VAR_TIT & "' "
''      'CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
''      CR02.Formulas(2) = "@periodo = '" & lbl_texto2 & "' "
''      'CR02.Formulas(2) = "periodo = '" & Cmb_Mes & "' "
'
'    'CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!fmes_plan
'    CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!zpiloto_codigo
'    'CR02.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl

    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized
End Sub

Private Sub BtnModDetalle_Click()
'If glusuario = "ADMIN" Or glusuario = "AURBINA" Or glusuario = "JSAAVEDRA" Or glusuario = "JAGUTIERREZ" Or glusuario = "KGARCIA" Or glusuario = "OCOLODRO" Or glusuario = "VMEJIA" Or glusuario = "RVALDIVIEZO" Or glusuario = "BMONTAÑO" Or glusuario = "RALARCON" Or glusuario = "FDELGADILLO" Or glusuario = "RRONDAL" Or glusuario = "FDELGADILLO" Or glusuario = "RRONDAL" Then
If glusuario = "ADMIN" Or glusuario = "AURBINA" Or glusuario = "JSAAVEDRA" Or glusuario = "MARTEAGA" Or glusuario = "KGARCIA" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "VMEJIA" Or glusuario = "FDELGADILLO" Or glusuario = "RRONDAL" Or glusuario = "EVILLALOBOS" Or glusuario = "LVEDIA" Or glusuario = "JCASTRO" Or glusuario = "ASANTIVAÑEZ" Or glusuario = "CSALINAS" Or glusuario = "ARODRIGUEZ" Or glusuario = "FFLORES" Then
   If Ado_detalle2.Recordset("estado_codigo") = "REG" Then
        VAR_BUSCA = 1
'      'sino = MsgBox("Para modificar elija una de las 2 opciones: (SI=Modifica SOLO el Registro Elegido, NO=Acceso a Modificar a TODOS los Registros de la ZONA) ", vbYesNo + vbQuestion, "Atención")
'      sino = MsgBox("Para modificar elija una de las 2 opciones:" & vbCr & _
'             "SI = Modifica SOLO el Registro Elegido " & vbCr & _
'             "NO = Acceso a Modificar a TODOS los Registros de la ZONA", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
        BtnImprimir2.Visible = False
        FraDet2.Visible = True
'      Else
'        fraOpciones.Visible = False
'        Frame1.Enabled = False
'        Frame2.Visible = True
'        dg_det3.AllowUpdate = True
'
'        'BtnModDetalle.Visible = False
'        'BtnImprimir2.Visible = False
'        'BtnGrabarDet.Visible = True
'        'BtnCancelarDet.Visible = True
'        'FraGrabarCancelar.Visible = False
'        'FraNavega.Enabled = False
'      End If
   Else
        MsgBox "No se puede Modificar, el registro ya fue Aprobado (Estado=APR) o está Anulado (Estado=ANL) ...", vbExclamation, "Validación de Registro"
   End If
Else
    MsgBox "El Usuario No tiene Acceso ...", vbExclamation, "Validación de Registro"
End If
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

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    If dtc_desc2.SelectedItem <> "" Then
         Call ABRIR_TABLA_DET(2)
    End If
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    VAR_BUSCA = 0
'    busca3 = 0
    'cmd_campo2.Text = "2"
    'Fra_Gestion.Visible = True
    VAR_GES = Year(Date)        'Cmb_gestion.Text
    
    Set rs_aux8 = New ADODB.Recordset
    If rs_aux8.State = 1 Then rs_aux8.Close
    rs_aux8.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux8.RecordCount > 0 Then
        usuario2 = rs_aux8!beneficiario_codigo
        VAR_DA = rs_aux8!da_codigo
        VAR_DPTOC = rs_aux8!depto_codigo
    Else
        usuario2 = "6753027"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    If Aux = "DNINS" Then
        Select Case VAR_DPTOC
            Case "1"    ' Chuquisaca
                VAR_UORIGEN = "DINSC"
            Case "2"    'La Paz - Tecnico
                VAR_UORIGEN = "DNINS"
            Case "3"    'Cochabamba
                VAR_UORIGEN = "DINSB"
                'VAR_DPTOC = "3"
            Case "7"    'Santa Cruz
                VAR_UORIGEN = "DINSS"
                'VAR_DPTOC = "7"
            Case "4"    'Oruro - Tecnico
                VAR_UORIGEN = "DNINS"
                'VAR_DPTOC = "2"
            Case "5"    ' Potosi
                VAR_UORIGEN = "DNINS"
            Case "6"    ' Tarija
                VAR_UORIGEN = "DNINS"
            Case "8"    ' Beni
                VAR_UORIGEN = "DNINS"
            Case "9"    ' Pando
                VAR_UORIGEN = "DNINS"
            Case Else    ' TODO
                VAR_UORIGEN = "DNINS"
                VAR_DPTOC = "0"
         End Select

'    If Aux = "DNMAN" Then
'        Select Case VAR_DA
'            Case "1.8"    'Cochabamba
'                VAR_UORIGEN = "DMANB"
'                VAR_DPTOC = "3"
'            Case "1.7"    'Santa Cruz
'                VAR_UORIGEN = "DMANS"
'                VAR_DPTOC = "7"
'            Case "1.3"    'La Paz - Tecnico
'                VAR_UORIGEN = "DNMAN"
'                VAR_DPTOC = "2"
'            Case "1.9"    ' Chuquisaca
'                VAR_UORIGEN = "DMANC"
'                VAR_DPTOC = "1"
'            Case Else    ' TODO
'                VAR_UORIGEN = "DNMAN"
'                VAR_DPTOC = "0"
'         End Select
     End If
    parametro = Aux
    VAR_ANL = ""
    

    Call ABRIR_TABLAS_AUX
    
    'db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.carta   = 'NO' WHERE carta IS NULL"
    
    'db.Execute "tp_certificados_res"
    
    'Call OptFilGral1_Click
    queryinicial0 = "select * from tv_ejecucion_final_inst  "           'where bien_codigo <> ''
    Call ABRIR_TABLA_DET(1)
    var_cod = "0"
    
'    Set rs_det1 = New ADODB.Recordset
'    If rs_det1.State = 1 Then rs_det1.Close
'    rs_det1.Open "select * from to_cronograma_diario_final where bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    'Set Ado_detalle1.Recordset = rs_det1
'    'Set dg_det1.DataSource = Ado_detalle1.Recordset
'    If rs_det1.RecordCount > 0 Then
'             rs_det1.MoveFirst
'
'             While Not rs_det1.EOF
'                    rs_det1!hora_registro = "0"
'                If var_cod = rs_det1!bien_codigo Then
'                    rs_det1!hora_registro = "1"
'                End If
'                var_cod = rs_det1!bien_codigo
'                rs_det1.Update
'                rs_det1.MoveNext
'             Wend
'
'    End If
    If glusuario = "MLLOSA" Then
'        BtnModificar.Visible = False
'        BtnEliminar.Visible = False
        BtnAprobar.Visible = False
        BtnModDetalle.Visible = False
        BtnGrabarDet.Visible = False
        BtnGraba3.Visible = False
    End If

    Fra_datos.Enabled = False
    dg_datos.Enabled = True
'    dg_det2.Enabled = False
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
    'rs_datos1.Open "Select * from gc_unidad_ejecutora order by unidad_descripcion", db, adOpenStatic
    rs_datos1.Open "gp_listar_apr_gc_unidad_ejecutora ", db, adOpenStatic
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
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
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

'Private Sub dtc_desc1_LostFocus()
''    dtc_codigo5.Text = dtc_aux1.Text
''    dtc_desc5.BoundText = dtc_codigo5.BoundText
''    Call pnivel5(dtc_codigo5.BoundText)
''    dtc_desc6.Enabled = True
'End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
End Sub

Private Sub OptFilGral0_Click()
    '===== Proceso para filtrado general de datos (los registros 2019)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2019') "
        Case "2"    'La Paz - Tecnico
            If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "RCUELA" Or glusuario = "OCOLODRO" Or glusuario = "JSAAVEDRA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Then
                queryinicial = "select * From to_cronograma_mensual WHERE ( ges_gestion = '2019') "     ' estado_certifica <> 'ANL' AND
            Else
                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2019' ) "
            End If
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2019') "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2019') "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2019') "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2019') "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2019') "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2019') "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2019') "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual  WHERE ( ges_gestion = '2019') "
     End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (los registros 2020)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2020') "
        Case "2"    'La Paz - Tecnico
            If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "RCUELA" Or glusuario = "OCOLODRO" Or glusuario = "JSAAVEDRA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Then
                queryinicial = "select * From to_cronograma_mensual WHERE ( ges_gestion = '2020') "     ' estado_certifica <> 'ANL' AND
            Else
                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2020' ) "
            End If
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2020') "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2020') "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2020') "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2020') "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2020') "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2020') "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2020') "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual  WHERE ( ges_gestion = '2020') "
     End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (los registros 2021)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2022') "
        Case "2"    'La Paz - Tecnico
            If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "RCUELA" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "JSAAVEDRA" Or glusuario = "VPAREDES" Or glusuario = "LVASQUEZ" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2022') "     ' estado_certifica <> 'ANL' AND
            Else
                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='33' OR zpiloto_codigo='36' OR zpiloto_codigo='37' OR zpiloto_codigo='39' OR zpiloto_codigo='40')  AND ges_gestion = '2022' ) "
            End If
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20' OR zpiloto_codigo='16' ) AND ges_gestion = '2022') "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2022') "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2022') "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2022') "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2022') "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2022') "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2022') "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual  WHERE ( ges_gestion = '2022') "
     End Select
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

Private Sub ABRIR_TABLA_DET(posicion As Integer)
    
  Select Case posicion
    Case 1
        var_cod = "0"
        Set rs_det1 = New ADODB.Recordset
        If rs_det1.State = 1 Then rs_det1.Close
    
    'rs_det1.Open "select * from to_cronograma_diario_final where bien_codigo <> '' and fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'", db, adOpenKeyset, adLockOptimistic, adCmdText
    'Set Ado_detalle1.Recordset = rs_det1
    'Set dg_det1.DataSource = Ado_detalle1.Recordset
'    If rs_det1.RecordCount > 0 Then
'             rs_det1.MoveFirst
'
'             While Not rs_det1.EOF
'                rs_det1!hora_registro = "0"
'                If var_cod = rs_det1!bien_codigo Then
'                    rs_det1!hora_registro = "1"
'                End If
'                var_cod = rs_det1!bien_codigo
'                rs_det1.Update
'                rs_det1.MoveNext
'             Wend

        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        'rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' and hora_registro = '0' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Open "select * from tv_ejecucion_final_inst where bien_codigo<>'' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        
        'VERIFICAR APROBADOS WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
        'queryinicial = "select * from tv_ejecucion_final_inst where estado_codigo_crono = 'APR' AND bien_codigo <> '' "
        If VAR_BUSCA = 0 Then
            queryinicial = "select * from tv_ejecucion_final_inst "         'where bien_codigo <> ''
        Else
            queryinicial = queryinicial0            'queryinicial99
        End If
        rs_det2.Open queryinicial, db, adOpenKeyset, adLockOptimistic
        rs_det2.Sort = "Ges_gestion, zpiloto_codigo, edif_codigo, horario_codigo"
        Set Ado_detalle2.Recordset = rs_det2
        dg_det2.Visible = True
'        If rs_det1.RecordCount > 0 Then
            Set dg_det2.DataSource = Ado_detalle2.Recordset
            dg_det2.Visible = True
'        Else
'            Set dg_det2.DataSource = rsNada
'            dg_det2.Visible = False
'        End If
'    Else
'        dg_det2.Visible = False
'    End If
    Case 2
        '--------------- buscar
        If (dg_det2.SelBookmarks.Count <> 0) Then
                dg_det2.SelBookmarks.Remove 0
        End If
        If rs_det2.RecordCount > 0 Then
            rs_det2.Find "edif_descripcion like '" & dtc_desc2.Text & "'", , , 1
            'rs_det2.Find "edif_codigo like '" & xxxxx & "'", , , 1
            dg_det2.SelBookmarks.Add (rs_det2.Bookmark)
        Else
            sino = MsgBox("No se encontro edificios con ese nombre", vbInformation, "Atencion!")
            Call ABRIR_TABLA_DET(2)
            dtc_desc2.Text = ""
        End If
  End Select
End Sub

'Private Sub ABRIR_TABLA_DET()
''    Set rs_det1 = New ADODB.Recordset
''    If rs_det1.State = 1 Then rs_det1.Close
''    rs_det1.Open "select * from to_cronograma_diario where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
''    Set Ado_detalle1.Recordset = rs_det1
''    Set dg_det1.DataSource = Ado_detalle1.Recordset
'
'    Set rs_det2 = New ADODB.Recordset
'    If rs_det2.State = 1 Then rs_det2.Close
'    rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' and hora_registro = '0' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    'rs_det2.Open "SELECT distinct fmes_plan, dia_correl, bien_orden, bien_codigo, unidad_codigo_tec, tec_plan_codigo, beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha, dia_nombre, nro_total_horas, observaciones, edif_descripcion, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, " & _
'    " bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, fecha_carta, fecha_conformidad, fecha_equipo_hdm, nro_fojas, doc_numero , estado_activo, estado_codigo, usr_codigo, fecha_registro, hora_registro  " & _
'    " From dbo.to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'
'    Set Ado_detalle2.Recordset = rs_det2
'    Set dg_det2.DataSource = Ado_detalle2.Recordset
'
'End Sub

Private Sub OptFilGral3_Click()
    '===== Proceso para filtrado general de datos (los registros 2020)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DPTOC
        Case "1"    ' Chuquisaca
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2021') "
        Case "2"    'La Paz - Tecnico
            If glusuario = "ASANTIVAÑEZ" Or glusuario = "ADMIN" Or glusuario = "APALACIOS" Or glusuario = "JCASTRO" Or glusuario = "RCUELA" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "JSAAVEDRA" Or glusuario = "VPAREDES" Or glusuario = "CSALINAS" Then
                queryinicial = "select * From to_cronograma_mensual WHERE ( ges_gestion = '2021') "     ' estado_certifica <> 'ANL' AND
            Else
                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' )  AND ges_gestion = '2021' ) "
            End If
        Case "3"    'Cochabamba
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2021') "
        Case "7"    'Santa Cruz
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2021') "
        Case "4"    'Oruro - Tecnico
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2021') "
        Case "5"    ' Potosi
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2021') "
        Case "6"    ' Tarija
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2021') "
        Case "8"    ' Beni
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2021') "
        Case "9"    ' Pando
            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2021') "
        Case Else    ' TODO
            queryinicial = "select * From to_cronograma_mensual  WHERE ( ges_gestion = '2021') "
     End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

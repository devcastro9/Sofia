VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_to_cronograma_mensual 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tecnico - Cronograma Mensual"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "frm_to_cronograma_mensual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8895
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
         Left            =   6120
         TabIndex        =   3
         Top             =   2835
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
         Left            =   1800
         TabIndex        =   2
         Top             =   2835
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2760
         Width           =   8625
         _ExtentX        =   15214
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
         Height          =   2490
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8640
         _ExtentX        =   15240
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
         ColumnCount     =   9
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
            DataField       =   "fmes_plan"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         BeginProperty Column08 
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   2520
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija para cambiar el Número de horas para el Servicio ..."
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
      Height          =   1680
      Left            =   7080
      TabIndex        =   70
      Top             =   5160
      Visible         =   0   'False
      Width           =   4860
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H80000015&
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   2760
         Picture         =   "frm_to_cronograma_mensual.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   840
         Width           =   1125
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H80000015&
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   960
         Picture         =   "frm_to_cronograma_mensual.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Aprueba Registro"
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox txtnrohrs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   72
         Text            =   "frm_to_cronograma_mensual.frx":0E16
         Top             =   360
         Width           =   645
      End
      Begin VB.ComboBox cmd_campo2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_to_cronograma_mensual.frx":0E18
         Left            =   3960
         List            =   "frm_to_cronograma_mensual.frx":0E25
         TabIndex        =   71
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbl_orden 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.de Horas actual"
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
         Left            =   120
         TabIndex        =   74
         Top             =   375
         Width           =   1725
      End
      Begin VB.Label lbl_orden_camb 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar a -->"
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
         Left            =   2760
         TabIndex        =   73
         Top             =   375
         Width           =   1140
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija la nueva Zona a la que se enviará el registro elegido"
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
      Height          =   2040
      Left            =   6000
      TabIndex        =   65
      Top             =   6960
      Visible         =   0   'False
      Width           =   7140
      Begin VB.CommandButton BtnCancelar3 
         BackColor       =   &H80000015&
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   3840
         Picture         =   "frm_to_cronograma_mensual.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1200
         Width           =   1125
      End
      Begin VB.CommandButton BtnGraba3 
         BackColor       =   &H80000015&
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   2040
         Picture         =   "frm_to_cronograma_mensual.frx":103C
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Aprueba Registro"
         Top             =   1200
         Width           =   1125
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6600
         TabIndex        =   66
         Top             =   690
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "frm_to_cronograma_mensual.frx":1246
         Height          =   315
         Left            =   240
         TabIndex        =   67
         Top             =   680
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "frm_to_cronograma_mensual.frx":125F
         Height          =   315
         Left            =   5880
         TabIndex        =   68
         Top             =   680
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
      Begin VB.Label lbl_campo5 
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
         TabIndex        =   69
         Top             =   405
         Width           =   990
      End
   End
   Begin VB.PictureBox fraOpciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   42
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   17760
         Picture         =   "frm_to_cronograma_mensual.frx":1278
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   62
         ToolTipText     =   "Cierra la Ventana Activa"
         Top             =   0
         Width           =   1245
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4320
         Picture         =   "frm_to_cronograma_mensual.frx":1A3A
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   52
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "frm_to_cronograma_mensual.frx":21EF
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   51
         ToolTipText     =   "Aprueba Cronograma"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1560
         Picture         =   "frm_to_cronograma_mensual.frx":2A22
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   50
         ToolTipText     =   "Anular Cronograma"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   105
         Picture         =   "frm_to_cronograma_mensual.frx":316E
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   49
         ToolTipText     =   "Modifica Datos Cabecera"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   7200
         Picture         =   "frm_to_cronograma_mensual.frx":3A83
         ScaleHeight     =   615
         ScaleWidth      =   1200
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton BtnVer 
         BackColor       =   &H00808000&
         Caption         =   "Digitaliza"
         Height          =   600
         Left            =   8760
         Picture         =   "frm_to_cronograma_mensual.frx":4242
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Guarda en Archivo Digital"
         Top             =   10
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton BtnDesAprobar 
         BackColor       =   &H00808080&
         Height          =   600
         Left            =   9720
         Picture         =   "frm_to_cronograma_mensual.frx":4684
         Style           =   1  'Graphical
         TabIndex        =   43
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
         Left            =   12255
         TabIndex        =   45
         Top             =   200
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
      Begin VB.PictureBox BtnGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         Picture         =   "frm_to_cronograma_mensual.frx":488E
         ScaleHeight     =   615
         ScaleWidth      =   1305
         TabIndex        =   47
         Top             =   0
         Width           =   1300
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4275
         Picture         =   "frm_to_cronograma_mensual.frx":5064
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   46
         Top             =   0
         Width           =   1400
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
         Left            =   14175
         TabIndex        =   41
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CRONOGRAMA FINAL (DESTINO)"
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
      Height          =   9015
      Left            =   10160
      TabIndex        =   37
      Top             =   720
      Width           =   8895
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Todos los Horarios"
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
         Left            =   5160
         TabIndex        =   88
         Top             =   8760
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Solo Horarios Laborables"
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
         Left            =   1560
         TabIndex        =   87
         Top             =   8760
         Width           =   2955
      End
      Begin VB.PictureBox fraOpciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   75
         ScaleHeight     =   660
         ScaleWidth      =   8760
         TabIndex        =   83
         Top             =   240
         Width           =   8760
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2160
            Picture         =   "frm_to_cronograma_mensual.frx":5950
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   86
            ToolTipText     =   "Habilita Horario"
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3960
            Picture         =   "frm_to_cronograma_mensual.frx":6265
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   85
            ToolTipText     =   "Anula Horario"
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
            Picture         =   "frm_to_cronograma_mensual.frx":69B1
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   84
            ToolTipText     =   "Imprime R-302 Cronograma Mensual Final (Destino)"
            Top             =   0
            Width           =   1400
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "frm_to_cronograma_mensual.frx":727E
         Height          =   7785
         Left            =   75
         TabIndex        =   38
         Top             =   960
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   13732
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
         ColumnCount     =   14
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
         BeginProperty Column13 
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
               ColumnWidth     =   1110.047
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
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
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
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1695.118
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   8985
      Left            =   8955
      ScaleHeight     =   8955
      ScaleWidth      =   1125
      TabIndex        =   18
      Top             =   790
      Width           =   1160
      Begin VB.CommandButton BtnAddDetalle3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "A otra zona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         MaskColor       =   &H80000014&
         Picture         =   "frm_to_cronograma_mensual.frx":7299
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5880
         Width           =   1125
      End
      Begin VB.CommandButton BtnAnlDetalle3 
         BackColor       =   &H80000006&
         Caption         =   "Retor.Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         Picture         =   "frm_to_cronograma_mensual.frx":76DB
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Cambia el Horario a NO LABORABLE (Anula horario)"
         Top             =   7680
         Width           =   1125
      End
      Begin VB.CommandButton BtnModDetalle2 
         BackColor       =   &H80000006&
         Caption         =   "Envía Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         MaskColor       =   &H00FFFF80&
         Picture         =   "frm_to_cronograma_mensual.frx":7B1D
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   6840
         Width           =   1125
      End
      Begin VB.CommandButton BtnAnlDetalle2 
         BackColor       =   &H80000006&
         Caption         =   "Retornar 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         Picture         =   "frm_to_cronograma_mensual.frx":7F5F
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cambia el Horario a NO LABORABLE (Anula horario)"
         Top             =   4920
         Width           =   1125
      End
      Begin VB.CommandButton BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         Caption         =   "Envia Uno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         MaskColor       =   &H80000014&
         Picture         =   "frm_to_cronograma_mensual.frx":83A1
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4080
         Width           =   1125
      End
      Begin VB.CommandButton BtnImprimir3 
         BackColor       =   &H80000015&
         Caption         =   "Edif.X.Zona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   0
         Picture         =   "frm_to_cronograma_mensual.frx":87E3
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Imprime Edificios por Zonas"
         Top             =   720
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CRONOGRAMA ELABORADO (ORIGEN)"
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
      Height          =   5775
      Left            =   0
      TabIndex        =   14
      Top             =   3960
      Width           =   8925
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Todos los Horarios"
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
         Left            =   5040
         TabIndex        =   90
         Top             =   5520
         Width           =   2355
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Solo Horarios Llenados"
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
         Left            =   1440
         TabIndex        =   89
         Top             =   5520
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   75
         ScaleHeight     =   660
         ScaleWidth      =   8775
         TabIndex        =   80
         Top             =   240
         Width           =   8775
         Begin VB.PictureBox BtnAnlDetalle4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            Picture         =   "frm_to_cronograma_mensual.frx":9F65
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   82
            ToolTipText     =   "Anula Horario"
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4080
            Picture         =   "frm_to_cronograma_mensual.frx":A6B1
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   81
            ToolTipText     =   "Imprime R-302 Origen (Borrador)"
            Top             =   0
            Width           =   1400
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "frm_to_cronograma_mensual.frx":AF7E
         Height          =   4565
         Left            =   75
         TabIndex        =   15
         Top             =   960
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8043
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
         ColumnCount     =   15
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
         BeginProperty Column13 
            DataField       =   "estado_codigo"
            Caption         =   "Enviado"
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
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   2715.024
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   0
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
      Top             =   10200
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
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   6480
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
   Begin MSAdodcLib.Adodc Ado_detalle2 
      Height          =   330
      Left            =   2280
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
      Left            =   5040
      Top             =   10560
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
      Height          =   2540
      Left            =   6360
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   13260
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
         Left            =   11440
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   59
         Top             =   1320
         Width           =   1410
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8990
         TabIndex        =   58
         Top             =   2095
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8990
         TabIndex        =   57
         Top             =   1570
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8290
         TabIndex        =   56
         Top             =   1040
         Width           =   255
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "observaciones"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "frm_to_cronograma_mensual.frx":AF99
         Top             =   2640
         Visible         =   0   'False
         Width           =   10320
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
         Left            =   7200
         TabIndex        =   12
         Top             =   1575
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
         Left            =   11445
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   7
         Top             =   555
         Width           =   1410
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
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2085
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7220
         TabIndex        =   5
         Top             =   1035
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "frm_to_cronograma_mensual.frx":AF9B
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   1020
         Width           =   1005
         _ExtentX        =   1773
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
         Bindings        =   "frm_to_cronograma_mensual.frx":AFB4
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
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
         Bindings        =   "frm_to_cronograma_mensual.frx":AFCD
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
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
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "frm_to_cronograma_mensual.frx":AFE6
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1020
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
         Bindings        =   "frm_to_cronograma_mensual.frx":AFFF
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1920
         TabIndex        =   28
         Top             =   2085
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "frm_to_cronograma_mensual.frx":B018
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   7560
         TabIndex        =   29
         Top             =   2085
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "0"
      End
      Begin MSComCtl2.DTPicker DTPfecha1 
         DataField       =   "fmes_fecha_registro"
         DataSource      =   "Ado_datos"
         Height          =   300
         Left            =   9840
         TabIndex        =   75
         Top             =   2085
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   82575361
         CurrentDate     =   42408
         MaxDate         =   55153
         MinDate         =   2
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
         Left            =   10875
         TabIndex        =   61
         Top             =   1080
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
         Left            =   9120
         TabIndex        =   55
         Top             =   525
         Width           =   1335
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
         Left            =   8925
         TabIndex        =   54
         Top             =   285
         Width           =   1650
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
         TabIndex        =   36
         Top             =   2650
         Visible         =   0   'False
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2095
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
         TabIndex        =   34
         Top             =   1570
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
         Left            =   11025
         TabIndex        =   32
         Top             =   315
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
         Left            =   12240
         TabIndex        =   31
         Top             =   1845
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9840
         TabIndex        =   27
         Top             =   1845
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
         Left            =   6975
         TabIndex        =   26
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dias X Mes"
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
         Left            =   6915
         TabIndex        =   25
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
         Left            =   4680
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   1030
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
         Left            =   4800
         TabIndex        =   16
         Top             =   525
         Width           =   1335
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
      TabIndex        =   60
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frm_to_cronograma_mensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos2 As New ADODB.Recordset
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
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
Dim queryinicial As String

'OTROS
'Dim swnuevo As String
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES, gestion0 As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_ZONA, VAR_UNITEC As String

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
Dim VAR_ORDEN, VAR_MES As Integer
Dim buscados As Integer

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
                Call ABRIR_TABLA_DET
                If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                    lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                End If
                'mes2 = MonthName(Month(DTPFec_Inicio.Value))
                buscados = buscados + 1
            End If
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
    'GENERA CRONOGRAMA FINAL ITEM x ITEM (INI)
    fraOpciones.Enabled = False
    fraOpciones2.Enabled = False
    FrmABMDet.Enabled = False
    FraDet3.Visible = True
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "Select * from to_cronograma_detalle WHERE unidad_codigo_tec = '" & Ado_detalle1.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_detalle1.Recordset!tec_plan_codigo & "  and bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  ", db, adOpenStatic
    If rs_aux7.RecordCount > 0 Then
        'txtnrohrs.Text = rs_aux7!bien_cantidad_por_empaque
        'cmd_campo2.Text = rs_aux7!bien_cantidad_por_empaque
        txtnrohrs.Text = Ado_detalle1.Recordset!nro_total_horas
        cmd_campo2.Text = Ado_detalle1.Recordset!nro_total_horas
    Else
        txtnrohrs.Text = "2"
        cmd_campo2.Text = "2"
    End If
    'GENERA CRONOGRAMA FINAL ITEM x ITEM (FIN)
End Sub

Private Sub BtnAddDetalle3_Click()
    'CCCCCCCCCCCCCCCCCCCCCCCCCCCBBBBBBBBBBBBBBB
    fraOpciones.Enabled = False
    fraOpciones2.Enabled = False
    FrmABMDet.Enabled = False
    FraDet2.Visible = True
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    rs_aux7.Open "Select * from tc_zonas_piloto WHERE zpiloto_codigo <> '" & Ado_datos.Recordset!zpiloto_codigo & "'   ", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_aux7
    If rs_aux7.RecordCount > 0 Then
        dtc_desc5.BoundText = dtc_codigo5.BoundText
        VAR_ZONA = rs_aux7!zpiloto_codigo
    Else
        VAR_ZONA = "0"
    End If
    'WWWWW GENERA CRONOGRAMA DIARIO (FIN)
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle2.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma Final - Destino) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle2.Recordset!estado_activo = "ANL"
        Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
        Ado_detalle2.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue Aprobado (Estado=APR) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle2_Click()
   If Ado_detalle2.Recordset!estado_codigo = "REG" And Ado_detalle2.Recordset!estado_activo = "APR" Then
      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Final) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "update to_cronograma_diario set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        db.Execute "update to_cronograma_diario_final set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        'db.Execute "update to_cronograma_diario_final set bien_orden = 0 WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        'db.Execute "update to_cronograma_diario_final set estado_activo = 'REG' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAnlDetalle3_Click()
    'to_cronograma_diario_final
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_codigo <> '' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
      sino = MsgBox("Está Seguro de RETORNAR TODO ? (Se Retornará TODO el Cronograma DESTINO al ORIGEN) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "UPDATE to_cronograma_diario_final SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', estado_activo = 'REG' WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "

        db.Execute "UPDATE to_cronograma_diario set estado_codigo   = 'REG' where fmes_plan  = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "
      
        Call ABRIR_TABLA_DET
      End If
    Else
        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique los registros ...", vbExclamation, "Validación de Registro"
    End If
End Sub

Private Sub BtnAnlDetalle4_Click()
   If Ado_detalle1.Recordset!estado_codigo = "REG" Then     'And Ado_detalle2.Recordset!estado_activo = "APR"
      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Elaborado - Origen) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "update to_cronograma_diario set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle1.Recordset!bien_orden & " AND  bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
        db.Execute "update to_cronograma_diario set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle1.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
        Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede ANULAR, el registro ya fue ENVIADO al Cronograma Destino o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
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
        buscados = 0
'        OptFilGral1.Visible = True
'        OptFilGral2.Visible = True
''        If Ado_datos.Recordset!estado_codigo = "REG" Then
''            Call OptFilGral1_Click
''        Else
''            Call OptFilGral2_Click
''        End If
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

Private Sub BtnCancelar2_Click()
    fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet3.Visible = False
End Sub

Private Sub BtnCancelar3_Click()
    fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet2.Visible = False
End Sub

Private Sub BtnGraba3_Click()
   'CCCCCCCCCCCCCCCCCCCCCCCCCCCBBBBBBBBBBBBBBB
   VAR_ZONA = dtc_codigo5.Text
   VAR_MES = lbl_texto1.Caption
   gestion0 = txt_codigo1.Text
   
     Set rs_aux4 = New ADODB.Recordset
     If rs_aux4.State = 1 Then rs_aux4.Close
     rs_aux4.Open "select * from to_cronograma_diario where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
     If rs_aux4.RecordCount > 0 Then
        If rs_aux4!estado_codigo = "APR" Then
            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
            Exit Sub
        End If
        VAR_UNITEC = Ado_detalle1.Recordset!unidad_codigo_tec
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "  and unidad_codigo_tec = '" & VAR_UNITEC & "'   ", db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
             VAR_AUX2 = rs_aux2!fmes_plan
             VAR_COD0 = 0
             'db.Execute "SELECT VAR_ORDEN = isnull(max(bien_orden),0) from to_cronograma_diario WHERE     (fmes_plan = " & VAR_AUX2 & " ) "
            Set rs_aux5 = New ADODB.Recordset
            If rs_aux5.State = 1 Then rs_aux5.Close
            rs_aux5.Open "select isnull(max(bien_orden),0) as bien_orden2 from to_cronograma_diario WHERE fmes_plan = " & VAR_AUX2 & "  ", db, adOpenStatic
            If rs_aux5.RecordCount > 0 Then
               VAR_ORDEN = rs_aux5!bien_orden2 + 1
            End If
             Set rs_aux3 = New ADODB.Recordset
             If rs_aux3.State = 1 Then rs_aux3.Close
             rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
             If rs_aux3.RecordCount > 0 Then
                 rs_aux3.MoveFirst
                 While Not rs_aux3.EOF
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then
                        db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux4!bien_codigo & "', unidad_codigo_tec = '" & rs_aux4!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux4!tec_plan_codigo & ", observaciones = '" & rs_aux4!observaciones & "', bien_orden = " & VAR_ORDEN & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario set estado_codigo = 'REG', estado_activo = 'REG', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', observaciones = 'HORARIO LABORABLE'  WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & rs_aux4!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                    End If
                    rs_aux3.MoveNext
                    'Habilitar .....
                    'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
                 Wend
             End If
        End If
     End If
     db.Execute "update to_cronograma_diario set to_cronograma_diario.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario.bien_codigo  = av_bienes_vs_edificios.bien_codigo "
     Call ABRIR_TABLA_DET
    fraOpciones.Enabled = True
    fraOpciones2.Enabled = True
    FrmABMDet.Enabled = True
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

     rs_datos!Fecha_Registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & ", beneficiario_codigo_resp2 = " & dtc_codigo4.Text & " Where fmes_plan = " & Ado_datos.Recordset!fmes_plan & "   "
     db.Execute "Update to_cronograma_diario_final Set beneficiario_codigo_resp = " & dtc_codigo4.Text & ", beneficiario_codigo_resp2 = " & dtc_codigo4.Text & " Where fmes_plan = " & Ado_datos.Recordset!fmes_plan & "   "
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

Private Sub BtnGrabar2_Click()
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
     If rs_aux2.RecordCount > 0 Then
        If rs_aux2!estado_codigo = "APR" Then
            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
            Exit Sub
        End If
         VAR_AUX2 = rs_aux2!fmes_plan
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                   If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                         db.Execute "update to_cronograma_diario set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                         db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", observaciones = '" & rs_aux2!observaciones & "', bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         'db.Execute "update to_cronograma_diario_final set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         'db.Execute "update to_cronograma_diario_final set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         
                         VAR_COD0 = VAR_COD0 + 1
                         CONT3 = 1
                   End If
                   If cmd_campo2.Text = "4" Then
                      rs_aux3.MoveNext
                      If VAR_COD0 < 2 And rs_aux3!estado_activo = "REG" Then        '
                         db.Execute "update to_cronograma_diario set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                         db.Execute "update to_cronograma_diario_final set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", observaciones = '" & rs_aux2!observaciones & "', bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         'db.Execute "update to_cronograma_diario_final set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         'db.Execute "update to_cronograma_diario_final set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                         
                         VAR_COD0 = VAR_COD0 + 1
                         CONT3 = 1
                      End If
                   End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
     End If
     db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
     fraOpciones.Enabled = True
     fraOpciones2.Enabled = True
     FrmABMDet.Enabled = True
     FraDet3.Visible = False
     Call ABRIR_TABLA_DET
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
End Sub

Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual_eqp.rpt"
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
    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    'db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    '" where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_cronograma_mensual_origen.rpt"
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
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
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR02.Formulas(2) = "periodo = '" & Cmb_Mes & "' "

    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
    CR02.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
    CR02.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
    CR02.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir3_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    'db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    '" From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    'db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
    '" where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_zonas_vs_edificios.rpt"
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
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
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR02.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      'CR02.Formulas(2) = "periodo = '" & Cmb_Mes & "' "

    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!fmes_plan
      
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized

End Sub

Private Sub BtnModDetalle_Click()
    If Ado_detalle2.Recordset("estado_activo") = "ANL" Then  '<> "REG"
      sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        Ado_detalle2.Recordset!estado_activo = "REG"
        Ado_detalle2.Recordset!observaciones = "HORARIO LABORABLE"
        Ado_detalle2.Recordset.Update
        'Call ABRIR_TABLA_DET
      End If
   Else
        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnModDetalle2_Click()
    'to_cronograma_diario_final
    Set rs_aux6 = New ADODB.Recordset
    If rs_aux6.State = 1 Then rs_aux6.Close
    rs_aux6.Open "Select * from to_cronograma_diario_final where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_codigo <> '' ", db, adOpenStatic
    If rs_aux6.RecordCount > 0 Then
        MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), debe deshabilitarlos (Retornar) o utilizar el botón (Envia Uno) ...", vbExclamation, "Validación de Registro"
    Else
      sino = MsgBox("Está Seguro de ENVIA TODO ? (Se enviará TODO el Cronograma ORIGEN al DESTINO) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        db.Execute "UPDATE to_cronograma_diario_final SET to_cronograma_diario_final.bien_orden  = to_cronograma_diario.bien_orden, to_cronograma_diario_final.bien_codigo = to_cronograma_diario.bien_codigo, to_cronograma_diario_final.unidad_codigo_tec = to_cronograma_diario.unidad_codigo_tec, " & _
        " to_cronograma_diario_final.tec_plan_codigo = to_cronograma_diario.tec_plan_codigo, to_cronograma_diario_final.edif_descripcion = to_cronograma_diario.edif_descripcion, to_cronograma_diario_final.estado_activo = 'APR' FROM to_cronograma_diario_final INNER JOIN to_cronograma_diario " & _
        " ON to_cronograma_diario_final.fmes_plan  = to_cronograma_diario.fmes_plan AND to_cronograma_diario_final.dia_correl  = to_cronograma_diario.dia_correl AND to_cronograma_diario_final.horario_codigo = to_cronograma_diario.horario_codigo WHERE to_cronograma_diario.fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND to_cronograma_diario.estado_activo = 'APR' "

        db.Execute "UPDATE to_cronograma_diario set to_cronograma_diario.estado_codigo   = 'APR' where to_cronograma_diario.fmes_plan  = " & Ado_detalle1.Recordset!fmes_plan & " AND to_cronograma_diario.estado_activo = 'APR' "

        Call ABRIR_TABLA_DET
      End If
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
        VAR_SW = "MOD"
        'tc_zonas_piloto
        Set rs_aux4 = New ADODB.Recordset
        If rs_aux4.State = 1 Then rs_aux4.Close
        rs_aux4.Open "Select * from tc_zonas_piloto where zpiloto_codigo = " & dtc_codigo3.Text & " ", db, adOpenStatic
        If rs_aux4.RecordCount > 0 Then
            dtc_codigo4.Text = rs_aux4!beneficiario_codigo
            dtc_desc4.BoundText = dtc_codigo4.BoundText
        End If
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

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    'Fra_Gestion.Visible = True
    VAR_GES = Year(Date)        'Cmb_gestion.Text
    parametro = Aux
    Call ABRIR_TABLAS_AUX
    db.Execute "update to_cronograma_diario set to_cronograma_diario.edif_descripcion   = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    db.Execute "update to_cronograma_diario set to_cronograma_diario.edif_descripcion   = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    Call OptFilGral1_Click
    
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    Option1.Value = True
    Option3.Value = True
    buscados = 0
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

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (registros no aprobados)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    'queryinicial = "select * From tv_cronograma_mensual_zona WHERE estado_codigo = 'REG' "
    queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' "
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "Select * from to_cronograma_mensual "          'where  unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & glGestion & "' "
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

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub ABRIR_TABLA_DET()
    If Option3.Value = True Then
        Set rs_det1 = New ADODB.Recordset
        If rs_det1.State = 1 Then rs_det1.Close
        rs_det1.Open "select * from to_cronograma_diario where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle1.Recordset = rs_det1
        Set dg_det1.DataSource = Ado_detalle1.Recordset
    End If
    If Option4.Value = True Then
        Set rs_det1 = New ADODB.Recordset
        If rs_det1.State = 1 Then rs_det1.Close
        rs_det1.Open "select * from to_cronograma_diario where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle1.Recordset = rs_det1
        Set dg_det1.DataSource = Ado_detalle1.Recordset
    End If
    If Option1.Value = True Then
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "' and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle2.Recordset = rs_det2
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    End If
    If Option2.Value = True Then
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_detalle2.Recordset = rs_det2
        Set dg_det2.DataSource = Ado_detalle2.Recordset
    End If
End Sub

Private Sub Option1_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option2_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option3_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option4_Click()
    Call ABRIR_TABLA_DET
End Sub

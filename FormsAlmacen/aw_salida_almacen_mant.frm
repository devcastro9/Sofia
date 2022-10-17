VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form aw_salida_almacen_mant 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Almacenes - Salida de Almacen de Insumos"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   10620
   Icon            =   "aw_salida_almacen_mant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   10620
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DETALLE DE INSUMOS"
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
      Height          =   8175
      Left            =   6720
      TabIndex        =   50
      Top             =   720
      Width           =   12645
      Begin VB.Frame FraDet2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos Complementarios"
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
         Height          =   3600
         Left            =   2760
         TabIndex        =   70
         Top             =   2640
         Visible         =   0   'False
         Width           =   7140
         Begin MSDataListLib.DataCombo dtc_desc4A 
            Bindings        =   "aw_salida_almacen_mant.frx":0A02
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   1560
            TabIndex        =   84
            Top             =   2280
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ListField       =   "beneficiario_denominacion"
            BoundColumn     =   "beneficiario_codigo"
            Text            =   "Todos"
         End
         Begin VB.TextBox txt_cm 
            BackColor       =   &H00808080&
            DataField       =   "doc_numero_m"
            DataSource      =   "Ado_detalle2"
            Height          =   285
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txt_hdm 
            BackColor       =   &H00808080&
            DataField       =   "doc_codigo"
            DataSource      =   "Ado_detalle2"
            Height          =   285
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   600
            Width           =   1440
         End
         Begin VB.TextBox txt_obs 
            DataField       =   "observaciones2"
            DataSource      =   "Ado_detalle2"
            Height          =   645
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   1320
            Width           =   6600
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   290
            Left            =   6600
            TabIndex        =   73
            Top             =   2730
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.CommandButton BtnGraba3 
            BackColor       =   &H80000015&
            Caption         =   "Aceptar"
            Height          =   615
            Left            =   2160
            Picture         =   "aw_salida_almacen_mant.frx":0A1B
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Aprueba Registro"
            Top             =   2760
            Width           =   1125
         End
         Begin VB.CommandButton BtnCancelar3 
            BackColor       =   &H80000015&
            Caption         =   "Cancelar"
            Height          =   615
            Left            =   3960
            Picture         =   "aw_salida_almacen_mant.frx":0C25
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   2760
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker DTPEjecucion 
            DataField       =   "fecha_almi"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   81
            Top             =   600
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            Format          =   118423553
            CurrentDate     =   42682
            MaxDate         =   55153
            MinDate         =   32874
         End
         Begin MSDataListLib.DataCombo dtc_codigo4A 
            Bindings        =   "aw_salida_almacen_mant.frx":0E2F
            DataField       =   "beneficiario_codigo_resp"
            DataSource      =   "Ado_datos"
            Height          =   315
            Left            =   240
            TabIndex        =   83
            Top             =   2280
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
         Begin VB.Label Label7 
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
            TabIndex        =   82
            Top             =   2040
            Width           =   2820
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
            TabIndex        =   80
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. de Registro"
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
            Left            =   5280
            TabIndex        =   79
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo de Registro"
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
            Left            =   2880
            TabIndex        =   78
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lbl_campo5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000012&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Salida"
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
            TabIndex        =   77
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   12420
         TabIndex        =   51
         Top             =   240
         Width           =   12420
         Begin MSDataListLib.DataCombo dtc_codigo2 
            Bindings        =   "aw_salida_almacen_mant.frx":0E48
            Height          =   315
            Left            =   6600
            TabIndex        =   52
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
         Begin MSDataListLib.DataCombo dtc_desc2 
            Bindings        =   "aw_salida_almacen_mant.frx":0E63
            Height          =   315
            Left            =   3240
            TabIndex        =   53
            Top             =   240
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "edif_descripcion"
            BoundColumn     =   "edif_codigo"
            Text            =   "Todos"
         End
         Begin VB.PictureBox BtnImprimir4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   10920
            Picture         =   "aw_salida_almacen_mant.frx":0E7E
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   61
            ToolTipText     =   "Estado de Ejecución"
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.PictureBox BtnBuscar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7680
            Picture         =   "aw_salida_almacen_mant.frx":174B
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   9480
            Picture         =   "aw_salida_almacen_mant.frx":1F00
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   59
            ToolTipText     =   "Comprobante de Salida Almacen"
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnAprobarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6360
            Picture         =   "aw_salida_almacen_mant.frx":27CD
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   58
            ToolTipText     =   "Envia a Cobranzas"
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnCancelarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5040
            Picture         =   "aw_salida_almacen_mant.frx":3000
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   57
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
            Left            =   3720
            Picture         =   "aw_salida_almacen_mant.frx":38EC
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   56
            Top             =   0
            Visible         =   0   'False
            Width           =   1280
         End
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            Picture         =   "aw_salida_almacen_mant.frx":40C2
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   55
            Top             =   0
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1680
            Picture         =   "aw_salida_almacen_mant.frx":49D7
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   54
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
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
            Index           =   9
            Left            =   3480
            TabIndex        =   62
            Top             =   0
            Width           =   2250
         End
      End
      Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
         Bindings        =   "aw_salida_almacen_mant.frx":5123
         Height          =   6615
         Left            =   120
         Negotiate       =   -1  'True
         OleObjectBlob   =   "aw_salida_almacen_mant.frx":513E
         TabIndex        =   63
         Top             =   960
         Width           =   12375
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
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
         Left            =   2760
         TabIndex        =   69
         Top             =   7680
         Width           =   765
      End
      Begin VB.Label lblcant1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_detalle2"
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
         Left            =   3600
         TabIndex        =   68
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label lblcant2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_detalle2"
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
         Left            =   4680
         TabIndex        =   67
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label lblcant4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_detalle2"
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
         Left            =   6840
         TabIndex        =   66
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label lblcant5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_detalle2"
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
         Left            =   7920
         TabIndex        =   65
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label lblcant3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataSource      =   "Ado_detalle2"
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
         Left            =   5760
         TabIndex        =   64
         Top             =   7680
         Width           =   1095
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
      TabIndex        =   37
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox fraOpciones 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   20280
         TabIndex        =   41
         Top             =   0
         Width           =   20280
         Begin VB.PictureBox BtnAñadir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            Picture         =   "aw_salida_almacen_mant.frx":1428B
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1305
            Picture         =   "aw_salida_almacen_mant.frx":14A4A
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   47
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2640
            Picture         =   "aw_salida_almacen_mant.frx":1535F
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   46
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox BtnAprobar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   7320
            Picture         =   "aw_salida_almacen_mant.frx":15AAB
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   45
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3720
            Picture         =   "aw_salida_almacen_mant.frx":162DE
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   44
            Top             =   0
            Width           =   1215
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5760
            Picture         =   "aw_salida_almacen_mant.frx":16A93
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   43
            ToolTipText     =   "Salida Almacen - Cronograma"
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
            Picture         =   "aw_salida_almacen_mant.frx":17360
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   42
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
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
            Left            =   12615
            TabIndex        =   49
            Top             =   195
            Width           =   1815
         End
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "aw_salida_almacen_mant.frx":17B22
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   39
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
         Picture         =   "aw_salida_almacen_mant.frx":1840E
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   38
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
         Left            =   12735
         TabIndex        =   40
         Top             =   195
         Width           =   1005
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000040&
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   6540
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6105
         TabIndex        =   32
         Top             =   3660
         Width           =   255
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "aw_salida_almacen_mant.frx":18BE4
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4800
         TabIndex        =   24
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
         Left            =   6135
         TabIndex        =   36
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
         TabIndex        =   33
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
         TabIndex        =   25
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   4440
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "aw_salida_almacen_mant.frx":18BFD
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5640
         TabIndex        =   7
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
         Bindings        =   "aw_salida_almacen_mant.frx":18C16
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   4920
         TabIndex        =   8
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
         Bindings        =   "aw_salida_almacen_mant.frx":18C2F
         DataField       =   "unidad_codigo_tec"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   9
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
         Format          =   118423555
         CurrentDate     =   41678
         MaxDate         =   109939
         MinDate         =   36526
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "aw_salida_almacen_mant.frx":18C48
         DataField       =   "zpiloto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   2220
         Width           =   4605
         _ExtentX        =   8123
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
         Bindings        =   "aw_salida_almacen_mant.frx":18C61
         DataField       =   "beneficiario_codigo_resp"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   240
         TabIndex        =   23
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
         TabIndex        =   35
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.Frame FraNavega 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO"
      ForeColor       =   &H00800000&
      Height          =   3120
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6540
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
         Left            =   2160
         TabIndex        =   87
         Top             =   2835
         Width           =   855
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
         Left            =   960
         TabIndex        =   86
         Top             =   2835
         Value           =   -1  'True
         Width           =   855
      End
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
         TabIndex        =   85
         Top             =   2835
         Width           =   915
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
         TabIndex        =   3
         Top             =   2835
         Width           =   915
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
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   2835.213
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   750.047
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
      Left            =   120
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
      Left            =   5160
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
   Begin MSAdodcLib.Adodc Ado_busqueda 
      Height          =   330
      Left            =   0
      Top             =   9720
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
      Caption         =   "Ado_busqueda"
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
      TabIndex        =   34
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "aw_salida_almacen_mant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_datos As New ADODB.Recordset
Dim rs_datos1 As New ADODB.Recordset
Dim rs_datos3 As New ADODB.Recordset
Dim rs_datos4 As New ADODB.Recordset
Dim rs_busqueda As New ADODB.Recordset

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
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset
Dim rs_aux10 As New ADODB.Recordset

Dim rs_almacen2 As New ADODB.Recordset

Dim puntero As String
'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

Dim VAR_CANT1 As Integer
Dim VAR_CANT2 As Integer
Dim VAR_NUM As Integer
Dim VAR_DOC As String
Dim VAR_DC As Integer
Dim VAR_ALMH, VAR_ALMI As Integer
Dim VAR_CRTL, VAR_CRONO As Integer

'OTROS
'Dim swnuevo As String
Dim imag2 As Long

Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim sqlAux As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim var_cod, VAR_GES As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, VAR_SOLA, VAR_TIT, VAR_B, VAR_ED, VAR_BC  As String
Dim VAR_B1, VAR_B2, VAR_B3, VAR_B4, VAR_B5 As String
Dim VAR_EDIFD, VAR_DA As String
Dim VAR_UORIGEN, VAR_DPTOC As String
Dim VAR_AUX, VAR_CONT2 As Double
Dim VAR_AUX2, VAR_COD0, CONT3, VAR_VT As Integer
Dim DIAS_HAB, NRO_HRS, NRO_HORARIO As Integer
Dim VAR_C1, VAR_C2, VAR_C3, VAR_C4, VAR_C5 As String
Dim VAR_FSAL As Date
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
            Set rs_aux6 = New ADODB.Recordset
            If rs_aux6.State = 1 Then rs_aux6.Close
            rs_aux6.Open "select sum(cantidad1) as cant1, sum(cantidad2) as cant2,sum(cantidad3) as cant3,sum(cantidad4) as cant4, sum(cantidad5) as cant5  from to_cronograma_diario_final where fmes_plan= " & Ado_datos.Recordset!fmes_plan & " ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux6.RecordCount > 0 Then
                lblcant1.Caption = IIf(IsNull(rs_aux6!cant1), 0, rs_aux6!cant1)
                lblcant2.Caption = IIf(IsNull(rs_aux6!cant2), 0, rs_aux6!cant2)
                lblcant3.Caption = IIf(IsNull(rs_aux6!cant3), 0, rs_aux6!cant3)
                lblcant4.Caption = IIf(IsNull(rs_aux6!cant4), 0, rs_aux6!cant4)
                lblcant5.Caption = IIf(IsNull(rs_aux6!cant5), 0, rs_aux6!cant5)
            Else
                lblcant1.Caption = "0"
                lblcant2.Caption = "0"
                lblcant3.Caption = "0"
                lblcant4.Caption = "0"
                lblcant5.Caption = "0"
            End If
            Call ABRIR_TABLA_DET(1)
            If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
                lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
            End If
            'mes2 = MonthName(Month(DTPFec_Inicio.Value))
        End If
     Else
        'Set rs_det1 = New ADODB.Recordset
        Set TDBGrid1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
     End If
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle2.Recordset("estado_almacen") = "APR" Then
      sino = MsgBox("Está Seguro de cambiar a ANULAR LA SALIDA DE ALMACEN ? (Este ya no podrá ser habilitado ...) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        VAR_DC = Ado_detalle2.Recordset!doc_numero_m
        VAR_GES = Year(Ado_detalle2.Recordset!fecha_almacen)
        Ado_detalle2.Recordset!estado_almacen = "ANL"
        Ado_detalle2.Recordset!ok_almacen = "False"
        Ado_detalle2.Recordset.Update
        'Call ABRIR_TABLA_DET
        'rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & "  and bien_codigo = '" & VAR_B5 & "'   ", db, adOpenKeyset, adLockOptimistic
        db.Execute "delete ao_almacen_salidas where ges_gestion = '" & VAR_GES & "' AND almacen_codigo = '2' AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & " "
      End If
   Else
        MsgBox "No se puede ANULAR, el registro NO fue procesado (Estado=REG) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
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

Private Sub BtnCancelar3_Click()
    'fraOpciones.Enabled = True
    ' fraOpciones2.Enabled = True
    ' FrmABMDet.Enabled = True
    ' FraDet2.Visible = False
   dtc_desc2.Enabled = True
   TDBGrid1.Enabled = True
   FraDet2.Visible = False

'   TDBGrid1.AllowUpdate = False
End Sub

Private Sub BtnCancelarDet_Click()
    BtnModDetalle.Visible = True
'    BtnImprimir2.Visible = True
    BtnGrabarDet.Visible = False
    BtnCancelarDet.Visible = False
'    dg_det2.Enabled = False
    TDBGrid1.AllowUpdate = False
End Sub

Private Sub BtnGraba3_Click()
    On Error GoTo UpdateErr
    VAR_DOC = "R-115"
    VAR_DC = IIf(IsNull(Ado_detalle2.Recordset!doc_numero_m), 0, Ado_detalle2.Recordset!doc_numero_m)
    VAR_B = Ado_detalle2.Recordset!bien_codigo
    VAR_ED = Ado_detalle2.Recordset!EDIF_CODIGO
    VAR_VT = Ado_detalle2.Recordset!venta_codigo
    VAR_BC = Ado_detalle2.Recordset!beneficiario_codigo
    If VAR_CRTL = 1 Then
        VAR_FSAL = IIf(IsNull(DTPEjecucion.Value), Date, DTPEjecucion.Value)
        'db.Execute "Update to_cronograma_mensual SET beneficiario_codigo_resp = '" & dtc_codigo4A & "' Where fmes_plan= " & Ado_detalle2.Recordset!fmes_plan & " "
        db.Execute "Update to_cronograma_diario_final SET beneficiario_codigo_resp = '" & dtc_codigo4A & "' Where fmes_plan= " & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' "
    Else
        VAR_FSAL = Date
    End If
    '
    db.Execute "Update to_cronograma_diario_final SET doc_codigo = '" & VAR_DOC & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' "
    'ACTUALIZA CORRELATIVO DE DOC. RESPALDO
    'If Ado_detalle2.Recordset!doc_numero_m = 0 Or IsNull(Ado_detalle2.Recordset!doc_numero_m) Then
    If VAR_DC = 0 Then
        Set rs_aux2 = New ADODB.Recordset
        If rs_aux2.State = 1 Then rs_aux2.Close
        Select Case Left(VAR_ED, 1)
            Case "1"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i1' "
                VAR_ALMI = 2            '13
            Case "2"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i2' "
                VAR_ALMI = 2
            Case "3"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i3' "
                VAR_ALMI = 12
            Case "4"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i3' "
                VAR_ALMI = 12           '14
            Case "5"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i2' "
                VAR_ALMI = 2            '13
            Case "6"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i2' "
                VAR_ALMI = 2            '24
            Case "7"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i7' "
                VAR_ALMI = 11
            Case "8"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i7' "
                VAR_ALMI = 11
            Case "9"
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i7' "
                VAR_ALMI = 11
            Case Else
                SQL_FOR = "select * from fc_Correl where tipo_tramite = 'R-115i2' "
                VAR_ALMI = 2
        End Select
        'SQL_FOR = "select * from ac_almacenes where almacen_codigo = '2'"        ''" & VAR_DOC & "' "
        rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux2.RecordCount > 0 Then
           'rs_aux2!correl_doc = rs_aux2!correl_doc + 1
           'VAR_NUM = rs_aux2!correl_doc        'numero_correlativo
           'rs_aux2!correl_sal = rs_aux2!correl_sal + 1
           'VAR_DC
           rs_aux2!numero_correlativo = rs_aux2!numero_correlativo + 1
           'VAR_NUM = rs_aux2!correl_sal
           VAR_NUM = rs_aux2!numero_correlativo
           VAR_DC = VAR_NUM
           rs_aux2.Update
        Else
            VAR_NUM = VAR_DC
        End If
    Else
        Select Case Left(VAR_ED, 1)
            Case "1"        'CHQ
                VAR_ALMI = 2    '13
            Case "2"        'LPZ
                VAR_ALMI = 2
            Case "3"        'CBB
                VAR_ALMI = 12
            Case "4"        'ORU
                VAR_ALMI = 12   '14
            Case "5"        'PTS
                VAR_ALMI = 2    '13
            Case "6"        'TJA
                VAR_ALMI = 2    '24
            Case "7"        'SCZ
                VAR_ALMI = 11
            Case "8"        'BEN
                VAR_ALMI = 11
            Case "9"        'PDO
                VAR_ALMI = 11
            Case Else
                VAR_ALMI = 2
        End Select
        VAR_NUM = VAR_DC
    End If
    
    If (Ado_detalle2.Recordset!doc_numero_m = 0 Or IsNull(Ado_detalle2.Recordset!doc_numero_m)) Or (glusuario = "ADMIN" Or glusuario = "LVASQUEZ" Or glusuario = "RCUELA") Then
           db.Execute "Update to_cronograma_diario_final SET doc_numero_m = " & VAR_NUM & "  Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
           db.Execute "Update to_cronograma_diario_final SET fecha_almi = '" & VAR_FSAL & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
           db.Execute "Update to_cronograma_diario_final SET ok_almacen = 'True' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
           db.Execute "Update to_cronograma_diario_final SET observaciones2 = '" & txt_obs.Text & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
           db.Execute "Update to_cronograma_diario_final SET estado_almacen = 'APR' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
           db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and unidad_codigo_tec like '%MAN%' "
           'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and unidad_codigo_tec = 'DNMAN' "
           If Ado_datos.Recordset!unidad_codigo_tec = "DNMAN" Or Ado_datos.Recordset!unidad_codigo_tec = "DMANS" Or Ado_datos.Recordset!unidad_codigo_tec = "DMANB" Or Ado_datos.Recordset!unidad_codigo_tec = "DMANC" Then
                VAR_B1 = ""
                VAR_B2 = ""
                VAR_B3 = ""
                VAR_B4 = ""
                VAR_B5 = ""
                VAR_C1 = 0
                VAR_C2 = 0
                VAR_C3 = 0
                VAR_C4 = 0
                VAR_C5 = 0
            'INICIO ACTUALIZA ao_almacen_salidas   'Copia el registro completo
            Set rs_aux8 = New ADODB.Recordset
            If rs_aux8.State = 1 Then rs_aux8.Close
            sqlAux = "SELECT almacen_codigo, doc_numero_m, bien_codigo, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5 FROM to_cronograma_diario_final WHERE doc_numero_m = " & VAR_DC & " AND almacen_codigo='" & VAR_ALMI & "' group by almacen_codigo, doc_numero_m, bien_codigo, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5"
            rs_aux8.Open sqlAux, db, adOpenKeyset, adLockOptimistic
            If rs_aux8.RecordCount > 0 Then
               rs_aux8.MoveFirst
               While Not rs_aux8.EOF
                 VAR_B1 = rs_aux8!bien_codigo1
                 VAR_B2 = rs_aux8!bien_codigo2
                 VAR_B3 = rs_aux8!bien_codigo3
                 VAR_B4 = rs_aux8!bien_codigo4
                 VAR_B5 = IIf(IsNull(rs_aux8!bien_codigo5), "", rs_aux8!bien_codigo5)
                 VAR_C1 = VAR_C1 + IIf(IsNull(rs_aux8!cantidad1), 0, rs_aux8!cantidad1)
                 VAR_C2 = VAR_C2 + IIf(IsNull(rs_aux8!cantidad2), 0, rs_aux8!cantidad2)
                 VAR_C3 = VAR_C3 + IIf(IsNull(rs_aux8!cantidad3), 0, rs_aux8!cantidad3)
                 VAR_C4 = VAR_C4 + IIf(IsNull(rs_aux8!cantidad4), 0, rs_aux8!cantidad4)
                 VAR_C5 = VAR_C5 + IIf(IsNull(rs_aux8!cantidad5), 0, rs_aux8!cantidad5)
                 rs_aux8.MoveNext
               Wend
               VAR_ALMH = VAR_ALMI     '"2"
                    '1
                    'db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
                    " ('" & glGestion & "', " & VAR_ALMH & ", '" & VAR_DOC & "', " & VAR_COD2 & ", '" & VAR_BIEN2 & "', '" & VAR_PROY2 & "', " & correlv & ", '" & VAR_BEN3 & "', '" & Ado_datos.Recordset!fecha_verif & "', '', " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_bs), 0, ado_datos14.Recordset!venta_precio_total_bs) & ", " & IIf(IsNull(ado_datos14.Recordset!venta_precio_total_dol), 0, ado_datos14.Recordset!venta_precio_total_dol) & ", 'REG', '" & Date & "', '" & glusuario & "') "
                    VAR_ALMI = VAR_ALMI
                    Set rs_aux7 = New ADODB.Recordset
                    If rs_aux7.State = 1 Then rs_aux7.Close
                    rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMI & " AND doc_codigo = 'R-115' AND doc_numero=" & VAR_DC & " and bien_codigo = '" & VAR_B1 & "'   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        db.Execute "update ao_almacen_salidas set edif_codigo = '" & VAR_ED & "', beneficiario_codigo = '" & VAR_BC & "', fecha_salida = '" & VAR_FSAL & "', cantidad_salida = " & VAR_C1 & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' and bien_codigo = '" & VAR_B1 & "'  AND doc_numero = " & rs_aux7!doc_numero & " "
                    Else
                        db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo, concepto) VALUES " & _
                            " ('" & glGestion & "', " & VAR_ALMI & " , 'R-115', " & VAR_DC & ", '" & VAR_B1 & "', '" & VAR_ED & "', " & VAR_VT & ",'" & VAR_BC & "', '" & VAR_FSAL & "', '', " & VAR_C1 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "', '" & Ado_detalle2.Recordset!observaciones2 & "') "
                    End If
                    '2
                    Set rs_aux7 = New ADODB.Recordset
                    If rs_aux7.State = 1 Then rs_aux7.Close
                    rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMI & " AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & " and bien_codigo = '" & VAR_B2 & "'   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        'db.Execute "update ao_almacen_salidas set edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "', beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_alm & "', fecha_salida = '" & Ado_datos.Recordset!fecha_verif & "', cantidad_salida = " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = " & ado_datos14.Recordset!doc_numero_alm & " and bien_codigo = '" & VAR_BIEN2 & "'  "
                        db.Execute "update ao_almacen_salidas set edif_codigo = '" & VAR_ED & "', beneficiario_codigo = '" & VAR_BC & "', fecha_salida = '" & VAR_FSAL & "', cantidad_salida = " & VAR_C2 & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' and bien_codigo = '" & VAR_B2 & "'  AND doc_numero = " & rs_aux7!doc_numero & " "
                    Else
                        db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo, concepto) VALUES " & _
                            " ('" & glGestion & "', " & VAR_ALMI & " , 'R-115', " & VAR_DC & ", '" & VAR_B2 & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & VAR_FSAL & "', '', " & VAR_C2 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "', '" & Ado_detalle2.Recordset!observaciones2 & "') "
                    End If
                    '3
                    Set rs_aux7 = New ADODB.Recordset
                    If rs_aux7.State = 1 Then rs_aux7.Close
                    rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMI & " AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & "  and bien_codigo = '" & VAR_B3 & "'   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        'db.Execute "update ao_almacen_salidas set edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "', beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_alm & "', fecha_salida = '" & Ado_datos.Recordset!fecha_verif & "', cantidad_salida = " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = " & ado_datos14.Recordset!doc_numero_alm & " and bien_codigo = '" & VAR_BIEN2 & "'  "
                        db.Execute "update ao_almacen_salidas set edif_codigo = '" & VAR_ED & "', beneficiario_codigo = '" & VAR_BC & "', fecha_salida = '" & VAR_FSAL & "', cantidad_salida = " & VAR_C3 & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' and bien_codigo = '" & VAR_B3 & "' AND doc_numero = " & rs_aux7!doc_numero & " "
                    Else
                        db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo, concepto) VALUES " & _
                            " ('" & glGestion & "', " & VAR_ALMI & " , 'R-115', " & VAR_DC & ", '" & VAR_B3 & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & VAR_FSAL & "', '', " & VAR_C3 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "', '" & Ado_detalle2.Recordset!observaciones2 & "') "
                    End If
                    '4
                    Set rs_aux7 = New ADODB.Recordset
                    If rs_aux7.State = 1 Then rs_aux7.Close
                    rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMI & " AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & " and bien_codigo = '" & VAR_B4 & "'   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        'db.Execute "update ao_almacen_salidas set edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "', beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_alm & "', fecha_salida = '" & Ado_datos.Recordset!fecha_verif & "', cantidad_salida = " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = " & ado_datos14.Recordset!doc_numero_alm & " and bien_codigo = '" & VAR_BIEN2 & "'  "
                        db.Execute "update ao_almacen_salidas set edif_codigo = '" & VAR_ED & "', beneficiario_codigo = '" & VAR_BC & "', fecha_salida = '" & VAR_FSAL & "', cantidad_salida = " & VAR_C4 & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' and bien_codigo = '" & VAR_B4 & "'   AND doc_numero = " & rs_aux7!doc_numero & " "
                    Else
                        db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo, concepto) VALUES " & _
                            " ('" & glGestion & "', " & VAR_ALMI & " , 'R-115', " & VAR_DC & ", '" & VAR_B4 & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & VAR_FSAL & "', '', " & VAR_C4 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "', '" & Ado_detalle2.Recordset!observaciones2 & "') "
                    End If
                    '5
                    Set rs_aux7 = New ADODB.Recordset
                    If rs_aux7.State = 1 Then rs_aux7.Close
                    rs_aux7.Open "Select * from ao_almacen_salidas where ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMI & " AND doc_codigo = 'R-115' AND doc_numero = " & VAR_DC & "  and bien_codigo = '" & VAR_B5 & "'   ", db, adOpenKeyset, adLockOptimistic
                    If rs_aux7.RecordCount > 0 Then
                        'db.Execute "update ao_almacen_salidas set edif_codigo = '" & Ado_datos.Recordset!edif_codigo & "', beneficiario_codigo = '" & Ado_datos.Recordset!beneficiario_codigo_alm & "', fecha_salida = '" & Ado_datos.Recordset!fecha_verif & "', cantidad_salida = " & ado_datos14.Recordset!bien_cantidad_por_empaque & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = " & ado_datos14.Recordset!doc_numero_alm & " and bien_codigo = '" & VAR_BIEN2 & "'  "
                        db.Execute "update ao_almacen_salidas set edif_codigo = '" & VAR_ED & "', beneficiario_codigo = '" & VAR_BC & "', fecha_salida = '" & VAR_FSAL & "', cantidad_salida = " & VAR_C5 & ", fecha_registro = '" & Date & "', usr_codigo = '" & glusuario & "' WHERE ges_gestion = '" & glGestion & "' AND almacen_codigo = " & VAR_ALMH & " AND doc_codigo = 'R-115' and bien_codigo = '" & VAR_B5 & "'  "
                    Else
                        db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, doc_numero, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo, concepto) VALUES " & _
                            " ('" & glGestion & "', " & VAR_ALMI & " , 'R-115', " & VAR_DC & ", '" & VAR_B5 & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & VAR_FSAL & "', '', " & VAR_C5 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "', '" & Ado_detalle2.Recordset!observaciones2 & "') "
                    End If
               '     rs_aux8.MoveNext
               'Wend
               db.Execute "DELETE ao_almacen_salidas where cantidad_salida = '0' AND almacen_codigo = " & VAR_ALMI & " AND doc_numero = " & VAR_DC & " "
             
             VAR_ALMH = VAR_ALMI     '"2"
             Select Case VAR_DA
               Case "1.8"    'Cochabamba
                   'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and (unidad_codigo_tec = 'DNMAN' OR unidad_codigo_tec = 'DMANB') "
               Case "1.7"    'Santa Cruz
                   'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and (unidad_codigo_tec = 'DNMAN' OR unidad_codigo_tec = 'DMANS') "
               Case "1.3"    'La Paz - Tecnico
                   'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and (unidad_codigo_tec = 'DNMAN' ) "
               Case "1.9"    ' Chuquisaca
                   'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and (unidad_codigo_tec = 'DNMAN' OR unidad_codigo_tec = 'DMANC') "
               Case Else    ' TODO
                   'db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and (unidad_codigo_tec = 'DNMAN' ) "
            End Select
             db.Execute "Update to_cronograma_diario_final SET almacen_codigo = " & VAR_ALMI & " Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "' and unidad_codigo_tec like '%MAN%' "
             'ACTUALIZA ao_almacen_totales   'Actualiza en el Almacen Especificado
'             '1
'             Set rs_almacen2 = New ADODB.Recordset
'             If rs_almacen2.State = 1 Then rs_almacen2.Close
'             'rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & VAR_BIEN2 & "' ", db, adOpenKeyset, adLockOptimistic
'             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo1 & "' ", db, adOpenKeyset, adLockOptimistic
'             If rs_almacen2.RecordCount > 0 Then
'                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =" & rs_almacen2!stock_salida & "+ av_acumula_insumos_tot_alm.cantidad1 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo1 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
'                db.Execute "UPDATE ao_almacen_totales SET stock_salida =  (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo1 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo1 & "' "
'             Else
'                 db.Execute "INSERT INTO ao_almacen_totales(almacen_codigo, bien_codigo, stock_ingreso, stock_salida, stock_salida, total_compra_bs, total_venta_bs, utilidad_Bs, total_compra_dol, total_venta_dol, utilidad_dol, estado_codigo, fecha_registro, usr_codigo )  " & _
'                    " VALUES (" & VAR_ALMH & ", '" & Ado_detalle2.Recordset!bien_codigo1 & "', 0, " & Ado_detalle2.Recordset!cantidad1 & ", " & Ado_detalle2.Recordset!cantidad1 & ", 0, '0', '0', 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "' ) "
'             End If
'             '2
'             Set rs_almacen2 = New ADODB.Recordset
'             If rs_almacen2.State = 1 Then rs_almacen2.Close
'             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo2 & "' ", db, adOpenKeyset, adLockOptimistic
'             If rs_almacen2.RecordCount > 0 Then
'                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =" & rs_almacen2!stock_salida & " + av_acumula_insumos_tot_alm.cantidad2 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo2 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
'                db.Execute "UPDATE ao_almacen_totales SET stock_salida =  (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo2 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo2 & "' "
'             Else
'                 db.Execute "INSERT INTO ao_almacen_totales(almacen_codigo, bien_codigo, stock_ingreso, stock_salida, stock_salida, total_compra_bs, total_venta_bs, utilidad_Bs, total_compra_dol, total_venta_dol, utilidad_dol, estado_codigo, fecha_registro, usr_codigo )  " & _
'                    " VALUES (" & VAR_ALMH & ", '" & Ado_detalle2.Recordset!bien_codigo2 & "', 0, " & Ado_detalle2.Recordset!cantidad2 & ", " & Ado_detalle2.Recordset!cantidad2 & ", 0, '0', '0', 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "' ) "
'             End If
'             '3
'             Set rs_almacen2 = New ADODB.Recordset
'             If rs_almacen2.State = 1 Then rs_almacen2.Close
'             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo3 & "' ", db, adOpenKeyset, adLockOptimistic
'             If rs_almacen2.RecordCount > 0 Then
'                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =" & rs_almacen2!stock_salida & " + av_acumula_insumos_tot_alm.cantidad3 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo3 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
'                 db.Execute "UPDATE ao_almacen_totales SET stock_salida = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo3 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo3 & "' "
'             Else
'                 db.Execute "INSERT INTO ao_almacen_totales(almacen_codigo, bien_codigo, stock_ingreso, stock_salida, stock_salida, total_compra_bs, total_venta_bs, utilidad_Bs, total_compra_dol, total_venta_dol, utilidad_dol, estado_codigo, fecha_registro, usr_codigo )  " & _
'                    " VALUES (" & VAR_ALMH & ", '" & Ado_detalle2.Recordset!bien_codigo3 & "', 0, " & Ado_detalle2.Recordset!cantidad3 & ", " & Ado_detalle2.Recordset!cantidad3 & ", 0, '0', '0', 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "' ) "
'             End If
'             '4
'             Set rs_almacen2 = New ADODB.Recordset
'             If rs_almacen2.State = 1 Then rs_almacen2.Close
'             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' ", db, adOpenKeyset, adLockOptimistic
'             If rs_almacen2.RecordCount > 0 Then
'                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =" & rs_almacen2!stock_salida & " + av_acumula_insumos_tot_alm.cantidad4 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo4 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
'                 db.Execute "UPDATE ao_almacen_totales SET stock_salida = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
'             Else
'                 db.Execute "INSERT INTO ao_almacen_totales(almacen_codigo, bien_codigo, stock_ingreso, stock_salida, stock_salida, total_compra_bs, total_venta_bs, utilidad_Bs, total_compra_dol, total_venta_dol, utilidad_dol, estado_codigo, fecha_registro, usr_codigo )  " & _
'                    " VALUES (" & VAR_ALMH & ", '" & Ado_detalle2.Recordset!bien_codigo4 & "', 0, " & Ado_detalle2.Recordset!cantidad4 & ", " & Ado_detalle2.Recordset!cantidad4 & ", 0, '0', '0', 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "' ) "
'             End If
'             '5
'             Set rs_almacen2 = New ADODB.Recordset
'             If rs_almacen2.State = 1 Then rs_almacen2.Close
'             rs_almacen2.Open "select * from ao_almacen_totales where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo5 & "' ", db, adOpenKeyset, adLockOptimistic
'             If rs_almacen2.RecordCount > 0 Then
'                 'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =" & rs_almacen2!stock_salida & " + av_acumula_insumos_tot_alm.cantidad5 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo5 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
'                 db.Execute "UPDATE ao_almacen_totales SET stock_salida = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo5 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo5 & "' "
'             Else
'                 db.Execute "INSERT INTO ao_almacen_totales(almacen_codigo, bien_codigo, stock_ingreso, stock_salida, stock_salida, total_compra_bs, total_venta_bs, utilidad_Bs, total_compra_dol, total_venta_dol, utilidad_dol, estado_codigo, fecha_registro, usr_codigo )  " & _
'                    " VALUES (" & VAR_ALMH & ", '" & Ado_detalle2.Recordset!bien_codigo5 & "', 0, " & Ado_detalle2.Recordset!cantidad5 & ", " & Ado_detalle2.Recordset!cantidad5 & ", 0, '0', '0', 0, 0, 0, 'REG', '" & Date & "', '" & glusuario & "' ) "
'             End If

             db.Execute "update ao_almacen_totales set stock_salida  = tatales_almacenes_js.cantidad_salida FROM tatales_almacenes_js WHERE tatales_almacenes_js.bien_codigo = ao_almacen_totales.bien_codigo and tatales_almacenes_js.almacen_codigo = ao_almacen_totales.almacen_codigo"
             'ACTUALIZA ao_almacen_totales   'Actualiza en el Almacen Especificado
             'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =almacen_totales.stock_salida+ av_acumula_insumos_tot_alm.cantidad1 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo1 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
             'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =almacen_totales.stock_salida+ av_acumula_insumos_tot_alm.cantidad2 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo2 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
             'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =almacen_totales.stock_salida+ av_acumula_insumos_tot_alm.cantidad3 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo3 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
             'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =almacen_totales.stock_salida+ av_acumula_insumos_tot_alm.cantidad4 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo4 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
             'db.Execute "update ao_almacen_totales set ao_almacen_totales.stock_salida =almacen_totales.stock_salida+ av_acumula_insumos_tot_alm.cantidad5 from ao_almacen_totales inner join av_acumula_insumos_tot_alm on ao_almacen_totales.bien_codigo = av_acumula_insumos_tot_alm.bien_codigo5 AND av_acumula_insumos_tot_alm.almacen_codigo = ao_almacen_totales.almacen_codigo"
            
             db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida"
'             db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida where almacen_codigo= '2' and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo2 & "' "
'             db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida where almacen_codigo= '2' and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo3 & "' "
'             db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida where almacen_codigo= '2' and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
'             db.Execute "update ao_almacen_totales set stock_actual = stock_ingreso - stock_salida where almacen_codigo= '2' and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo5 & "' "
           'End If
           
           End If
           'FIN ACTUALIZA ao_almacen_salidas   'Copia el registro completo
'            db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
'                    " ('" & glGestion & "', '2' , " & VAR_DC & ", '" & VAR_B & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & Date & "', '', " & Ado_detalle2.Recordset!cantidad1 & ",'0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "') "
'
'            db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
'                    " ('" & glGestion & "', '2' , " & VAR_DC & ",'" & VAR_B & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & Date & "', '', '" & Ado_detalle2.Recordset!cantidad2 & "','0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "') "
'
'            db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
'                    " ('" & glGestion & "', '2' , " & VAR_DC & ",'" & VAR_B & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & Date & "', '', '" & Ado_detalle2.Recordset!cantidad3 & "','0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "') "
'
'            db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
'                    " ('" & glGestion & "', '2' , " & VAR_DC & ",'" & VAR_B & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & Date & "', '', '" & Ado_detalle2.Recordset!cantidad4 & "','0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "') "
'
'            db.Execute "INSERT INTO ao_almacen_salidas (ges_gestion, almacen_codigo, doc_codigo, bien_codigo, edif_codigo, venta_codigo, beneficiario_codigo, fecha_salida, bien_nro_lote, cantidad_salida, importe_venta_bs, importe_venta_dol, estado_codigo, fecha_registro, usr_codigo) VALUES " & _
'                    " ('" & glGestion & "', '2' , " & VAR_DC & ",'" & VAR_B & "', '" & VAR_ED & "', '" & VAR_VT & "','" & VAR_BC & "', '" & Date & "', '', '" & Ado_detalle2.Recordset!cantidad5 & "','0','0','" & Ado_detalle2.Recordset!estado_almacen & "', '" & Date & "', '" & glusuario & "') "
           
           'ACtualiza ac_bienes    'Todos Los Almacenes
'           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida_mant = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
'           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida_mant = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
'           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida_mant = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
'           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida_mant = (SELECT SUM(cantidad_salida) FROM ao_almacen_salidas WHERE bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' AND almacen_codigo = " & VAR_ALMH & ")where almacen_codigo = " & VAR_ALMH & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo4 & "' "
           db.Execute "update ac_bienes set ac_bienes.bien_stock_salida = total_salidas_js.cantidad_salida from total_salidas_js Where ac_bienes.bien_codigo = total_salidas_js.bien_codigo"
           
           db.Execute "update ac_bienes set bien_stock_actual = bien_stock_ingreso - bien_stock_salida"
        End If
    End If

   
    BtnModDetalle.Visible = True
'   BtnImprimir2.Visible = True
'   BtnGrabarDet.Visible = True
'   BtnCancelarDet.Visible = True
    TDBGrid1.AllowUpdate = True
    TDBGrid1.Enabled = True
    dtc_desc2.Enabled = True
   
    puntero = rs_det2!bien_codigo
    Call ABRIR_TABLA_DET(1)
    'Call BtnImprimir2_Click
'     If OptFilGral1.Value = True Then
'        Call OptFilGral1_Click        'Pendientes
'     Else
'        Call OptFilGral2_Click        'TODOS
'     End If
'     If (dg_datos.SelBookmarks.Count <> 0) Then
'        dg_datos.SelBookmarks.Remove 0
'     End If
'     If Ado_datos.Recordset.RecordCount > 0 And VAR_SW = "MOD" Then
'        rs_datos.Find "fmes_plan = " & VAR_SOLA & "   ", , , 1
'        dg_datos.SelBookmarks.Add (rs_datos.BookMark)
'     Else
'        rs_datos.MoveLast
'     End If

    If (TDBGrid1.SelBookmarks.Count <> 0) Then
        TDBGrid1.SelBookmarks.Remove 0
    End If
   
   If rs_det2.RecordCount > 0 Then
        rs_det2.Find "bien_codigo = '" & puntero & "'", , , 1
        TDBGrid1.SelBookmarks.Add (rs_det2.Bookmark)
   Else
     sino = MsgBox("No se encontro a nadie con ese nombre", vbInformation, "Aviso")
 'Call Carga_Beneficiario(1)
' dtc_buscar_desc.Text = ""
   End If
 
    FraDet2.Visible = False
    
'    db.Execute "update to_cronograma_diario_final set fecha_conformidad = '" & DTPEjecucion.Value & "', nro_fojas = " & txt_hdm.Text & ", doc_numero = " & txt_cm.Text & ", observaciones = '" & txt_obs.Text & "', carta = '" & Cmb_carta.Text & "', doc_numero_carta = '" & Cmb_carta.Text & "' where fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " and dia_correl = " & Ado_detalle2.Recordset!dia_correl & " and bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'"
'    FraDet2.Visible = False
'    BtnImprimir2.Visible = True
  Exit Sub
UpdateErr:
  MsgBox Err.Description

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
VAR_DOC = "R-115"
db.Execute "Update to_cronograma_diario_final SET doc_codigo  =  '" & VAR_DOC & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
'ACTUALIZA CORRELATIVO DE DOC. RESPALDO

If Ado_detalle2.Recordset!doc_numero_m = 0 Then
            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            SQL_FOR = "select * from gc_documentos_respaldo where doc_codigo = '" & VAR_DOC & "' "
            rs_aux2.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2!correl_doc = rs_aux2!correl_doc + 1
                  VAR_NUM = rs_aux2!correl_doc
                rs_aux2.Update
                
                db.Execute "Update to_cronograma_diario_final SET doc_numero_m = " & VAR_NUM & "  Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
                db.Execute "Update to_cronograma_diario_final SET fecha_almi = '" & Date & "' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
            End If
End If

    BtnModDetalle.Visible = True
    BtnImprimir2.Visible = True
    BtnGrabarDet.Visible = True
    BtnCancelarDet.Visible = True
    TDBGrid1.AllowUpdate = True
'    TDBGrid1.Enabled = True
'    dg_det2.Enabled = False
End Sub


Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
'    db.Execute "Update to_cronograma_diario_final SET to_cronograma_diario_final.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    " From to_cronograma_diario_final INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'
'    db.Execute "Update to_cronograma_diario_final set to_cronograma_diario_final.cantidad3 = '0' From to_cronograma_diario_final INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacen_mant_lista.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
          Case "DNINS"
              var_titulo = "Módulo Instalaciones"
          Case "DNAJS"
              var_titulo = "Módulo Ajustes"
          Case "DNMAN"
              var_titulo = "Módulo Almacenes"
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
    If Ado_detalle2.Recordset.RecordCount > 0 Then
      If Ado_detalle2.Recordset!cantidad1 = "0" And Ado_detalle2.Recordset!cantidad2 = "0" And Ado_detalle2.Recordset!cantidad3 = "0" And Ado_detalle2.Recordset!cantidad4 = "0" And Ado_detalle2.Recordset!cantidad5 = "0" Then
         MsgBox "NO se puede imprimir, porque la cantidad de los insumos debe ser mayor a cero, Consulte con el Responsable de Mantenimiento... ", vbInformation, "Atención!"
      Else
        If Ado_detalle2.Recordset!doc_numero_m = 0 Or IsNull(Ado_detalle2.Recordset!doc_numero_m) Then
            MsgBox "No se puede Imprimir. Debe MODIFICAR los datos previamente ...", , "Atención"
        Else
            VAR_CRTL = 2
            txt_obs.Text = "Salida Almacen de Insumos Mantenimiento"
            VAR_EDIFD = Trim(Ado_detalle2.Recordset!edif_descripcion)
            VAR_CRONO = Ado_detalle2.Recordset!fmes_plan
            'Call BtnGraba3_Click
            Dim iResult As Integer
            'Dim co As New ADODB.Command
            If GlBaseDatos = "ADMIN_EMPRESA" Then
                CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacen_mant.rpt"
            Else
                CR02.ReportFileName = App.Path & "\Reportes\Almacenes\ar_salida_almacen_mant_prueba.rpt"
            End If
            CR02.WindowShowPrintSetupBtn = True
            CR02.WindowShowRefreshBtn = True
            CR02.StoredProcParam(0) = VAR_CRONO         'Ado_detalle2.Recordset!fmes_plan
            CR02.StoredProcParam(1) = VAR_EDIFD         'Trim(Ado_detalle2.Recordset!edif_descripcion)
            iResult = CR02.PrintReport
            If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
            CR02.WindowState = crptMaximized
        End If
      End If
    Else
        MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
    End If
End Sub

Private Sub BtnImprimir4_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_cronograma_mensual_ejecucion.rpt"
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
      VAR_TIT = "EJECUCION SERVICIO DE MANTENIMMIENTO"
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & VAR_TIT & "' "
      'CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
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

Private Sub BtnModDetalle_Click()
  If (Ado_detalle2.Recordset!estado_almacen = "REG" Or IsNull(Ado_detalle2.Recordset!estado_almacen)) Or (glusuario = "ADMIN" Or glusuario = "LVASQUEZ" Or glusuario = "RCUELA") Then
    If Ado_detalle2.Recordset!cantidad1 = "0" And Ado_detalle2.Recordset!cantidad2 = "0" And Ado_detalle2.Recordset!cantidad3 = "0" And Ado_detalle2.Recordset!cantidad4 = "0" And Ado_detalle2.Recordset!cantidad5 = "0" Then
         MsgBox "La cantidad de los insumos debe ser mayor a cero, Consulte con el Responsable de Mantenimiento... ", vbInformation, "Atención!"
         Exit Sub
    Else
         VAR_CRTL = 1
         TDBGrid1.Enabled = False
         dtc_desc2.Enabled = False
         FraDet2.Visible = True
         DTPEjecucion.Value = Date
         If (glusuario = "ADMIN" Or glusuario = "LVASQUEZ" Or glusuario = "RCUELA") Then
            DTPEjecucion.Enabled = True
         Else
            DTPEjecucion.Enabled = False
         End If
         txt_obs.Text = "Salida Almacen de Insumos Mantenimiento"
         'Ado_detalle2.Recordset("observaciones2").Value = "Salida Almacen de Insumos Mantenimiento"
        '    db.Execute "Update to_cronograma_diario_final SET txt_obs. =   'Salida Almacen de Insumos Mantenimiento' Where fmes_plan=" & Ado_detalle2.Recordset!fmes_plan & " and edif_descripcion='" & Ado_detalle2.Recordset!edif_descripcion & "'"
        '        BtnModDetalle.Visible = False
        ''        BtnImprimir2.Visible = False
        '        BtnGrabarDet. = True
        '        BtnCancelarDet.Visible = TrueVisible
                TDBGrid1.AllowUpdate = True
        '        dg_det2.Enabled = True
    End If
  Else
     MsgBox "El Registro ya fue Procesado, Elija otro Registro... ", vbInformation, "Atención!"
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



Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_codigo4A_Click(Area As Integer)
    dtc_desc4A.BoundText = dtc_codigo4A.BoundText
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


Private Sub dtc_desc4A_Click(Area As Integer)
    dtc_codigo4A.BoundText = dtc_desc4A.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    VAR_CRTL = 0
    'Fra_Gestion.Visible = True
    VAR_GES = Year(Date)        'Cmb_gestion.Text
    
    Set rs_aux9 = New ADODB.Recordset
    If rs_aux9.State = 1 Then rs_aux8.Close
    rs_aux9.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux9.RecordCount > 0 Then
        usuario2 = rs_aux9!beneficiario_codigo
        VAR_DA = rs_aux9!da_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
    End If
    Select Case VAR_DA
        Case "1.8"    'Cochabamba
            VAR_UORIGEN = "UALMB"
            VAR_DPTOC = "3"
            VAR_ALMI = 12
        Case "1.7"    'Santa Cruz
            VAR_UORIGEN = "UALMS"
            VAR_DPTOC = "7"
            VAR_ALMI = 11
        Case "1.3"    'La Paz - Tecnico
            VAR_UORIGEN = "UALMI"
            VAR_DPTOC = "2"
            VAR_ALMI = 2
        Case "1.9"    ' Chuquisaca
            VAR_UORIGEN = "UALMC"
            VAR_DPTOC = "1"
            VAR_ALMI = 13
        Case Else    ' TODO
            VAR_UORIGEN = "UALMI"
            VAR_DPTOC = "0"
            VAR_ALMI = 2
     End Select
    
    If Aux = "" Then
        Aux = "DNMAN"
    End If
    parametro = Aux
    
    Call ABRIR_TABLAS_AUX
    
'    db.Execute "update to_cronograma_diario_final set to_cronograma_diario_final.carta   = 'NO' WHERE carta IS NULL  "
'
'    db.Execute "UPDATE to_cronograma_diario_final SET to_cronograma_diario_final.beneficiario_codigo_resp = to_cronograma_mensual.beneficiario_codigo_resp, to_cronograma_diario_final.beneficiario_codigo_resp2 = to_cronograma_mensual.beneficiario_codigo_resp From to_cronograma_diario_final " & _
'    " INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final.fmes_plan  = to_cronograma_mensual.fmes_plan) where to_cronograma_diario_final.beneficiario_codigo_resp is null or to_cronograma_diario_final.beneficiario_codigo_resp =''  "
    
    Call OptFilGral1_Click
    
'    var_cod = "0"
'    Set rs_det1 = New ADODB.Recordset
'    If rs_det1.State = 1 Then rs_det1.Close
'    rs_det1.Open "select * from to_cronograma_diario_final where bien_codigo <> ''  ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    'Set Ado_detalle1.Recordset = rs_det1
'    'Set dg_det1.DataSource = Ado_detalle1.Recordset
''    TDBGrid1.Enabled = False
'    If rs_det1.RecordCount > 0 Then
'             rs_det1.MoveFirst
'             While Not rs_det1.EOF
'                rs_det1!hora_registro = "0"
'                If var_cod = rs_det1!bien_codigo Then
'                    rs_det1!hora_registro = "1"
'                End If
'                var_cod = rs_det1!bien_codigo
'                rs_det1.Update
'                rs_det1.MoveNext
'             Wend
'    End If
    If glusuario = "MLLOSA" Then
        BtnModificar.Visible = False
        BtnEliminar.Visible = False
        BtnAprobar.Visible = False
        BtnModDetalle.Visible = False
        BtnGrabarDet.Visible = False
'        BtnGraba3.Visible = False
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
    rs_datos4.Open "select * from rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
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
    '===== Proceso para filtrado general de datos (todos los registros 2019)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
       Case "1.8"    'Cochabamba
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND (ges_gestion = '2019')) "
       Case "1.7"    'Santa Cruz
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33') AND (ges_gestion = '2019')) "
       Case "1.3"    'La Paz - Tecnico
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'17' OR zpiloto_codigo>'27' ) AND (ges_gestion = '2019')) "
       Case "1.9"    ' Chuquisaca
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36') AND (ges_gestion = '2019'))"
       Case Else    ' TODO
           queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2019') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset

End Sub

Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2020)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
       Case "1.8"    'Cochabamba
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND (ges_gestion = '2020')) "
       Case "1.7"    'Santa Cruz
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33') AND (ges_gestion = '2020')) "
       Case "1.3"    'La Paz - Tecnico
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'17' OR zpiloto_codigo>'27' ) AND (ges_gestion = '2020')) "
       Case "1.9"    ' Chuquisaca
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' ) AND (ges_gestion = '2020'))"
       Case Else    ' TODO
           queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2020') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2022)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
       Case "1.8"    'Cochabamba
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND (ges_gestion = '2022')) "
       Case "1.7"    'Santa Cruz
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33') AND (ges_gestion = '2022')) "
       Case "1.3"    'La Paz - Tecnico
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'17' OR zpiloto_codigo>'27' ) AND (ges_gestion = '2022')) "
       Case "1.9"    ' Chuquisaca
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' ) AND (ges_gestion = '2022'))"
       Case Else    ' TODO
           queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2022') "
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
'    TDBGrid1.Enabled = False
        
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
'
Private Sub ABRIR_TABLA_DET(posicion As Integer)
  Select Case posicion
    Case 1
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from tv_cronograma_salida_insumos where fmes_plan = " & Ado_datos.Recordset!fmes_plan & "  AND estado_activo = 'APR' ORDER BY dia_fecha ", db, adOpenKeyset, adLockOptimistic, adCmdText       'and hora_registro = '0'
        'rs_det2.Open "SEECT distinct fmes_plan, dia_correl, bien_orden, bien_codigo, unidad_codigo_tec, tec_plan_codigo, beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha, dia_nombre, nro_total_horas, observaciones, edif_descripcion, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, " & _
        " bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, fecha_carta, fecha_conformidad, fecha_equipo_hdm, nro_fojas, doc_numero , estado_activo, estado_codigo, usr_codigo, fecha_registro, hora_registro  " & _
        " From dbo.to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Sort = "dia_fecha"
        rs_det2.Sort = "dia_fecha, horario_codigo"
        Set Ado_detalle2.Recordset = rs_det2
        Set TDBGrid1.DataSource = Ado_detalle2.Recordset
        Set TDBGrid1.DataSource = rs_det2   ' Ado_detalle2.Recordset
        dtc_codigo2.BoundText = dtc_desc2.BoundText
'        Columns(26).CheckBoxes = True
    
     
    Case 2
        Set rs_busqueda = New ADODB.Recordset
        If rs_busqueda.State = 1 Then rs_busqueda.Close
        rs_busqueda.Open "select * from tv_cronograma_salida_insumos where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR'  and edif_codigo = '" & dtc_codigo2.Text & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText      'and hora_registro = '0'
         rs_busqueda.Sort = "edif_descripcion"
        'rs_det2.Open "SELECT distinct fmes_plan, dia_correl, bien_orden, bien_codigo, unidad_codigo_tec, tec_plan_codigo, beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha, dia_nombre, nro_total_horas, observaciones, edif_descripcion, bien_codigo1, bien_codigo2, bien_codigo3, bien_codigo4, " & _
        " bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, fecha_carta, fecha_conformidad, fecha_equipo_hdm, nro_fojas, doc_numero , estado_activo, estado_codigo, usr_codigo, fecha_registro, hora_registro  " & _
        " From dbo.to_cronograma_diario_final where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'  AND estado_activo = 'APR' ", db, adOpenKeyset, adLockOptimistic, adCmdText
        Set Ado_busqueda.Recordset = rs_busqueda
       ' Set TDBGrid1.DataSource = Ado_busqueda.Recordset
    
        '--------------- buscar
        If (TDBGrid1.SelBookmarks.Count <> 0) Then
                TDBGrid1.SelBookmarks.Remove 0
        End If
        If rs_busqueda.RecordCount > 0 Then
            rs_det2.Find "edif_codigo LIKE '" & dtc_codigo2.Text & "'", , , 1
            TDBGrid1.SelBookmarks.Add (rs_det2.Bookmark)
        Else
            sino = MsgBox("No se encontro edificios con ese nombre", vbInformation, "Atencion!")
            Call ABRIR_TABLA_DET(1)
            dtc_desc2.Text = ""
        End If
  End Select

End Sub

Private Sub OptFilGral3_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2021)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    Select Case VAR_DA
       Case "1.8"    'Cochabamba
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND (ges_gestion = '2021')) "
       Case "1.7"    'Santa Cruz
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33') AND (ges_gestion = '2021')) "
       Case "1.3"    'La Paz - Tecnico
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'17' OR zpiloto_codigo>'27' ) AND (ges_gestion = '2021')) "
       Case "1.9"    ' Chuquisaca
           queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' ) AND (ges_gestion = '2021'))"
       Case Else    ' TODO
           queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2021') "
    End Select
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

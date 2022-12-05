VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_cronograma_mensual_inst 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Instalaciones - Cronograma por Grupo Piloto"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13260
   Icon            =   "tw_cronograma_mensual_inst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   13260
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraDet7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modifica las Fechas para Cronograma Instalación"
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
      Height          =   2160
      Left            =   10200
      TabIndex        =   45
      Top             =   2040
      Visible         =   0   'False
      Width           =   5580
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "0"
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPfecha4 
         DataField       =   "fecha_ini_max"
         Height          =   315
         Left            =   360
         TabIndex        =   71
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128581633
         CurrentDate     =   44890
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   70
         ScaleHeight     =   660
         ScaleWidth      =   5445
         TabIndex        =   52
         Top             =   1440
         Width           =   5450
         Begin VB.PictureBox BtnGrabar7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1200
            Picture         =   "tw_cronograma_mensual_inst.frx":0A02
            ScaleHeight     =   615
            ScaleWidth      =   1305
            TabIndex        =   75
            Top             =   0
            Width           =   1300
         End
         Begin VB.PictureBox BtnCancelar7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2760
            Picture         =   "tw_cronograma_mensual_inst.frx":11D8
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   53
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSComCtl2.DTPicker DTPfecha5 
         DataField       =   "fecha_fin_max"
         Height          =   315
         Left            =   2280
         TabIndex        =   72
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128581633
         CurrentDate     =   45291
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.de Días"
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
         Left            =   4080
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
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
         Left            =   360
         TabIndex        =   47
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final"
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
         Left            =   2280
         TabIndex        =   46
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ORGANIZACION DE EDIFICIOS EN INSTALACION"
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
      Height          =   5175
      Left            =   8520
      TabIndex        =   3
      Top             =   0
      Width           =   10725
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Terminados"
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
         Left            =   6600
         TabIndex        =   29
         Top             =   4920
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pendientes (en Proceso)"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   4920
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.PictureBox fraOpciones3 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   10455
         TabIndex        =   19
         Top             =   240
         Width           =   10455
         Begin VB.PictureBox BtnModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3480
            Picture         =   "tw_cronograma_mensual_inst.frx":1AC4
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   70
            ToolTipText     =   "Modifica Cronograma del Edificio"
            Top             =   20
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnVer2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1800
            Picture         =   "tw_cronograma_mensual_inst.frx":23D9
            ScaleHeight     =   615
            ScaleWidth      =   1575
            TabIndex        =   44
            ToolTipText     =   "Actualiza #Horas y Orden"
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.PictureBox BtnAnlDetalle4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5040
            Picture         =   "tw_cronograma_mensual_inst.frx":33C2
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   21
            ToolTipText     =   "Anula Horario"
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
            Left            =   6480
            Picture         =   "tw_cronograma_mensual_inst.frx":3B0E
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   20
            ToolTipText     =   "Imprime R-302 Origen (Borrador)"
            Top             =   0
            Visible         =   0   'False
            Width           =   1400
         End
         Begin VB.Label lbl_texto0 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   300
            Left            =   1200
            TabIndex        =   80
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#Grupo"
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
            Left            =   150
            TabIndex        =   79
            Top             =   120
            Width           =   885
         End
      End
      Begin TrueOleDBGrid60.TDBGrid dg_det1 
         Bindings        =   "tw_cronograma_mensual_inst.frx":43DB
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "tw_cronograma_mensual_inst.frx":43F6
         TabIndex        =   81
         Top             =   960
         Width           =   10455
      End
   End
   Begin VB.Frame FraInsumos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una Opción para Actualizar INSUMOS ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2400
      Left            =   10560
      TabIndex        =   64
      Top             =   6840
      Visible         =   0   'False
      Width           =   7980
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Para programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC) los insumos 3 y 4."
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
         Left            =   240
         TabIndex        =   69
         Top             =   1080
         Width           =   7335
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Para programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV) los insumos 3 y 4."
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
         Left            =   240
         TabIndex        =   68
         Top             =   600
         Width           =   7575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   7860
         TabIndex        =   65
         Top             =   1680
         Width           =   7860
         Begin VB.PictureBox BtnCancelar8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4320
            Picture         =   "tw_cronograma_mensual_inst.frx":10E02
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   67
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGrabar8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1920
            Picture         =   "tw_cronograma_mensual_inst.frx":116EE
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   66
            Top             =   0
            Width           =   1280
         End
      End
   End
   Begin VB.Frame FraDet6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija una Opción para Enviar Registros..."
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
      Height          =   2400
      Left            =   11160
      TabIndex        =   40
      Top             =   6480
      Visible         =   0   'False
      Width           =   6900
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   6780
         TabIndex        =   54
         Top             =   1680
         Width           =   6780
         Begin VB.PictureBox BtnGrabar6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1920
            Picture         =   "tw_cronograma_mensual_inst.frx":11EDC
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   62
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3360
            Picture         =   "tw_cronograma_mensual_inst.frx":126CA
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   55
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3. Envía SOLO los equipos PENDIENTES del Origen al Destino"
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
         Left            =   480
         TabIndex        =   43
         Top             =   1200
         Width           =   6135
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Envía TODO, a todos los días calendario, incluyendo días NO laborales"
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
         Left            =   480
         TabIndex        =   42
         Top             =   840
         Width           =   6375
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Envía TODO, solo a los Horarios Laborales definidos en el Destino"
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
         Left            =   480
         TabIndex        =   41
         Top             =   480
         Width           =   6015
      End
   End
   Begin VB.Frame FraDet5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija los parámetros para retornar al Crono. Origen..."
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
      Height          =   2760
      Left            =   11760
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   6300
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   6180
         TabIndex        =   56
         Top             =   2040
         Width           =   6180
         Begin VB.PictureBox BtnGrabar5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            Picture         =   "tw_cronograma_mensual_inst.frx":12FB6
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   61
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3000
            Picture         =   "tw_cronograma_mensual_inst.frx":137A4
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   57
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "tw_cronograma_mensual_inst.frx":14090
         Height          =   315
         Left            =   240
         TabIndex        =   34
         Top             =   675
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_descripcion"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DTPfecha2 
         Bindings        =   "tw_cronograma_mensual_inst.frx":140A9
         Height          =   315
         Left            =   840
         TabIndex        =   35
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo DTPfecha3 
         Bindings        =   "tw_cronograma_mensual_inst.frx":140C3
         Height          =   315
         Left            =   3720
         TabIndex        =   39
         Top             =   1440
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ListField       =   "dia_fecha"
         BoundColumn     =   "dia_fecha"
         Text            =   "Todos"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "fecha_ini_max"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   840
         TabIndex        =   76
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128581633
         CurrentDate     =   44890
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "fecha_fin_max"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   3720
         TabIndex        =   77
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   128581633
         CurrentDate     =   45291
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final (hasta...)"
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
         TabIndex        =   38
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha inicial (desde...)"
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
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio..."
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
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame FraDet3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija para cambiar el Número de horas de Servicio ..."
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
      Left            =   12720
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   4860
      Begin VB.CommandButton BtnCancelar2 
         BackColor       =   &H80000015&
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   2760
         Picture         =   "tw_cronograma_mensual_inst.frx":140DD
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancela sin Guardar"
         Top             =   840
         Width           =   1125
      End
      Begin VB.CommandButton BtnGrabar2 
         BackColor       =   &H80000015&
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   960
         Picture         =   "tw_cronograma_mensual_inst.frx":142E7
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Grabar los Datos"
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
         TabIndex        =   14
         Text            =   "tw_cronograma_mensual_inst.frx":144F1
         Top             =   360
         Width           =   645
      End
      Begin VB.ComboBox cmd_campo2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_cronograma_mensual_inst.frx":144F3
         Left            =   3960
         List            =   "tw_cronograma_mensual_inst.frx":14503
         TabIndex        =   13
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   375
         Width           =   1140
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
      Height          =   4200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox fraOpciones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   8280
         TabIndex        =   82
         Top             =   240
         Width           =   8280
         Begin VB.PictureBox BtnSalir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6960
            Picture         =   "tw_cronograma_mensual_inst.frx":14513
            ScaleHeight     =   615
            ScaleWidth      =   1245
            TabIndex        =   83
            ToolTipText     =   "Cierra la Ventana Activa"
            Top             =   0
            Width           =   1245
         End
         Begin VB.PictureBox BtnAñadir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5160
            Picture         =   "tw_cronograma_mensual_inst.frx":14CD5
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox BtnEliminar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6360
            Picture         =   "tw_cronograma_mensual_inst.frx":15494
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   86
            ToolTipText     =   "Anular Cronograma"
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
            Left            =   1200
            Picture         =   "tw_cronograma_mensual_inst.frx":15BE0
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   85
            ToolTipText     =   "Aprueba Cronograma"
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
            Left            =   0
            Picture         =   "tw_cronograma_mensual_inst.frx":16413
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   84
            Top             =   0
            Width           =   1215
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
            Left            =   3375
            TabIndex        =   88
            Top             =   195
            Width           =   1815
         End
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
         Left            =   2160
         TabIndex        =   63
         Top             =   3915
         Visible         =   0   'False
         Width           =   795
      End
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
         Left            =   5520
         TabIndex        =   2
         Top             =   3915
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1035
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   3840
         Width           =   8265
         _ExtentX        =   14579
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
         Bindings        =   "tw_cronograma_mensual_inst.frx":16BC8
         Height          =   2850
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   5027
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "zpiloto_codigo"
            Caption         =   "#.Grupo"
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
            DataField       =   "zpiloto_descripcion"
            Caption         =   "Grupo.Descripcion"
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
            DataField       =   "depto_codigo"
            Caption         =   "Depto.Codigo"
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
            DataField       =   "beneficiario_codigo"
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
         BeginProperty Column05 
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
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
               ColumnWidth     =   3869.858
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
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
      Height          =   2160
      Left            =   11520
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   7140
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   60
         ScaleHeight     =   660
         ScaleWidth      =   7020
         TabIndex        =   58
         Top             =   1440
         Width           =   7020
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3720
            Picture         =   "tw_cronograma_mensual_inst.frx":16BE0
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   60
            Top             =   0
            Width           =   1335
         End
         Begin VB.PictureBox BtnGraba3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2160
            Picture         =   "tw_cronograma_mensual_inst.frx":174CC
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   59
            Top             =   0
            Width           =   1280
         End
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   6600
         TabIndex        =   8
         Top             =   690
         Width           =   270
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "tw_cronograma_mensual_inst.frx":17CBA
         Height          =   315
         Left            =   240
         TabIndex        =   9
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
         Bindings        =   "tw_cronograma_mensual_inst.frx":17CD3
         Height          =   315
         Left            =   5880
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   405
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CRONOGRAMA POR EDIFICIO (Cliente)"
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
      Height          =   5415
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   19215
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver Solo Horarios Libres"
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
         TabIndex        =   30
         Top             =   5160
         Visible         =   0   'False
         Width           =   2475
      End
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
         Left            =   6360
         TabIndex        =   27
         Top             =   5160
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver los Horarios Laborables"
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
         Left            =   360
         TabIndex        =   26
         Top             =   5160
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.PictureBox fraOpciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   75
         ScaleHeight     =   660
         ScaleWidth      =   8400
         TabIndex        =   22
         Top             =   240
         Width           =   8400
         Begin VB.PictureBox BtnAddDetalle3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000015&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6960
            Picture         =   "tw_cronograma_mensual_inst.frx":17CEC
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   89
            Top             =   20
            Width           =   1335
         End
         Begin VB.PictureBox BtnImprimir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   3480
            Picture         =   "tw_cronograma_mensual_inst.frx":18909
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   23
            ToolTipText     =   "Imprime Cronograma Instalaciones"
            Top             =   0
            Width           =   1400
         End
         Begin VB.PictureBox BtnAñadir2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6000
            Picture         =   "tw_cronograma_mensual_inst.frx":191D6
            ScaleHeight     =   615
            ScaleWidth      =   1320
            TabIndex        =   31
            ToolTipText     =   "Habilita Horario (cambia a  HORARIO LABORABLE)"
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.PictureBox BtnModDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2040
            Picture         =   "tw_cronograma_mensual_inst.frx":19AA3
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   25
            ToolTipText     =   "Cambia Estado del Horario"
            Top             =   0
            Visible         =   0   'False
            Width           =   1430
         End
         Begin VB.PictureBox BtnAnlDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4800
            Picture         =   "tw_cronograma_mensual_inst.frx":1A3B8
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   24
            ToolTipText     =   "Anula Horario (cambia a NO LABORABLE)"
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "#Crono."
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
            Left            =   75
            TabIndex        =   78
            Top             =   120
            Width           =   945
         End
         Begin VB.Label lbl_texto2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   300
            Left            =   1080
            TabIndex        =   32
            Top             =   120
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid dg_det2 
         Bindings        =   "tw_cronograma_mensual_inst.frx":1AB04
         Height          =   4425
         Left            =   75
         TabIndex        =   6
         Top             =   960
         Width           =   18960
         _ExtentX        =   33443
         _ExtentY        =   7805
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
         ColumnCount     =   19
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
            DataField       =   "horario_codigo"
            Caption         =   "#.Tarea"
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
            DataField       =   "observaciones"
            Caption         =   "Tarea.Descripcion"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "nro_total_horas"
            Caption         =   "#.Dias"
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
            DataField       =   "hora_ingreso"
            Caption         =   "Fecha.Inicio"
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
            Caption         =   "Fecha.Fin"
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
            DataField       =   "estado_activo"
            Caption         =   "Estado.Tarea"
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
            Caption         =   "Estado.Todo"
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
            DataField       =   "dia_nombre"
            Caption         =   "Nombre.Mes"
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
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
         BeginProperty Column14 
            DataField       =   "cantidad1"
            Caption         =   "Haipe/Trapo"
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
         BeginProperty Column15 
            DataField       =   "cantidad2"
            Caption         =   "Gasolina"
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
            DataField       =   "cantidad3"
            Caption         =   "ISO-680"
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
            DataField       =   "cantidad4"
            Caption         =   "ISO-2050"
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
            DataField       =   "cantidad5"
            Caption         =   "Grasa"
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
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3764.977
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column12 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   629.858
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
   Begin MSAdodcLib.Adodc Ado_datos10 
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
   Begin MSAdodcLib.Adodc Ado_datos9 
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
   Begin MSAdodcLib.Adodc Ado_datos6 
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
      Top             =   9720
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
   Begin VB.PictureBox FrmABMDet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   9360
      ScaleHeight     =   3075
      ScaleWidth      =   1245
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1275
      Begin VB.PictureBox BtnAnlDetalle3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":1AB1F
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   51
         Top             =   2280
         Width           =   1095
      End
      Begin VB.PictureBox BtnModDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":1B56D
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   50
         Top             =   1560
         Width           =   1095
      End
      Begin VB.PictureBox BtnAnlDetalle2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":1BE87
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   49
         Top             =   840
         Width           =   1095
      End
      Begin VB.PictureBox BtnAddDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_cronograma_mensual_inst.frx":1C7B0
         ScaleHeight     =   615
         ScaleWidth      =   1095
         TabIndex        =   48
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos11 
      Height          =   330
      Left            =   15600
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
End
Attribute VB_Name = "tw_cronograma_mensual_inst"
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
Dim rs_datos6 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset
Dim rs_det2 As New ADODB.Recordset

Dim rs_aux0 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux8 As New ADODB.Recordset
Dim rs_aux9 As New ADODB.Recordset      'OK
Dim rs_aux10 As New ADODB.Recordset
Dim rs_aux11 As New ADODB.Recordset
Dim rs_aux12 As New ADODB.Recordset     'OK
Dim rs_aux13 As New ADODB.Recordset     'OK
Dim rs_aux14 As New ADODB.Recordset     'OK

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

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
Dim VAR_EDIF, VAR_EQP As String
Dim VAR_OBS, VAR_EQP2 As String
Dim VAR_ANL, VAR_SW2, VAR_MSG As String
Dim VAR_DA, VAR_UORIGEN, VAR_DPTOC As String
Dim VAR_BENINST, VAR_BENAJST, VAR_BENSUP As String
Dim VAR_LUN, VAR_PRIM, VAR_UNIDCOD As String
Dim MControl, VAR_DESTAREA, VAR_BIEN As String

Dim VAR_AUX, VAR_CONT2 As Double
Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double
Dim VAR_RECORRIDO, VAR_VELOCIDAD As Double

Dim VAR_AUX2, VAR_COD0, CONT3 As Integer
Dim DIAS_HAB, NRO_HRS, NRO_HORARIO As Integer
Dim VAR_ORDEN, VAR_MES, VAR_FMES As Integer
Dim buscados, busca3, VAR_CONT As Integer
Dim VAR_REG, VAR_CANT1 As Integer
Dim VAR_SW0, VAR_PLANID, VAR_SOL As Integer
Dim VAR_DIA, VAR_NRODIAS, VAR_IDTAREA, VAR_PERIODOS As Integer
Dim VAR_PASAJEROS, VAR_PARADAS As Integer

Dim VAR_FECH1, VAR_FECH2 As Date
Dim VAR_FECHAINI, VAR_FECHACTRL, VAR_FCTRLINI, VAR_FCTRLFIN As Date

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
            'VAR_FMES = Ado_datos.Recordset!fmes_plan
            lbl_texto0 = Ado_datos.Recordset!zpiloto_codigo
            buscados = buscados + 1
            If busca3 = 1 Then
                If buscados = 1 Then
                    Call Option3_Click
                    'Call ABRIR_TABLA_DET
                    'If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
'                        lbl_texto2.Caption = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
                    '    lbl_texto3.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                    'End If
                    'mes2 = MonthName(Month(DTPFec_Inicio.Value))
                    buscados = buscados + 1
                End If
            Else
                'Call ABRIR_TABLA_DET
                Call Option3_Click
                'If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
'                    lbl_texto2.Caption = UCase(MonthName(Ado_datos.Recordset!fmes_correl))
                '    lbl_texto3.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
                'End If
                buscados = buscados + 1
            End If
        Else
            
            'Set dg_det1.DataSource = rsNada
            Set dg_det1.DataSource = rsNada
            Set dg_det2.DataSource = rsNada
        End If
        If glusuario = "MLLOSA" Then
            BtnModificar.Visible = False
            BtnEliminar.Visible = False
            BtnAprobar.Visible = False
            BtnAnlDetalle4.Visible = False
            BtnModDetalle.Visible = False
            BtnAnlDetalle.Visible = False
            BtnAñadir2.Visible = False
'            BtnGraba4.Visible = False
            BtnGrabar2.Visible = False
            BtnGraba3.Visible = False
            BtnAddDetalle.Visible = False
            BtnAnlDetalle2.Visible = False
            BtnAddDetalle3.Visible = False
            BtnModDetalle2.Visible = False
            BtnAnlDetalle3.Visible = False
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        Set dg_det2.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub Ado_detalle1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If VAR_SW0 = 1 Then
        VAR_FMES = Ado_detalle1.Recordset!correlativo
        lbl_texto2.Caption = Ado_detalle1.Recordset!correlativo
        If Ado_detalle1.Recordset!estado_activo = "APR" Then
            BtnModificar.Visible = False
            BtnAddDetalle3.Visible = False
        Else
            BtnModificar.Visible = True
            BtnAddDetalle3.Visible = True
        End If
        
        Set rs_det2 = New ADODB.Recordset
        If rs_det2.State = 1 Then rs_det2.Close
        rs_det2.Open "select * from to_cronograma_diario_final_INST where fmes_plan = '" & VAR_FMES & "' and estado_activo <> 'ANL' AND bien_codigo <> '' ORDER BY horario_codigo ", db, adOpenKeyset, adLockOptimistic, adCmdText
        'rs_det2.Sort = "bien_orden"
        Set Ado_detalle2.Recordset = rs_det2
        If Ado_detalle2.Recordset.RecordCount > 0 Then
            Ado_detalle2.Recordset.MoveLast
            Set dg_det2.DataSource = Ado_detalle2.Recordset
            dg_det2.Visible = True
            
            
        Else
            Set dg_det2.DataSource = rsNada
            dg_det2.Visible = False
        End If
    Else
        dg_det2.Visible = False
    End If
End Sub

Private Sub BtnAddDetalle_Click()
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    'GENERA CRONOGRAMA FINAL ITEM x ITEM (INI)
'    fraOpciones.Enabled = False
'    fraOpciones2.Enabled = False
'    FrmABMDet.Enabled = False
'    FraDet3.Visible = True
'    Set rs_aux7 = New ADODB.Recordset
'    If rs_aux7.State = 1 Then rs_aux7.Close
'    rs_aux7.Open "Select * from to_cronograma_detalle WHERE unidad_codigo_tec = '" & Ado_detalle1.Recordset!unidad_codigo_tec & "' and tec_plan_codigo = " & Ado_detalle1.Recordset!tec_plan_codigo & "  and bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  ", db, adOpenStatic
'    If rs_aux7.RecordCount > 0 Then
'        'txtnrohrs.Text = rs_aux7!bien_cantidad_por_empaque
'        'cmd_campo2.Text = rs_aux7!bien_cantidad_por_empaque
'        txtnrohrs.Text = Ado_detalle1.Recordset!nro_total_horas
'        cmd_campo2.Text = Ado_detalle1.Recordset!nro_total_horas
'    Else
'        txtnrohrs.Text = "2"
'        cmd_campo2.Text = "2"
'    End If
'    'GENERA CRONOGRAMA FINAL ITEM x ITEM (FIN)
'  Else
'      MsgBox "No se puede ENVIAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
'  End If
End Sub

Private Sub BtnAddDetalle3_Click()
    'CRONO_INSTALACION()
    If IsNull(Ado_detalle1.Recordset!venta_codigo) Then
        MsgBox "No se puede Generar el Cronograma, debe Verificar los datos del Contrato y/o consulte con el Administrador del Sistema ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    If IsNull(Ado_detalle1.Recordset!fecha_ini_max) Then
        MsgBox "No se puede Generar el Cronograma, debe Verificar los datos del Contrato y/o consulte con el Administrador del Sistema ...", vbExclamation, "Validación de Registro"
        Exit Sub
    End If
    GlEdificio = Ado_detalle1.Recordset!EDIF_CODIGO
    VAR_FECHAINI = Ado_detalle1.Recordset!fecha_ini_max
    VAR_PLANID = Ado_detalle1.Recordset!fmes_plan                 'Ado_detalle1.Recordset!correlativo
    VAR_BENINST = Ado_detalle1.Recordset!beneficiario_codigo        'RESPONSABLE INSTALACION
    VAR_BENAJST = Ado_detalle1.Recordset!beneficiario_codigo_rep    'RESPONSABLE AJUSTE
    VAR_BENSUP = Ado_detalle1.Recordset!beneficiario_codigo_cobr      'SUPERVISOR INSTALACION
    NumComp = Ado_detalle1.Recordset!venta_codigo
    
    db.Execute "update to_cronograma_mensual_inst SET estado_activo = 'ANL' WHERE dia_fecha < '" & VAR_FECHAINI & "' AND fmes_plan = " & VAR_PLANID & "  "
    db.Execute "update to_cronograma_mensual_inst SET bien_codigo = '0' WHERE fmes_plan = " & VAR_PLANID & " AND bien_codigo IS NULL "
    
    'UNIDAD ORIGEN
    Set rs_aux1 = New ADODB.Recordset
    If rs_aux1.State = 1 Then rs_aux1.Close
    rs_aux1.Open "Select * from AO_VENTAS_CABECERA WHERE venta_codigo = " & NumComp & "   ", db, adOpenStatic
    If rs_aux1.RecordCount > 0 Then
        VAR_UNIDCOD = rs_aux1!unidad_codigo
        VAR_SOL = rs_aux1!solicitud_codigo
    Else
        VAR_UNIDCOD = "DVTA"
        VAR_SOL = 0
    End If
    
    'EDIFICIO
    Set rs_aux0 = New ADODB.Recordset
    If rs_aux0.State = 1 Then rs_aux0.Close
    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & GlEdificio & "'   ", db, adOpenStatic
    If rs_aux0.RecordCount > 0 Then
        VAR_EDIF = rs_aux0!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
    End If
    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes

    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
    'mes_inicio_crono
    MControl = UCase(MonthName(Month(VAR_FECHAINI)))
    'MonthName(Month(fecha))
    VAR_FECHACTRL = VAR_FECHAINI
    VAR_FCTRLINI = VAR_FECHACTRL
    VAR_FCTRLFIN = VAR_FECHACTRL - 1
    Set rs_aux9 = New ADODB.Recordset
    rs_aux9.Open "select * from tc_tareas_crono_instalacion  ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux9.RecordCount > 0 Then
        rs_aux9.MoveFirst
        While Not rs_aux9.EOF
            'FECHA, MES Y DIA
            VAR_FCTRLINI = VAR_FCTRLFIN + 1
            VAR_MOD2 = UCase(WeekdayName(Weekday(VAR_FCTRLINI - 1)))
'            If VAR_MOD2 = "SABADO" Or VAR_MOD2 = "SÁBADO" Then
'                VAR_FCTRLINI = VAR_FCTRLINI + 1
'                VAR_FECHACTRL = VAR_FCTRLINI
'                VAR_MOD2 = UCase(WeekdayName(Weekday(VAR_FCTRLINI - 1)))
'            End If
'            If VAR_MOD2 = "DOMINGO" Then
'                VAR_FCTRLINI = VAR_FCTRLINI + 1
'                VAR_FECHACTRL = VAR_FCTRLINI
'                VAR_MOD2 = UCase(WeekdayName(Weekday(VAR_FCTRLINI - 1)))
'            End If
            VAR_DIA = Day(VAR_FECHACTRL)
            VAR_MES = Month(VAR_FECHACTRL)
            MControl = UCase(MonthName(Month(VAR_FCTRLINI)))
            VAR_IDTAREA = rs_aux9!IdTareaInst
            VAR_DESTAREA = rs_aux9!TareaDescripcion

            Set rs_aux7 = New ADODB.Recordset
            If rs_aux7.State = 1 Then rs_aux7.Close
            rs_aux7.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux7.RecordCount > 0 Then
                rs_aux7.MoveFirst
                While Not rs_aux7.EOF
                    VAR_BIEN = rs_aux7!bien_codigo
                    Select Case rs_aux7!cotiza_codigo
                        Case 1
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo1 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo1 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = CDbl(rs_aux8!recorrido_codigo)
                                VAR_VELOCIDAD = CDbl(rs_aux8!vel_equipo_m_s)
                                VAR_PASAJEROS = CDbl(rs_aux8!pasajeros_descripcion)
                                VAR_PARADAS = CDbl(rs_aux8!trafico_num_paradas)
                            End If
                        Case 2
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo2 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = CDbl(rs_aux8!recorrido_codigo)
                                VAR_VELOCIDAD = CDbl(rs_aux8!vel_equipo_m_s)
                                VAR_PASAJEROS = CDbl(rs_aux8!pasajeros_descripcion)
                                VAR_PARADAS = CDbl(rs_aux8!trafico_num_paradas)
                            End If
                        Case 3
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo3 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = CDbl(rs_aux8!recorrido_codigo)
                                VAR_VELOCIDAD = CDbl(rs_aux8!vel_equipo_m_s)
                                VAR_PASAJEROS = CDbl(rs_aux8!pasajeros_descripcion)
                                VAR_PARADAS = CDbl(rs_aux8!trafico_num_paradas)
                            End If
                        Case 4
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo4 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = CDbl(rs_aux8!recorrido_codigo)
                                VAR_VELOCIDAD = CDbl(rs_aux8!vel_equipo_m_s)
                                VAR_PASAJEROS = CDbl(rs_aux8!pasajeros_descripcion)
                                VAR_PARADAS = CDbl(rs_aux8!trafico_num_paradas)
                            End If
                        Case Else
                            Set rs_aux8 = New ADODB.Recordset
                            If rs_aux8.State = 1 Then rs_aux8.Close
                            rs_aux8.Open "select * from av_arreglo1 where unidad_codigo = '" & VAR_UNIDCOD & "' AND solicitud_codigo = " & VAR_SOL & " AND arreglo1 = " & rs_aux7!cotiza_codigo & "  ", db, adOpenKeyset, adLockBatchOptimistic
                            If rs_aux8.RecordCount > 0 Then
                                VAR_RECORRIDO = CDbl(rs_aux8!recorrido_codigo)
                                VAR_VELOCIDAD = CDbl(rs_aux8!vel_equipo_m_s)
                                VAR_PASAJEROS = CDbl(rs_aux8!pasajeros_descripcion)
                                VAR_PARADAS = CDbl(rs_aux8!trafico_num_paradas)
                            End If
                    End Select
                    VAR_NRODIAS = rs_aux9!NroEstimadoDias
                    VAR_PERIODOS = rs_aux9!NroTiempoPeriodos
                    Select Case rs_aux9!IdTareaInst
                        Case 4
                            VAR_NRODIAS = Round((((CDbl(VAR_RECORRIDO) + 1 + 3) * 4) / 6) / 3, 0)
                            VAR_PERIODOS = VAR_NRODIAS * 2
                        Case 10
                            VAR_NRODIAS = Round(CDbl(VAR_PARADAS) / 1.9, 0)
                        Case 15
                            VAR_NRODIAS = Round(Abs(CDbl(VAR_PASAJEROS) * CDbl(VAR_VELOCIDAD) / CDbl(VAR_PASAJEROS) * CDbl(VAR_VELOCIDAD) - 2), 0)
                            If VAR_NRODIAS = 0 Then
                                VAR_NRODIAS = 1
                            End If
                        Case Else
                            VAR_NRODIAS = VAR_NRODIAS
                    End Select
                    VAR_PERIODOS = VAR_NRODIAS * 2
                    'If (VAR_PERIODOS Mod 2) <> 0 Then
                        
                    'Else
                        
                    'End If
                    
                    VAR_FCTRLFIN = VAR_FCTRLINI + VAR_NRODIAS - 1
                    VAR_MOD1 = UCase(WeekdayName(Weekday(VAR_FCTRLFIN - 1)))
                    If VAR_MOD1 = "SABADO" Or VAR_MOD1 = "SÁBADO" Or VAR_MOD1 = "DOMINGO" Then
                        If VAR_NRODIAS >= 1 And VAR_NRODIAS <= 13 Then
                            VAR_FCTRLFIN = VAR_FCTRLFIN + 2
                        End If
                        VAR_MOD1 = UCase(WeekdayName(Weekday(VAR_FCTRLFIN - 1)))
                    Else
                        If VAR_NRODIAS >= 14 And VAR_NRODIAS <= 20 Then
                            VAR_FCTRLFIN = VAR_FCTRLFIN + 4
                        End If
                        If VAR_NRODIAS >= 21 And VAR_NRODIAS <= 27 Then
                            VAR_FCTRLFIN = VAR_FCTRLFIN + 6
                        End If
                        If VAR_NRODIAS >= 28 And VAR_NRODIAS <= 34 Then
                            VAR_FCTRLFIN = VAR_FCTRLFIN + 8
                        End If
                        VAR_MOD1 = UCase(WeekdayName(Weekday(VAR_FCTRLFIN - 1)))
                        If VAR_MOD1 = "SABADO" Or VAR_MOD1 = "SÁBADO" Or VAR_MOD1 = "DOMINGO" Then
                            VAR_FCTRLFIN = VAR_FCTRLFIN + 2
                        End If
                    End If
'                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
'                    Set rs_aux4 = New ADODB.Recordset
'                    If rs_aux4.State = 1 Then rs_aux4.Close
'                    rs_aux4.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_PLANID & " AND bien_codigo = '" & VAR_BIEN & "' AND horario_codigo = " & VAR_IDTAREA & " AND dia_correl = " & VAR_DIA & " ", db, adOpenKeyset, adLockBatchOptimistic
'                    If rs_aux4.RecordCount > 0 Then
'                        'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'                        db.Execute "update to_cronograma_diario_final_INST set unidad_codigo_tec = '" & VAR_UNIDCOD & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = '" & VAR_DESTAREA & "', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & GlEdificio & "' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
'                        db.Execute "update to_cronograma_diario_final_INST set bien_orden = " & VAR_IDTAREA & ", venta_codigo = " & NumComp & " WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "   "
'                        db.Execute "update to_cronograma_diario_final_INST set estado_activo = 'REG' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
'                    Else
'                        db.Execute "INSERT INTO to_cronograma_diario_final_INST (fmes_plan, dia_correl, horario_codigo, bien_orden,     bien_codigo,        unidad_codigo_tec, tec_plan_codigo,     beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha,             dia_nombre,         hora_ingreso,           hora_salida,            nro_total_horas,      observaciones,      edif_descripcion, bien_codigo1, " & _
'                        " bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, nro_fojas, doc_numero, estado_activo, estado_codigo, usr_codigo,      fecha_registro, " & _
'                        " hora_registro, estado_almacen, ok_almacen, doc_codigo, doc_numero_m, observaciones2, almacen_codigo, cite_certificado, estado_certificado, venta_codigo,  edif_codigo) " & _
'                        " VALUES ( " & VAR_PLANID & ",      " & VAR_DIA & ",        " & VAR_IDTAREA & ",        " & VAR_IDTAREA & ", '" & VAR_BIEN & "', '" & VAR_UNIDCOD & "', " & VAR_SOL & ", '" & VAR_BENINST & "',     '" & VAR_BENAJST & "',  '" & VAR_FECHACTRL & "', '" & MControl & "', '" & VAR_FCTRLINI & "', '" & VAR_FCTRLFIN & "', " & VAR_NRODIAS & ", '" & VAR_DESTAREA & "', '" & GlEdificio & "', '4211', " & _
'                        " '479',        '500',          '4529',         '3113',     '0',        '0',        '0',      '0',       '0',   'NO',       '0',            '0',        '0',        'REG',      'REG',          '" & glusuario & "', '" & Date & "',  " & _
'                        " '0',              'REG',      '0',          'R-115',    '0',          '',             '0',            '0',                'REG',          " & NumComp & ", '" & GlEdificio & "'     )"
'
'                    End If
                    'CARGA CRONOGRAMA PRELIMINAR
                    Set rs_aux4 = New ADODB.Recordset
                    If rs_aux4.State = 1 Then rs_aux4.Close
                    'rs_aux4.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_PLANID & " AND bien_codigo = '" & VAR_BIEN & "' AND horario_codigo = " & VAR_IDTAREA & " AND dia_correl = " & VAR_DIA & " ", db, adOpenKeyset, adLockBatchOptimistic
                    rs_aux4.Open "Select * from to_cronograma_mensual_inst WHERE fmes_plan = " & VAR_PLANID & " AND estado_activo <> 'ANL' AND bien_codigo = '0' ORDER BY dia_fecha, horario_codigo ", db, adOpenStatic
                    If rs_aux4.RecordCount > 0 Then
                        VAR_CONT = 1
                        rs_aux4.MoveFirst
                        While VAR_CONT <= VAR_PERIODOS
                            db.Execute "update to_cronograma_mensual_inst set bien_codigo = '" & rs_aux7!bien_codigo & "', IdTareaInst = " & VAR_IDTAREA & "  WHERE fmes_plan = " & VAR_PLANID & " AND dia_fecha = '" & rs_aux4!dia_fecha & "' AND horario_codigo = " & rs_aux4!horario_codigo & "  "
                            VAR_CONT = VAR_CONT + 1
                            rs_aux4.MoveNext
                        Wend
'                        db.Execute "update to_cronograma_mensual_inst set unidad_codigo_tec = '" & VAR_UNIDCOD & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = '" & VAR_DESTAREA & "', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & GlEdificio & "' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
'                        db.Execute "update to_cronograma_mensual_inst set bien_orden = " & VAR_IDTAREA & ", venta_codigo = " & NumComp & " WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "   "
'                        db.Execute "update to_cronograma_mensual_inst set estado_activo = 'REG' WHERE fmes_plan = " & VAR_PLANID & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & VAR_IDTAREA & "  "
                    Else
                        '
'                        db.Execute "INSERT INTO to_cronograma_mensual_inst (fmes_plan, dia_correl, horario_codigo, bien_orden,     bien_codigo,        unidad_codigo_tec, tec_plan_codigo,     beneficiario_codigo_resp, beneficiario_codigo_resp2, dia_fecha,             dia_nombre,         hora_ingreso,           hora_salida,            nro_total_horas,      observaciones,      edif_descripcion, bien_codigo1, " & _
'                        " bien_codigo2, bien_codigo3, bien_codigo4, bien_codigo5, cantidad1, cantidad2, cantidad3, cantidad4, cantidad5, carta, doc_numero_carta, nro_fojas, doc_numero, estado_activo, estado_codigo, usr_codigo,      fecha_registro, " & _
'                        " hora_registro, estado_almacen, ok_almacen, doc_codigo, doc_numero_m, observaciones2, almacen_codigo, cite_certificado, estado_certificado, venta_codigo,  edif_codigo) " & _
'                        " VALUES ( " & VAR_PLANID & ",      " & VAR_DIA & ",        " & VAR_IDTAREA & ",        " & VAR_IDTAREA & ", '" & VAR_BIEN & "', '" & VAR_UNIDCOD & "', " & VAR_SOL & ", '" & VAR_BENINST & "',     '" & VAR_BENAJST & "',  '" & VAR_FECHACTRL & "', '" & MControl & "', '" & VAR_FCTRLINI & "', '" & VAR_FCTRLFIN & "', " & VAR_NRODIAS & ", '" & VAR_DESTAREA & "', '" & GlEdificio & "', '4211', " & _
'                        " '479',        '500',          '4529',         '3113',     '0',        '0',        '0',      '0',       '0',   'NO',       '0',            '0',        '0',        'REG',      'REG',          '" & glusuario & "', '" & Date & "',  " & _
'                        " '0',              'REG',      '0',          'R-115',    '0',          '',             '0',            '0',                'REG',          " & NumComp & ", '" & GlEdificio & "'     )"

                    End If
                    rs_aux7.MoveNext
                Wend
            VAR_FECHACTRL = VAR_FCTRLFIN + 1
            rs_aux9.MoveNext
            End If
        Wend
    End If
End Sub

Private Sub CRONO_INST()
'    VAR_PLANID = Ado_detalle1.Recordset!fmes_plan
'    VAR_LUN = "SI"
'    VAR_PRIM = "SI"
'    VAR_FECHAINI = Ado_detalle1.Recordset!fecha_ini_max
'    MControl = UCase(MonthName(Month(VAR_FECHAINI)))
'    'MonthName(Month(fecha))
'    VAR_FECHACTRL = VAR_FECHAINI
'    VAR_FCTRLINI = VAR_FECHACTRL
'    'VAR_FCTRLFIN = VAR_FECHACTRL - 1
'    db.Execute "update to_cronograma_mensual_inst SET estado_activo = 'ANL' WHERE fecha_ini_max < '" & VAR_FECHAINI & "' AND fmes_plan = " & VAR_PLANID & "  "
'    db.Execute "update to_cronograma_mensual_inst SET bien_codigo = '0' WHERE fmes_plan = " & VAR_PLANID & " AND bien_codigo IS NULL "
'    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'    ' TAREAS DEL CRONO DE INSTALACIONES
'    Set rs_datos9 = New ADODB.Recordset
'    If rs_datos9.State = 1 Then rs_datos6.Close
'    rs_aux9.Open "select * from tc_tareas_crono_instalacion  ", db, adOpenKeyset, adLockBatchOptimistic
'    If rs_datos9.RecordCount > 0 Then
'        rs_datos9.MoveFirst
'        While Not rs_datos6.EOF
'            VAR_CONT = 1
'            Set rs_datos12 = New ADODB.Recordset
'            If rs_datos12.State = 1 Then rs_datos12.Close
'            rs_datos12.Open "Select * from to_cronograma_mensual_inst WHERE fmes_plan = " & VAR_PLANID & " AND estado_activo = 'APR' AND bien_codigo <> '0' ORDER BY dia_fecha, horario_codigo ", db, adOpenStatic
'            If rs_datos12.RecordCount > 0 Then
'                rs_datos12.MoveFirst
'                While Not rs_datos12.EOF
'
'                    rs_datos12.MoveNext
'                Wend
'            rs_datos9.MoveNext
'        Wend
'    End If
'
'
'    Set rs_aux1 = New ADODB.Recordset
'    rs_aux1.Open "Select * from to_cronograma_mensual_inst WHERE fmes_plan = " & VAR_PLANID & "    ", db, adOpenStatic
'    If rs_aux1.RecordCount > 0 Then
'        var_cod5 = rs_aux1.RecordCount
'        rs_aux1.MoveFirst
'        While Not rs_aux1.EOF
'            VAR_AUX2 = rs_aux1!fmes_plan
'            Set rs_aux2 = New ADODB.Recordset
'            If rs_aux2.State = 1 Then rs_aux2.Close
'            'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
'            rs_aux2.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
'            If rs_aux2.RecordCount > 0 Then
'                rs_aux2.MoveFirst
'                While Not rs_aux2.EOF
'                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
'                    Set rs_aux21 = New ADODB.Recordset
'                    If rs_aux21.State = 1 Then rs_aux21.Close
'                    rs_aux21.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  ", db, adOpenKeyset, adLockBatchOptimistic
'                    If rs_aux21.RecordCount > 0 Then
'                        db.Execute "update to_cronograma_diario set unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
'                        db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "   "
'                        db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
'                    Else
'                        Set rs_aux3 = New ADODB.Recordset
'                        If rs_aux3.State = 1 Then rs_aux3.Close
'                        rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux3.RecordCount > 0 Then
'                            rs_aux3.MoveFirst
'                            'If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
'                                'db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                'VAR_COD0 = VAR_COD0 + 1
'                                'CONT3 = 1
'                                db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
'                                'VAR_EMES = "NADA"
'                            'End If
'                        Else
'                            'POR SI NO TIENE fmes_plan
'                        End If
'                    End If
'                    rs_aux2.MoveNext
'                Wend
'            rs_aux1.MoveNext
'            End If
'        Wend
'    End If

End Sub

Private Sub CRONO_MTTO()
'    VAR_PLANID = Ado_detalle1.Recordset!fmes_plan
'
''    Set rs_aux0 = New ADODB.Recordset
''    If rs_aux0.State = 1 Then rs_aux0.Close
''    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & VAR_PROY2 & "'   ", db, adOpenStatic
''    If rs_aux0.RecordCount > 0 Then
''        VAR_EDIF = Ado_datos.Recordset!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
''    End If
'    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
'    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
'    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'    ' jalar ORDEN de tc_zona_piloto_edif
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from tc_zona_piloto_edif_inst WHERE fmes_plan = " & VAR_PLANID & "    ", db, adOpenStatic
'    If rs_datos6.RecordCount > 0 Then
'    '    DIA_ORDEN = rs_datos6!zona_edif_orden
'    'Else
''        Set rs_aux18 = New ADODB.Recordset
''        If rs_aux18.State = 1 Then rs_aux18.Close
''        rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZONA & " ", db, adOpenKeyset, adLockOptimistic
''        If rs_aux18.RecordCount > 0 Then
''            VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
''        Else
''            VAR_ORDEN = 1
''        End If
'
'       db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
'                  " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo) " & _
'                  " VALUES (" & VAR_ZONA & ", '" & VAR_PROY2 & "', '" & gestion0 & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
'                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ")"
'        DIA_ORDEN = "1"
'    End If
'    'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
'    MControl = Ado_datos.Recordset!mes_inicio_crono_tec                     'mes_inicio_crono
'
'    Set rs_aux1 = New ADODB.Recordset
'    'rs_aux1.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
'    rs_aux1.Open "select * from ao_ventas_cobranza_prog where venta_codigo = " & NumComp & "   ", db, adOpenKeyset, adLockBatchOptimistic
'    If rs_aux1.RecordCount > 0 Then
'        var_cod5 = rs_aux1.RecordCount
'        rs_aux1.MoveFirst
'        While Not rs_aux1.EOF
'            VAR_AUX2 = rs_aux1!fmes_plan
'            Set rs_aux2 = New ADODB.Recordset
'            If rs_aux2.State = 1 Then rs_aux2.Close
'            'rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "    ", db, adOpenKeyset, adLockOptimistic
'            rs_aux2.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
'            If rs_aux2.RecordCount > 0 Then
'                rs_aux2.MoveFirst
'                While Not rs_aux2.EOF
'                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
'                    Set rs_aux21 = New ADODB.Recordset
'                    If rs_aux21.State = 1 Then rs_aux21.Close
'                    rs_aux21.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  ", db, adOpenKeyset, adLockBatchOptimistic
'                    If rs_aux21.RecordCount > 0 Then
'                        db.Execute "update to_cronograma_diario set unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
'                        db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "   "
'                        db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
'                    Else
'                        Set rs_aux3 = New ADODB.Recordset
'                        If rs_aux3.State = 1 Then rs_aux3.Close
'                        rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
'                        If rs_aux3.RecordCount > 0 Then
'                            rs_aux3.MoveFirst
'                            'If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
'                                'db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
'                                db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                                'VAR_COD0 = VAR_COD0 + 1
'                                'CONT3 = 1
'                                db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
'                                'VAR_EMES = "NADA"
'                            'End If
'                        Else
'                            'POR SI NO TIENE fmes_plan
'                        End If
'                    End If
'                    rs_aux2.MoveNext
'                Wend
'            rs_aux1.MoveNext
'            End If
'        Wend
'    End If
End Sub


Private Sub BtnAnlDetalle_Click()
  
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    If Ado_detalle2.Recordset!bien_codigo <> "" And Ado_detalle2.Recordset!edif_descripcion <> " " Then
        MsgBox "No se puede ANULAR, porque ya tiene un equipo asignado en este horario ..." & vbCrLf & "Solo puede ANULAR, horarios Libres..", vbExclamation, "Validación de Registro"
    Else
        VAR_SW2 = "ANL"
        sino = MsgBox("Elige SI: para cambiar a HORARIO NO LABORABLE SOLO el registro elegido ..." & vbCrLf & "Elija NO: Para cambiar a HORARIO NO LABORABLE, de acuerdo a los parámetros a elegir ...", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            Ado_detalle2.Recordset!estado_activo = "ANL"
            Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
            Ado_detalle2.Recordset!edif_descripcion = " "
            Ado_detalle2.Recordset.Update
        
    '      Set rs_aux6 = New ADODB.Recordset
    '      If rs_aux6.State = 1 Then rs_aux6.Close
    '      rs_aux6.Open "Select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
    '      If rs_aux6.RecordCount > 0 Then
    '        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and bien_codigo ='' "
    '        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo = 'REG' where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' "
    '        Call ABRIR_TABLA_DET
    '      End If
        Else
            VAR_MSG = "Anular (Marcar como Honario NO Laborable) ..."
            FraDet7.Caption = FraDet7.Caption + VAR_MSG
            FraDet7.Visible = True
            
            fraOpciones.Visible = False
            FrmABMDet.Visible = False
'            FraGrabarCancelar.Visible = False
            fraOpciones2.Visible = False
            
            'dia_fecha Inicial
            Set rs_aux10 = New ADODB.Recordset
            If rs_aux10.State = 1 Then rs_aux10.Close
            rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND estado_activo <> 'ANL' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
            Set Ado_datos10.Recordset = rs_aux10
            If Ado_datos10.Recordset.RecordCount > 0 Then
            End If
        
            'dia_fecha Final
            Set rs_aux11 = New ADODB.Recordset
            If rs_aux11.State = 1 Then rs_aux11.Close
            rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND estado_activo <> 'ANL' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
            Set Ado_datos11.Recordset = rs_aux11
            If Ado_datos11.Recordset.RecordCount > 0 Then
            End If
        
        End If
    End If
        
'    If Ado_detalle2.Recordset("estado_activo") = "REG" Or Ado_detalle2.Recordset("estado_activo") = "APC" Then
'      sino = MsgBox("Está Seguro de cambiar a HORARIO NO LABORABLE ? (Este ya no será considerado en el Cronograma Final - Destino) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between ('" & CDate(DTPfecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "') "
'        Ado_detalle2.Recordset!estado_activo = "ANL"
'        Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
'        Ado_detalle2.Recordset!edif_descripcion = " "
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'      End If
'    Else
'        MsgBox "No se puede ANULAR, el registro ya fue Aprobado (Estado=APR) o ya fue Anulado anteriormente (Estado=ANL)...", vbExclamation, "Validación de Registro"
'    End If
  Else
      MsgBox "No se puede ANULAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If

  'WWWWWWWWWWWWWWWWWWWWW
'  If Ado_detalle2.Recordset.RecordCount > 0 Then
'    If ExisteReg(Ado_detalle2.Recordset!fmes_plan) Then MsgBox "No se puede RETORNAR TODO, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
'    sino = MsgBox("Elige SI: para cambiar a HORARIO NO LABORABLE el registro elegido ..." & vbCrLf & "Elija NO: Para cambiar a HORARIO NO LABORABLE, de acuerdo a los parámetros elegidos a continuación ...", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'        Ado_detalle2.Recordset!estado_activo = "ANL"
'        Ado_detalle2.Recordset!observaciones = "HORARIO NO LABORABLE"
'        Ado_detalle2.Recordset!edif_descripcion = " "
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'    Else
'        'edif_descripcion
'        Set rs_aux9 = New ADODB.Recordset
'        If rs_aux9.State = 1 Then rs_aux9.Close
'        rs_aux9.Open "Select edif_descripcion from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND edif_descripcion <> '' group  by edif_descripcion order by edif_descripcion ", db, adOpenStatic
'        Set Ado_datos9.Recordset = rs_aux9
''        dtc_desc9.BoundText = dtc_codigo9.BoundText
'
'        'dia_fecha Inicial
'        Set rs_aux10 = New ADODB.Recordset
'        If rs_aux10.State = 1 Then rs_aux10.Close
'        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos10.Recordset = rs_aux10
'        If Ado_datos10.Recordset.RecordCount > 0 Then
'        End If
'
'        'dia_fecha Final
'        Set rs_aux11 = New ADODB.Recordset
'        If rs_aux11.State = 1 Then rs_aux11.Close
'        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos11.Recordset = rs_aux11
'        If Ado_datos11.Recordset.RecordCount > 0 Then
'        End If
'        VAR_ANL = "RET"
'        FraDet5.Caption = "Elija los parámetros para retornar al Crono. Origen..."
'        FraDet5.Visible = True
'    End If
'  Else
'        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
'  End If
  'WWWWWWWWWWWWWWWWWWWWW
End Sub

Private Sub BtnAnlDetalle2_Click()
'  'If ExisteReg2(Ado_detalle2.Recordset!fmes_plan, Ado_detalle2.Recordset!bien_codigo) Then MsgBox "No se puede RETORNAR 1, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
'
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'   'If Ado_detalle2.Recordset!estado_codigo = "REG" And Ado_detalle2.Recordset!estado_activo = "APR" Then
'   If Ado_detalle2.Recordset!estado_activo = "APR" Then
'      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Final) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        'db.Execute "update to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
'        'db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & Ado_detalle2.Recordset!fmes_plan & " AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "
'        db.Execute "update to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "           'AND bien_orden = " & Ado_detalle2.Recordset!bien_orden & "
'        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & Ado_detalle2.Recordset!bien_codigo & "'  "         'bien_orden = " & Ado_detalle2.Recordset!bien_orden & " AND
'        Call ABRIR_TABLA_DET
'      End If
'   Else
'        MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
'   End If
'  Else
'      MsgBox "No se puede RETORNAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
'  End If
End Sub

Private Sub BtnAnlDetalle3_Click()
'  If Ado_detalle2.Recordset.RecordCount > 0 Then
'    If ExisteReg(Ado_detalle2.Recordset!fmes_plan) Then MsgBox "No se puede RETORNAR TODO, porque ya existen datos de ejecución...", vbInformation + vbOKOnly, "Atención": Exit Sub
'
'    sino = MsgBox("Elige SI: para RETORNAR TODO el Cronograma DESTINO al ORIGEN..." & vbCrLf & "Elija NO: Para RETORNAR al Cronograma ORIGEN, registros de acuerdo a los parámetros elegidos...", vbYesNo + vbQuestion, "Atención")
'    If sino = vbYes Then
'      Set rs_aux6 = New ADODB.Recordset
'      If rs_aux6.State = 1 Then rs_aux6.Close
'      rs_aux6.Open "Select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
'      If rs_aux6.RecordCount > 0 Then
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' "
'
'        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo = 'REG' where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' "
'
'        Call ABRIR_TABLA_DET
'      End If
'    Else
'        'edif_descripcion
'        Set rs_aux9 = New ADODB.Recordset
'        If rs_aux9.State = 1 Then rs_aux9.Close
'        rs_aux9.Open "Select edif_descripcion from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND edif_descripcion <> '' group  by edif_descripcion order by edif_descripcion ", db, adOpenStatic
'        Set Ado_datos9.Recordset = rs_aux9
''        dtc_desc9.BoundText = dtc_codigo9.BoundText
'
'        'dia_fecha Inicial
'        Set rs_aux10 = New ADODB.Recordset
'        If rs_aux10.State = 1 Then rs_aux10.Close
'        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos10.Recordset = rs_aux10
'        If Ado_datos10.Recordset.RecordCount > 0 Then
'        End If
'
'        'dia_fecha Final
'        Set rs_aux11 = New ADODB.Recordset
'        If rs_aux11.State = 1 Then rs_aux11.Close
'        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND bien_codigo <> '' group  by dia_fecha order by dia_fecha ", db, adOpenStatic
'        Set Ado_datos11.Recordset = rs_aux11
'        If Ado_datos11.Recordset.RecordCount > 0 Then
'        End If
'        VAR_ANL = "RET"
'        FraDet5.Caption = "Elija los parámetros para retornar al Crono. Origen..."
'        FraDet5.Visible = True
'    End If
'  Else
'        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique y vuelva a intentar ...", vbExclamation, "Validación de Registro"
'  End If
'
''  If Ado_datos.Recordset!estado_codigo = "REG" Then
''    'to_cronograma_diario_final_INST
''    Set rs_aux6 = New ADODB.Recordset
''    If rs_aux6.State = 1 Then rs_aux6.Close
''    rs_aux6.Open "Select * from to_cronograma_diario_final_INST where fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND bien_codigo <> '' ", db, adOpenStatic
''    If rs_aux6.RecordCount > 0 Then
''      sino = MsgBox("Está Seguro de RETORNAR TODO ? (Se Retornará TODO el Cronograma DESTINO al ORIGEN) ", vbYesNo + vbQuestion, "Atención")
''      If sino = vbYes Then
''        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', estado_activo = 'REG' WHERE fmes_plan = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "
''
''        db.Execute "UPDATE to_cronograma_diario_inst set estado_codigo   = 'REG' where fmes_plan  = " & Ado_detalle1.Recordset!fmes_plan & " AND estado_activo = 'APR' "
''
''        Call ABRIR_TABLA_DET
''      End If
''    Else
''        MsgBox "NO existen registros en el CRONOGRAMA FINAL (DESTINO), verifique los registros ...", vbExclamation, "Validación de Registro"
''    End If
''  Else
''      MsgBox "No se puede RETORNAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
''  End If
End Sub

Private Function ExisteReg(codigo2 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM to_cronograma_diario_final_INST  WHERE fmes_plan = " & codigo2 & " AND bien_codigo <> '' and (nro_fojas IS NOT NULL) AND (doc_numero IS NOT NULL)  "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

Private Function ExisteReg2(codigo2 As String, codigo3 As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'GlSqlAux = "SELECT Count(*) AS Cuantos2 FROM to_cronograma_diario_final_INST  WHERE fmes_plan = " & codigo2 & " AND bien_codigo = '" & codigo3 & "' and (nro_fojas IS NOT NULL or nro_fojas='0') AND (doc_numero IS NOT NULL or doc_numero ='0')  "
    GlSqlAux = "SELECT Count(*) AS Cuantos2 FROM to_cronograma_diario_final_INST  WHERE fmes_plan = " & codigo2 & " AND bien_codigo = '" & codigo3 & "' and (nro_fojas IS NOT NULL or nro_fojas='0') AND (doc_numero IS NOT NULL or doc_numero ='0')  "
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg2 = rs!Cuantos2 > 0
End Function

Private Sub BtnAnlDetalle4_Click()
' If Ado_datos.Recordset!estado_activo = "REG" Then
'   If Ado_detalle1.Recordset!estado_codigo = "REG" Then
'      sino = MsgBox("Está Seguro de QUITAR el registro ? (Este no será considerado en el Cronograma Elaborado - Origen) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        db.Execute "update to_cronograma_diario_inst set estado_activo = 'ANL', estado_codigo = 'ANL' WHERE fmes_plan = " & VAR_FMES & " AND horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & " AND  bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
'        'db.Execute "update to_cronograma_diario_inst set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & " AND bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "'  "
'        Call ABRIR_TABLA_DET
'      End If
'   Else
'        MsgBox "No se puede ANULAR, el registro ya fue ENVIADO al Cronograma Destino o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
'   End If
' Else
'      MsgBox "No se puede ANULAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
' End If
End Sub

Private Sub BtnAñadir2_Click()
  If Ado_datos.Recordset!estado_codigo = "REG" Then
    VAR_SW2 = "HAB"
    
    sino = MsgBox("Elige SI: para Habilitar a HORARIO LABORABLE SOLO el registro elegido ..." & vbCrLf & "Elija NO: Para Habilitar a HORARIO LABORABLE, de acuerdo a los parámetros a elegir ...", vbYesNo + vbQuestion, "Atención")
    If sino = vbYes Then
        If Ado_detalle2.Recordset("estado_activo") = "ANL" Or Ado_detalle2.Recordset("estado_activo") = "APC" Then
           'sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
           'If sino = vbYes Then
            Ado_detalle2.Recordset!estado_activo = "REG"
            Ado_detalle2.Recordset!observaciones = " "       '"HORARIO LABORABLE"
            Ado_detalle2.Recordset!edif_descripcion = " "
            Ado_detalle2.Recordset.Update
             'Call ABRIR_TABLA_DET
           'End If
        Else
            If (Ado_detalle2.Recordset!bien_codigo = "" And Ado_detalle2.Recordset!tec_plan_codigo = "0") Then
             'sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
             'If sino = vbYes Then
               Ado_detalle2.Recordset!estado_activo = "REG"
               Ado_detalle2.Recordset!observaciones = " "        '"HORARIO LABORABLE"
               Ado_detalle2.Recordset!edif_descripcion = " "
               Ado_detalle2.Recordset.Update
             'End If
            Else
               MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR, APC, APP) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
            End If
        End If
   Else
        VAR_MSG = "Habilitar (Marcar como Honario LABORABLE) ..."
        FraDet7.Caption = FraDet7.Caption + VAR_MSG
        FraDet7.Visible = True
        
        fraOpciones.Visible = False
        FrmABMDet.Visible = False
'        FraGrabarCancelar.Visible = False
        fraOpciones2.Visible = False
        
        'dia_fecha Inicial
        Set rs_aux10 = New ADODB.Recordset
        If rs_aux10.State = 1 Then rs_aux10.Close
        rs_aux10.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos10.Recordset = rs_aux10
        If Ado_datos10.Recordset.RecordCount > 0 Then
        End If
    
        'dia_fecha Final
        Set rs_aux11 = New ADODB.Recordset
        If rs_aux11.State = 1 Then rs_aux11.Close
        rs_aux11.Open "Select dia_fecha from to_cronograma_diario_final_INST where fmes_plan  = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') group  by dia_fecha order by dia_fecha ", db, adOpenStatic
        Set Ado_datos11.Recordset = rs_aux11
        If Ado_datos11.Recordset.RecordCount > 0 Then
        End If
   End If

  Else
      MsgBox "No se puede HABILITAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub


Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
   If Ado_datos.Recordset!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Cronograma Nro. " + Str(VAR_FMES), vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
'        VAR_COD0 = 0
'        VAR_EQP2 = ""
'        VAR_CONT = 0
'        Set rs_aux12 = New ADODB.Recordset
'        If rs_aux12.State = 1 Then rs_aux12.Close
'        'SQL_FOR = "SELECT to_cronograma_mensual.fmes_plan, to_cronograma_mensual.ges_gestion, to_cronograma_mensual.fmes_correl, to_cronograma_mensual.zpiloto_codigo, to_cronograma_diario_final_INST.dia_correl, to_cronograma_diario_final_INST.horario_codigo, to_cronograma_diario_final_INST.bien_codigo FROM to_cronograma_mensual INNER JOIN to_cronograma_diario_final_INST ON to_cronograma_mensual.fmes_plan = to_cronograma_diario_final_INST.fmes_plan where (to_cronograma_mensual.fmes_plan = " & VAR_FMES & " AND bien_codigo <> '') ORDER BY to_cronograma_diario_final_INST.dia_correl, to_cronograma_diario_final_INST.horario_codigo"
'        SQL_FOR = "SELECT * FROM tv_cronograma_mensual_y_final where (fmes_plan = " & VAR_FMES & " AND bien_codigo <> '') ORDER BY dia_correl, horario_codigo"
'        rs_aux12.Open SQL_FOR, db, adOpenStatic  'group  by bien_codigo
'        VAR_CONT = rs_aux12.RecordCount
'        Set rs_aux13 = New ADODB.Recordset
'        If rs_aux13.State = 1 Then rs_aux13.Close
'        SQL_FOR = "SELECT fmes_plan, ges_gestion, fmes_correl, zpiloto_codigo  FROM to_cronograma_mensual where ges_gestion = '" & rs_aux12!ges_gestion & "' AND fmes_correl = " & rs_aux12!fmes_correl & " + 1 AND zpiloto_codigo = " & rs_aux12!zpiloto_codigo & " "
'        rs_aux13.Open SQL_FOR, db, adOpenStatic  'group  by bien_codigo
'        If rs_aux13.RecordCount > 0 Then
'           db.Execute "update to_cronograma_diario_inst set bien_orden = " & VAR_CONT & " where fmes_plan = " & rs_aux13!fmes_plan & "  aND (bien_codigo <> '' AND bien_orden <= " & VAR_CONT & " ) "
'           db.Execute "update tc_zona_piloto_edif set zona_edif_orden = " & VAR_CONT & " where zpiloto_codigo = " & rs_aux12!zpiloto_codigo & " AND zona_edif_orden <= " & VAR_CONT & "  "
'           If rs_aux12.RecordCount > 0 Then
'              rs_aux12.MoveFirst
'              While Not rs_aux12.EOF
'                If VAR_EQP2 <> rs_aux12!bien_codigo Then
'                  VAR_COD0 = VAR_COD0 + 1
'                  db.Execute "update to_cronograma_diario_inst set bien_orden = " & VAR_COD0 & " where fmes_plan = " & rs_aux13!fmes_plan & " AND bien_codigo = '" & rs_aux12!bien_codigo & "' "
'                  db.Execute "update tc_zona_piloto_edif set zona_edif_orden = " & VAR_COD0 & " where zpiloto_codigo = " & rs_aux12!zpiloto_codigo & "  and edif_codigo = '" & rs_aux12!edif_codigo & "' "
'                End If
'                VAR_EQP2 = rs_aux12!bien_codigo
'                rs_aux12.MoveNext
'              Wend
'           End If
'        End If
      End If
       Ado_datos.Recordset!estado_codigo = "APR"
       Ado_datos.Recordset!estado_activo = "APR"
       Ado_datos.Recordset!fecha_registro = Date
       Ado_datos.Recordset!usr_codigo = glusuario
       Ado_datos.Recordset.Update
   End If
    
   Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnBuscar_Click()
    
    If Ado_datos.Recordset.RecordCount > 0 Then
        buscados = 0
        busca3 = 1
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

Private Sub BtnCancelar2_Click()
'    fraOpciones.Enabled = True
'     fraOpciones2.Enabled = True
'     FrmABMDet.Enabled = True
'     FraDet3.Visible = False
'     cmd_campo2.Text = "2"
End Sub

Private Sub BtnCancelar3_Click()
'    fraOpciones.Enabled = True
'     fraOpciones2.Enabled = True
'     FrmABMDet.Enabled = True
'     FraDet2.Visible = False
End Sub

Private Sub BtnCancelar5_Click()
'     fraOpciones.Enabled = True
'     fraOpciones2.Enabled = True
'     FrmABMDet.Enabled = True
'     FraDet2.Visible = False
'     FraDet5.Visible = False
End Sub

Private Sub BtnCancelar6_Click()
'    FraDet6.Visible = False
End Sub

Private Sub BtnCancelar7_Click()
'    FraDet7.Visible = False
'    VAR_SW2 = ""
'    VAR_MSG = ""
'    fraOpciones.Visible = True
'    FrmABMDet.Visible = True
'    FraGrabarCancelar.Visible = True
'    fraOpciones2.Visible = True
End Sub

Private Sub BtnCancelar8_Click()
'    FraInsumos.Visible = False
End Sub

Private Sub BtnGraba3_Click()
'   'CCCCCCCCCCCCCCCCCCCCCCCCCCCBBBBBBBBBBBBBBB
'   VAR_ZONA = dtc_codigo5.Text
'   VAR_MES = lbl_texto1.Caption
'   gestion0 = txt_codigo1.Text
'
'     Set rs_aux4 = New ADODB.Recordset
'     If rs_aux4.State = 1 Then rs_aux4.Close
'     rs_aux4.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
'     If rs_aux4.RecordCount > 0 Then
'        If rs_aux4!estado_codigo = "APR" Then
'            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
'            Exit Sub
'        End If
'        VAR_UNITEC = Ado_detalle1.Recordset!unidad_codigo_tec
'        VAR_EQP = Ado_detalle1.Recordset!bien_codigo
''        VAR_FMES = Ado_detalle1.Recordset!fmes_plan
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from to_cronograma_mensual where ges_gestion = '" & gestion0 & "' and fmes_correl = " & VAR_MES & " and zpiloto_codigo = " & VAR_ZONA & "  and unidad_codigo_tec = '" & VAR_UNITEC & "'   ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 Then
'             VAR_AUX2 = rs_aux2!fmes_plan
'             VAR_COD0 = 0
'             'db.Execute "SELECT VAR_ORDEN = isnull(max(bien_orden),0) from to_cronograma_diario_inst WHERE     (fmes_plan = " & VAR_AUX2 & " ) "
'            Set rs_aux5 = New ADODB.Recordset
'            If rs_aux5.State = 1 Then rs_aux5.Close
'            rs_aux5.Open "select isnull(max(bien_orden),0) as bien_orden2 from to_cronograma_diario_inst WHERE fmes_plan = " & VAR_AUX2 & "  ", db, adOpenStatic
'            If rs_aux5.RecordCount > 0 Then
'               VAR_ORDEN = rs_aux5!bien_orden2 + 1
'            End If
'             Set rs_aux3 = New ADODB.Recordset
'             If rs_aux3.State = 1 Then rs_aux3.Close
'             rs_aux3.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
'             If rs_aux3.RecordCount > 0 Then
'                 rs_aux3.MoveFirst
'                 While Not rs_aux3.EOF
'                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then
'                        db.Execute "update to_cronograma_diario_inst set bien_codigo = '" & rs_aux4!bien_codigo & "', unidad_codigo_tec = '" & rs_aux4!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux4!tec_plan_codigo & ", observaciones = '" & rs_aux4!observaciones & "', bien_orden = " & VAR_ORDEN & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'REG', estado_activo = 'REG', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', observaciones = 'HORARIO LABORABLE'  WHERE fmes_plan = " & VAR_FMES & " AND dia_correl = " & rs_aux4!dia_correl & " AND horario_codigo = " & rs_aux4!horario_codigo & "  "
'                        VAR_COD0 = VAR_COD0 + 1
'                        CONT3 = 1
'                    End If
'                    rs_aux3.MoveNext
'                    'Habilitar .....
'                    'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'                 Wend
'             End If
'             db.Execute "update to_cronograma_diario_inst set bien_codigo = '', unidad_codigo_tec = '',  tec_plan_codigo = 0, observaciones = '', bien_orden = 0, estado_activo = 'REG', edif_descripcion = '' WHERE fmes_plan = " & VAR_FMES & " AND bien_codigo = '" & VAR_EQP & "'  "
'        End If
'     End If
'     db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_inst INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_inst.bien_codigo  = av_bienes_vs_edificios.bien_codigo "
'     Call ABRIR_TABLA_DET
'    fraOpciones.Enabled = True
'    fraOpciones2.Enabled = True
'    FrmABMDet.Enabled = True
'    FraDet2.Visible = False
End Sub

'Private Sub valida_campos()
'  'Valida compos para editables
''  If (dtc_codigo1.Text = "") Then
''    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
''    VAR_VAL = "ERR"
''    Exit Sub
''  End If
''  If (dtc_codigo3.Text = "") Then
''    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
''    VAR_VAL = "ERR"
''    Exit Sub
''  End If
'  If (dtc_codigo4 = "") Then
'    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
''  If (Txt_campo2.Text = "") Then
''    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
''    VAR_VAL = "ERR"
''    Exit Sub
''  End If
'
'End Sub

Private Sub BtnGrabar2_Click()
'     'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO
'     Set rs_aux2 = New ADODB.Recordset
'     If rs_aux2.State = 1 Then rs_aux2.Close
'     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "   ", db, adOpenKeyset, adLockOptimistic
'     If rs_aux2.RecordCount > 0 Then
'        If rs_aux2!estado_codigo = "APR" Then
'            MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
'            Exit Sub
'        End If
'         VAR_AUX2 = rs_aux2!fmes_plan
'         VAR_COD0 = 0
'         Set rs_aux3 = New ADODB.Recordset
'         If rs_aux3.State = 1 Then rs_aux3.Close
'         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
'         rs_aux3.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_AUX2 & "   ", db, adOpenKeyset, adLockBatchOptimistic
'         If rs_aux3.RecordCount > 0 Then
'             rs_aux3.MoveFirst
'             While Not rs_aux3.EOF
'                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
'                If cmd_campo2.Text > 2 Then
'                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
'                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
'                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
'                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                        VAR_COD0 = VAR_COD0 + 2
'                        CONT3 = 1
'                   End If
'                Else
'                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
'                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
'                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
'                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
'                        VAR_COD0 = VAR_COD0 + 1
'                        CONT3 = 1
'                   End If
'                End If
''                   If cmd_campo2.Text = "4" Then
''                      rs_aux3.MoveNext
''                      If VAR_COD0 < 2 And rs_aux3!estado_activo = "REG" Then        '
''                         'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
''                         db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
''                         db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         'db.Execute "update to_cronograma_diario_final_INST set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         'db.Execute "update to_cronograma_diario_final_INST set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         VAR_COD0 = VAR_COD0 + 1
''                         CONT3 = 1
''                      End If
''                   End If
''                   If cmd_campo2.Text = "8" Then
''                      rs_aux3.MoveNext
''                      If VAR_COD0 < 2 And rs_aux3!estado_activo = "REG" Then        '
''                         'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
''                         db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
''                         db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         'db.Execute "update to_cronograma_diario_final_INST set bien_orden = " & rs_aux2!bien_orden & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         'db.Execute "update to_cronograma_diario_final_INST set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
''                         VAR_COD0 = VAR_COD0 + 1
''                         CONT3 = 1
''                      End If
''                   End If
'                rs_aux3.MoveNext
'                'Habilitar .....
'                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
'             Wend
'         End If
'     End If
'     db.Execute "update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final_INST INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final_INST.bien_codigo  = av_bienes_vs_edificios.bien_codigo where to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " AND to_cronograma_diario_final_INST.bien_codigo <>'' "
'
'    'Actualiza Codigos de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final_INST.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'    'Actualiza Cantidad de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final_INST.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " "
'    'Quita Cantidad de Insumo3 en meses pares al Crono Final
''    db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0' From to_cronograma_diario_final_INST INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final_INST.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final_INST.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
''    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
'    'Actualiza Carta al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_insumos.carta " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'
''     Call BtnAñadir2_Click
'     fraOpciones.Enabled = True
'     fraOpciones2.Enabled = True
'     FrmABMDet.Enabled = True
'     FraDet3.Visible = False
'     cmd_campo2.Text = "2"
'     Call ABRIR_TABLA_DET
'    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
End Sub

Private Sub COPIA_TODOS()
     'WWWWW GENERA TODO EL CRONOGRAMA DIARIO FINAL DESDE ORIGEN
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and bien_codigo <> '' order by bien_orden  ", db, adOpenKeyset, adLockOptimistic       'and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "
     If rs_aux2.RecordCount > 0 Then
        'If rs_aux2!estado_codigo = "APR" Then
        '    MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
        '    Exit Sub
        'End If
       VAR_AUX2 = rs_aux2!fmes_plan
       rs_aux2.MoveFirst
       While Not rs_aux2.EOF
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_AUX2 & "  and estado_codigo = 'REG' ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                If cmd_campo2.Text > 2 Then
                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 2
                        CONT3 = 1
                   End If
                Else
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                   End If
                End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
        rs_aux2.MoveNext
       Wend
     End If
     db.Execute "update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final_INST INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final_INST.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    
    'Actualiza Codigos de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final_INST.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final_INST.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " "
    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0' From to_cronograma_diario_final_INST INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final_INST.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final_INST.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_insumos.carta " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'     Call BtnAñadir2_Click
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
    'wwwwwwwwwwwwwwwwwwwwwwwwwwwww
End Sub

Private Sub COPIA_ALGUNOS()
     'WWWWW GENERA ALGUNOS EL CRONOGRAMA DIARIO FINAL DESDE ORIGEN
     Set rs_aux2 = New ADODB.Recordset
     If rs_aux2.State = 1 Then rs_aux2.Close
     rs_aux2.Open "select * from to_cronograma_diario_inst where fmes_plan = " & VAR_FMES & " and bien_codigo <> '' AND estado_activo = 'REG' order by bien_orden  ", db, adOpenKeyset, adLockOptimistic       'and dia_correl = " & Ado_detalle1.Recordset!dia_correl & " and horario_codigo = " & Ado_detalle1.Recordset!horario_codigo & "
     If rs_aux2.RecordCount > 0 Then
        'If rs_aux2!estado_codigo = "APR" Then
        '    MsgBox "El registro ya fue ENVIADO, debe elegir otro registro ...", vbExclamation, "Validación de Registro"
        '    Exit Sub
        'End If
       VAR_AUX2 = rs_aux2!fmes_plan
       rs_aux2.MoveFirst
       While Not rs_aux2.EOF
         VAR_COD0 = 0
         Set rs_aux3 = New ADODB.Recordset
         If rs_aux3.State = 1 Then rs_aux3.Close
         'rs_aux3.Open "select * from to_cronograma_detalle where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   ", db, adOpenKeyset, adLockBatchOptimistic
         rs_aux3.Open "select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_AUX2 & " and estado_codigo = 'REG' ", db, adOpenKeyset, adLockBatchOptimistic
         If rs_aux3.RecordCount > 0 Then
             rs_aux3.MoveFirst
             While Not rs_aux3.EOF
                'If cmb_dia.Text = "AUTOMATICO" And dtc_codigo5.Text = "0" Then
                If cmd_campo2.Text > 2 Then
                   If VAR_COD0 < cmd_campo2.Text And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 2
                        CONT3 = 1
                   End If
                Else
                    If VAR_COD0 < 1 And rs_aux3!estado_activo = "REG" Then        '
                        'db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux2!dia_correl & " AND horario_codigo = " & rs_aux2!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario_inst set estado_codigo = 'APR', estado_activo = 'APR'  WHERE fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  "
                        db.Execute "update to_cronograma_diario_final_INST set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & rs_aux2!unidad_codigo_tec & "',  tec_plan_codigo = " & rs_aux2!tec_plan_codigo & ", bien_orden = " & rs_aux2!bien_orden & ", estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                        VAR_COD0 = VAR_COD0 + 1
                        CONT3 = 1
                   End If
                End If
                rs_aux3.MoveNext
                'Habilitar .....
                'db.Execute "Update to_cronograma Set estado_detalle = 'APR' Where unidad_codigo_tec = '" & VAR_UNITEC & "' and tec_plan_codigo = " & VAR_TECCOD & "   "
             Wend
         End If
        rs_aux2.MoveNext
       Wend
     End If
     db.Execute "update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.edif_descripcion = av_bienes_vs_edificios.edif_descripcion FROM to_cronograma_diario_final_INST INNER JOIN av_bienes_vs_edificios ON to_cronograma_diario_final_INST.bien_codigo  = av_bienes_vs_edificios.bien_codigo"
    
    'Actualiza Codigos de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1 = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final_INST.bien_codigo2 = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3 = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4 = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
    'Actualiza Cantidad de Insumos al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final_INST.bien_codigo = '" & Ado_detalle1.Recordset!bien_codigo & "' AND to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " "
    'Quita Cantidad de Insumo3 en meses pares al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0' From to_cronograma_diario_final_INST INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final_INST.fmes_plan = to_cronograma_mensual.fmes_plan AND to_cronograma_diario_final_INST.bien_codigo  = to_cronograma_mensual.bien_codigo) " & _
'    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
    'Actualiza Carta al Crono Final
    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_insumos.carta " & _
    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

'     Call BtnAñadir2_Click
    'WWWWW GENERA CRONOGRAMA DIARIO UNO POR UNO (FIN)
End Sub

Private Sub BtnGrabar5_Click()
'
'    If DTPfecha2.Text = "Todos" Then
'        DTPfecha2.Text = "01" & "/" & Trim(Ado_datos.Recordset!fmes_correl) & "/" & Trim(Ado_datos.Recordset!ges_gestion)
'    End If
'    If DTPfecha3.Text = "Todos" Then
'        DTPfecha3.Text = Trim(Ado_datos.Recordset!fmes_nro_dias) & "/" & Trim(Ado_datos.Recordset!fmes_correl) & "/" & Trim(Ado_datos.Recordset!ges_gestion)
'    End If
'    If dtc_desc9.Text = "Todos" And DTPfecha2.Text <> "Todos" Then
'        VAR_FECH1 = CDate(DTPfecha2.Text)
'        VAR_FECH2 = CDate(DTPfecha3.Text)
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = '00' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '')"
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = 'OK' FROM to_cronograma_diario_inst INNER JOIN to_cronograma_diario_final_INST ON to_cronograma_diario_inst.fmes_plan = to_cronograma_diario_final_INST.fmes_plan and to_cronograma_diario_inst.bien_codigo  = to_cronograma_diario_final_INST.bien_codigo WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') "
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.estado_activo  = 'REG', to_cronograma_diario_inst.estado_codigo  = 'REG' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') AND (to_cronograma_diario_inst.hora_registro  = '00')"
'
'        'db.Execute "UPDATE to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG'  where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' and dia_fecha between '" & CDate(dtpFecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "' "
'
'        Call ABRIR_TABLA_DET
'        'cod_comp between " & Val(Me.cboaprob_inicio.Text) & " and " & Val(Me.cbo_aprob_final.Text) & "
'    Else
'        VAR_FECH1 = CDate(DTPfecha2.Text)
'        VAR_FECH2 = CDate(DTPfecha3.Text)
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET bien_orden  = '0', bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = '0', edif_descripcion = '', observaciones = '', estado_activo = 'REG' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
'
'        'db.Execute "UPDATE to_cronograma_diario_inst set estado_activo = 'REG', estado_codigo = 'REG'  where fmes_plan  = " & VAR_FMES & " AND estado_activo = 'APR' and trim(edif_descripcion) = '" & Trim(dtc_desc9.Text) & "' and dia_fecha between ('" & CDate(dtpFecha2.Text) & "' and '" & CDate(DTPfecha3.Text) & "') "
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = '00' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '')"
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.hora_registro  = 'OK' FROM to_cronograma_diario_inst INNER JOIN to_cronograma_diario_final_INST ON to_cronograma_diario_inst.fmes_plan = to_cronograma_diario_final_INST.fmes_plan and to_cronograma_diario_inst.bien_codigo  = to_cronograma_diario_final_INST.bien_codigo WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') "
'
'        db.Execute "update to_cronograma_diario_inst set to_cronograma_diario_inst.estado_activo  = 'REG', to_cronograma_diario_inst.estado_codigo  = 'REG' WHERE (to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & ") AND (to_cronograma_diario_inst.bien_codigo <> '') AND (to_cronograma_diario_inst.hora_registro  = '00')"
'
'        Call ABRIR_TABLA_DET
'    End If
'    FraDet5.Visible = False
'
End Sub

Private Sub BtnGrabar6_Click()
'    Set rs_aux6 = New ADODB.Recordset
'    If rs_aux6.State = 1 Then rs_aux6.Close
'    rs_aux6.Open "Select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
'    If rs_aux6.RecordCount > 0 Then
'        'MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), debe deshabilitarlos (Retornar) o utilizar el botón (Envia Uno) ...", vbExclamation, "Validación de Registro"
'        MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), solo podrá procesar la opción 3. ...", vbExclamation, "Validación de Registro"
'        If Option8.Value = True Then
'            Call COPIA_ALGUNOS
'            db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'        End If
'    Else
'      If Option6.Value = True Then
'        Call COPIA_TODOS
'        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'      End If
'      If Option7.Value = True Then
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_orden  = to_cronograma_diario_inst.bien_orden, to_cronograma_diario_final_INST.bien_codigo = to_cronograma_diario_inst.bien_codigo, to_cronograma_diario_final_INST.unidad_codigo_tec = to_cronograma_diario_inst.unidad_codigo_tec, " & _
'        " to_cronograma_diario_final_INST.tec_plan_codigo = to_cronograma_diario_inst.tec_plan_codigo, to_cronograma_diario_final_INST.edif_descripcion = to_cronograma_diario_inst.edif_descripcion, to_cronograma_diario_final_INST.estado_activo = 'APR' FROM to_cronograma_diario_final_INST INNER JOIN to_cronograma_diario_inst " & _
'        " ON to_cronograma_diario_final_INST.fmes_plan  = to_cronograma_diario_inst.fmes_plan AND to_cronograma_diario_final_INST.dia_correl  = to_cronograma_diario_inst.dia_correl AND to_cronograma_diario_final_INST.horario_codigo = to_cronograma_diario_inst.horario_codigo WHERE to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'
'        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
'      End If
'
''      sino = MsgBox("Está Seguro de ENVIAR TODO el Cronograma ORIGEN al DESTINO ?." & vbCrLf & " SI-->(Envía solo a los Horarios Laborales definidos en el Destino) " & vbCrLf & " NO-->(Envía todo a todos los días calendario, incluyendo días NO laborales) " & vbCrLf & " Cancelar, la Operación", vbYesNoCancel + vbQuestion, "Atención")
''      If sino = vbYes Then
''      Else
''        If sino = vbNo Then
''        End If
''      End If
''        'Call BtnAñadir2_Click
'      Call ABRIR_TABLA_DET
'    End If
'    FraDet6.Visible = False
End Sub

Private Sub BtnGrabar7_Click()
    If DTPfecha5.Value > DTPfecha4.Value Then
        sino = MsgBox("Elija SI: para ACEPTAR el Cambio de Fechas y se eliminirá el Cronograma. Luego debe generar uno NUEVO ..." & vbCr & _
             "Elija NO: para CANCELAR sin realizar cambios ...", vbYesNo + vbQuestion, "Atención")
        If sino = vbYes Then
            If Ado_detalle2.Recordset.RecordCount > 0 Then
                db.Execute "DELETE to_cronograma_diario_final_INST where fmes_plan = " & Ado_detalle1.Recordset!correlativo & " "
            End If
            Ado_detalle1.Recordset!fecha_ini_max = DTPfecha4.Value
            Ado_detalle1.Recordset!fecha_fin_max = DTPfecha5.Value
            Ado_detalle1.Recordset!estado_activo = "REG"
            Ado_detalle1.Recordset!estado_codigo = "REG"
            Ado_detalle1.Recordset.Update
        End If
    Else
        MsgBox "La Fecha Inicio NO puede ser mayor a la Fecha Fin, corrija y vuelva a intentar ...", vbInformation, "Información"
    End If
    FraDet7.Visible = False
    
'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)

'    VAR_FECH1 = CDate(DTPfecha4.Text)
'    VAR_FECH2 = CDate(DTPfecha5.Text)
'
'    If VAR_SW2 = "HAB" Then
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET estado_activo  = 'REG', observaciones = 'HORARIO LABORABLE', edif_descripcion = '', tec_plan_codigo = '0' WHERE fmes_plan = " & VAR_FMES & " AND (estado_activo = 'ANL' OR estado_activo = 'APC') and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
'    Else
'        db.Execute "UPDATE to_cronograma_diario_final_INST SET estado_activo  = 'ANL', observaciones = 'HORARIO NO LABORABLE', edif_descripcion = '', tec_plan_codigo = '0' WHERE fmes_plan = " & VAR_FMES & " AND estado_activo <> 'APR' and dia_fecha between '" & VAR_FECH1 & "' and '" & VAR_FECH2 & "' "
'    End If
'    Call ABRIR_TABLA_DET
'    FraDet7.Visible = False
'    VAR_SW2 = ""
'    VAR_MSG = ""
'    fraOpciones.Visible = True
'    FrmABMDet.Visible = True
'    FraGrabarCancelar.Visible = True
'    fraOpciones2.Visible = True
End Sub

Private Sub BtnGrabar8_Click()
'    VAR_AUX2 = VAR_FMES     ' fmes_plan
'    'Carga "Codigos de Insumos" al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1 = tv_cronograma_y_detalle.bien_codigo1 , to_cronograma_diario_final_INST.bien_codigo2 = tv_cronograma_y_detalle.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3 = tv_cronograma_y_detalle.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4 = tv_cronograma_y_detalle.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_y_detalle.bien_codigo5 " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final_INST.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final_INST.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "
'
'    'Actualiza Cantidad de Insumos al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_y_detalle.cantidad1 , to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_y_detalle.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_y_detalle.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_y_detalle.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_y_detalle.cantidad5 " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_y_detalle ON (to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_y_detalle.bien_codigo AND to_cronograma_diario_final_INST.unidad_codigo_tec = tv_cronograma_y_detalle.unidad_codigo_tec) where (to_cronograma_diario_final_INST.fmes_plan >= " & VAR_AUX2 & ") and (tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND tv_cronograma_y_detalle.ges_gestion = '" & Ado_datos.Recordset!ges_gestion & "' ) "
'
''    'Quita Cantidad de Insumo3 e Insumo4 en meses pares al Crono Final
''    sino = MsgBox("Elija SI: para programar en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC) los insumos 3 y 4..." & vbCr & _
''             "Elija NO: para programar en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV) los insumos 3 y 4....", vbYesNo + vbQuestion, "Atención")
''    If sino = vbYes Then
'    If Option10.Value = True Then
'        'Programar Meses IMPARES y quitar PARES
'        db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0', to_cronograma_diario_final_INST.cantidad4 = '0' From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_mensual_par ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_mensual_par.fmes_plan ) " & _
'        " where (to_cronograma_diario_final_INST.fmes_plan >= " & VAR_AUX2 & " ) "
'    Else
'        'PROGRAMAR en Meses PARES y quitar Mes IMPARES
'        db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0', to_cronograma_diario_final_INST.cantidad4 = '0' From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_mensual_impar ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_mensual_impar.fmes_plan ) " & _
'        " where (to_cronograma_diario_final_INST.fmes_plan >= " & VAR_AUX2 & "  ) "
'    End If
'
'    'Actualiza Cantidad de Insumos al Crono Final Bmes, Tmes, etc.
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo) WHERE to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " and tv_cronograma_insumos.unimed_codigo <> 'MES' "
'
'    'Actualiza Carta al Crono Final
'    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_carta.carta " & _
'    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_carta ON (to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_carta.bien_codigo) WHERE to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final_INST.bien_codigo <> '' "
'
'    db.Execute " update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.carta = tv_cronograma_y_detalle.carta from to_cronograma_diario_final_INST inner join tv_cronograma_y_detalle on to_cronograma_diario_final_INST.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_final_INST.fmes_plan = " & VAR_AUX2 & " and to_cronograma_diario_final_INST.bien_codigo <> '' "
'
'    MsgBox "Se actualizaron los Insumos desde CRONOGRAMA POR CONTRATO correspondientes a la misma Gestión y Zona del CRONOGRAMA FINAL (DESTINO) ...", vbInformation, "Información"
End Sub

Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
''    'Actualiza Codigos de Insumos al Crono Final
''    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final_INST.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
''    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
''    'Actualiza Cantidad de Insumos al Crono Final
''    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.cantidad1 = tv_cronograma_insumos.cantidad1, to_cronograma_diario_final_INST.cantidad2 = tv_cronograma_insumos.cantidad2, to_cronograma_diario_final_INST.cantidad3 = tv_cronograma_insumos.cantidad3, to_cronograma_diario_final_INST.cantidad4 = tv_cronograma_insumos.cantidad4, to_cronograma_diario_final_INST.cantidad5 = tv_cronograma_insumos.cantidad5 " & _
''    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
''    'Quita Cantidad de Insumo3 en meses pares al Crono Final
''    db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0' From to_cronograma_diario_final_INST INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final_INST.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
''    " where to_cronograma_mensual.fmes_correl = '2' or to_cronograma_mensual.fmes_correl = '4' or to_cronograma_mensual.fmes_correl = '6' or to_cronograma_mensual.fmes_correl = '8' or to_cronograma_mensual.fmes_correl = '10' or to_cronograma_mensual.fmes_correl = '12' "
''    'Actualiza Carta al Crono Final
''    db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_insumos.carta " & _
''    " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"
'
'    'to_cronograma_diario_final_INST
'    Set rs_datos1 = New ADODB.Recordset
'    If rs_datos1.State = 1 Then rs_datos1.Close
'    rs_datos1.Open "select distinct bien_codigo  from to_cronograma_diario_final_INST where fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and bien_codigo <>'' ", db, adOpenStatic
'    If rs_datos1.RecordCount > 0 Then
'        VAR_REG = rs_datos1.RecordCount
'        VAR_CANT1 = rs_datos1.RecordCount
'        'Actualiza Carta al Crono Final
''        db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.carta  = tv_cronograma_carta.carta " & _
''        " From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_carta ON (to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_carta.bien_codigo)  " & _
''        " WHERE to_cronograma_diario_final_INST.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " and to_cronograma_diario_final_INST.bien_codigo <> '' "
'    Else
'        VAR_REG = "0"
'        VAR_CANT1 = "0"
'    End If
    
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_Instalacion.rpt"
    'CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_R302_Instalacion_PRUEBA.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
'    Select Case Me.Ado_datos.Recordset!unidad_codigo_tec
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
        VAR_REG = "0"
        
      'Cmb_Mes.Text = "ENERO"
      CR01.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
      CR01.Formulas(2) = "periodo = '" & lbl_texto2 & "' "
      CR01.Formulas(3) = "TotalReg = " & VAR_REG & " "
      CR01.Formulas(4) = "CANT1 = " & VAR_CANT1 & " "
      
     CR01.StoredProcParam(0) = Ado_detalle1.Recordset!correlativo                  'Me.Ado_datos.Recordset!fmes_plan
     CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!zpiloto_codigo
     
'    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!ges_gestion
'    CR01.StoredProcParam(1) = Me.Ado_datos.Recordset!unidad_codigo_tec
'    CR01.StoredProcParam(2) = Me.Ado_datos.Recordset!zpiloto_codigo
'    CR01.StoredProcParam(3) = Me.Ado_datos.Recordset!fmes_correl
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir2_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    'db.Execute "Update to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_codigo1  = tv_cronograma_insumos.bien_codigo1, to_cronograma_diario_final_INST.bien_codigo2   = tv_cronograma_insumos.bien_codigo2, to_cronograma_diario_final_INST.bien_codigo3   = tv_cronograma_insumos.bien_codigo3, to_cronograma_diario_final_INST.bien_codigo4   = tv_cronograma_insumos.bien_codigo4, to_cronograma_diario_final_INST.bien_codigo5 = tv_cronograma_insumos.bien_codigo5 " & _
    '" From to_cronograma_diario_final_INST INNER JOIN tv_cronograma_insumos ON (to_cronograma_diario_final_INST.fmes_plan = tv_cronograma_insumos.fmes_plan and to_cronograma_diario_final_INST.bien_codigo  = tv_cronograma_insumos.bien_codigo)"

    'db.Execute "Update to_cronograma_diario_final_INST set to_cronograma_diario_final_INST.cantidad3 = '0' From to_cronograma_diario_final_INST INNER JOIN to_cronograma_mensual ON (to_cronograma_diario_final_INST.fmes_plan = to_cronograma_mensual.fmes_plan) " & _
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

Private Sub BtnModDetalle_Click()
'    If Ado_detalle2.Recordset("estado_activo") = "ANL" Then  '<> "REG"
'      sino = MsgBox("Está Seguro de cambiar a HORARIO LABORABLE ? (Este volverá a ser considerado en el Cronograma) ", vbYesNo + vbQuestion, "Atención")
'      If sino = vbYes Then
'        Ado_detalle2.Recordset!estado_activo = "REG"
'        Ado_detalle2.Recordset!observaciones = "HORARIO LABORABLE"
'        Ado_detalle2.Recordset.Update
'        'Call ABRIR_TABLA_DET
'      End If
'   Else
'        MsgBox "No se puede Habilitar, el registro ya fue Procesado (Estado=APR) o ya está Habilitado (Estado=REG) ...", vbExclamation, "Validación de Registro"
'   End If
'    Call BtnAñadir2_Click
  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    FraDet4.Visible = True
'    txt_obs.Visible = False
'    Frame2.Visible = False
'    BtnGraba4.Visible = False
'    BtnCancelar4.Visible = False
    fraOpciones.Enabled = False
    FrmABMDet.Enabled = False
    fraOpciones2.Enabled = False
'    Select Case Ado_detalle2.Recordset!estado_activo
'        Case "APP"
'            cmd_campo1.Text = "HORARIO POR CONFIRMAR"
'        Case "APC"
'            cmd_campo1.Text = "COMPENSACION"
'        Case "APR"
'            cmd_campo1.Text = "HORARIO LABORAL Confirmado"
'        Case Else
'            cmd_campo1.Text = ""
'    End Select
    'HORARIO LABORAL Confirmado"
'    If dtc_codigo6.Text <> "4211" Then
'        dtc_codigo6.Text = "4211"                   'TRAPO
'        dtc_desc6.BoundText = dtc_codigo6.BoundText
'    End If
'    If dtc_codigo6A.Text <> "479" Then
'        dtc_codigo6A.Text = "479"                   'GASOLINA
'        dtc_desc6A.BoundText = dtc_codigo6A.BoundText
'    End If
'    If dtc_codigo6B.Text <> "500" Then          '3410003 (ANTES)
'        dtc_codigo6B.Text = "500"                   'ACEITE PREPARADO
'        dtc_desc6B.BoundText = dtc_codigo6B.BoundText
'    End If
'    If dtc_codigo6C.Text <> "4529" Then
'        dtc_codigo6C.Text = "4529"                  'ACEITE DELGADO 20/50
'        dtc_desc6C.BoundText = dtc_codigo6C.BoundText
'    End If
'    If dtc_codigo6D.Text <> "3113" Then
'        dtc_codigo6D.Text = "3113"                  'GRASA PARA RODAMIENTO
'        dtc_desc6D.BoundText = dtc_codigo6D.BoundText
'    End If
  
  Else
      MsgBox "No se puede MODIFICAR un cronograma APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnModDetalle2_Click()
'  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    'to_cronograma_diario_final_INST
'    FraDet6.Visible = True
''    Set rs_aux6 = New ADODB.Recordset
''    If rs_aux6.State = 1 Then rs_aux6.Close
''    rs_aux6.Open "Select * from to_cronograma_diario_final_INST where fmes_plan = " & VAR_FMES & " AND bien_codigo <> '' ", db, adOpenStatic
''    If rs_aux6.RecordCount > 0 Then
''        MsgBox "Ya existen registros en el CRONOGRAMA FINAL (DESTINO), debe deshabilitarlos (Retornar) o utilizar el botón (Envia Uno) ...", vbExclamation, "Validación de Registro"
''    Else
''      sino = MsgBox("Está Seguro de ENVIAR TODO el Cronograma ORIGEN al DESTINO ?." & vbCrLf & " SI-->(Envía solo a los Horarios Laborales definidos en el Destino) " & vbCrLf & " NO-->(Envía todo a todos los días calendario, incluyendo días NO laborales) " & vbCrLf & " Cancelar, la Operación", vbYesNoCancel + vbQuestion, "Atención")
''      If sino = vbYes Then
''        Call COPIA_TODOS
''        db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
''      Else
''        If sino = vbNo Then
''            db.Execute "UPDATE to_cronograma_diario_final_INST SET to_cronograma_diario_final_INST.bien_orden  = to_cronograma_diario_inst.bien_orden, to_cronograma_diario_final_INST.bien_codigo = to_cronograma_diario_inst.bien_codigo, to_cronograma_diario_final_INST.unidad_codigo_tec = to_cronograma_diario_inst.unidad_codigo_tec, " & _
''            " to_cronograma_diario_final_INST.tec_plan_codigo = to_cronograma_diario_inst.tec_plan_codigo, to_cronograma_diario_final_INST.edif_descripcion = to_cronograma_diario_inst.edif_descripcion, to_cronograma_diario_final_INST.estado_activo = 'APR' FROM to_cronograma_diario_final_INST INNER JOIN to_cronograma_diario_inst " & _
''            " ON to_cronograma_diario_final_INST.fmes_plan  = to_cronograma_diario_inst.fmes_plan AND to_cronograma_diario_final_INST.dia_correl  = to_cronograma_diario_inst.dia_correl AND to_cronograma_diario_final_INST.horario_codigo = to_cronograma_diario_inst.horario_codigo WHERE to_cronograma_diario_inst.fmes_plan = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
''
''            db.Execute "UPDATE to_cronograma_diario_inst set to_cronograma_diario_inst.estado_codigo   = 'APR' where to_cronograma_diario_inst.fmes_plan  = " & VAR_FMES & " AND to_cronograma_diario_inst.estado_activo = 'APR' "
''        End If
''      End If
''        'Call BtnAñadir2_Click
''      Call ABRIR_TABLA_DET
''    End If
'  Else
'      MsgBox "No se puede ENVIAR, el cronograma ya fue APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
'  End If
End Sub

Private Sub BtnModificar_Click()
    DTPfecha4.Value = Ado_detalle1.Recordset!fecha_ini_max
    DTPfecha5.Value = Ado_detalle1.Recordset!fecha_fin_max
    FraDet7.Visible = True
    
'  On Error GoTo EditErr
''  lblStatus.Caption = "Modificar registro"
'    If Ado_datos.Recordset!estado_codigo = "REG" Then
'        Fra_datos.Visible = True
'        Fra_datos.Enabled = True
'        fraOpciones.Visible = False
'        FraGrabarCancelar.Visible = True
'        dg_datos.Enabled = False
'        VAR_SW = "MOD"
'        'tc_zonas_piloto
'        Set rs_aux4 = New ADODB.Recordset
'        If rs_aux4.State = 1 Then rs_aux4.Close
'        rs_aux4.Open "Select * from tc_zonas_piloto where zpiloto_codigo = " & dtc_codigo3.Text & " ", db, adOpenStatic
'        If rs_aux4.RecordCount > 0 Then
'            dtc_codigo4.Text = rs_aux4!beneficiario_codigo
'            dtc_desc4.BoundText = dtc_codigo4.BoundText
'        End If
'    '    BtnVer.Visible = True
'    Else
'      MsgBox "No se puede MODIFICAR un cronograma APROBADO o ANULADO ...", vbExclamation, "Validación de Registro"
'    End If
'  Exit Sub

'EditErr:
'  MsgBox Err.Description
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub BtnVer2_Click()
'    Set rs_aux14 = New ADODB.Recordset
'    If rs_aux14.State = 1 Then rs_aux14.Close
'    rs_aux14.Open "select * from to_cronograma_diario_final_INST where fmes_plan = '" & VAR_FMES & "'  and estado_activo = 'APR' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    'rs_det1.Sort = "bien_orden"
'    If rs_aux14.RecordCount > 0 Then
'        MsgBox "No se puede Actualizar el #Horas ni Orden, porque ya existen registros en el Cronograma Final de esta Zona en el Mes a procesar, Vuelva a Intentar ...", vbExclamation, "Validación"
'    Else
'        db.Execute " update to_cronograma_diario_inst set to_cronograma_diario_inst.nro_total_horas = tv_cronograma_y_detalle.bien_cantidad_por_empaque from to_cronograma_diario_inst inner join tv_cronograma_y_detalle on to_cronograma_diario_inst.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_inst.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " AND tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
'        db.Execute " update to_cronograma_diario_inst set to_cronograma_diario_inst.bien_orden = tv_cronograma_y_detalle.zona_edif_orden from to_cronograma_diario_inst inner join tv_cronograma_y_detalle on to_cronograma_diario_inst.bien_codigo = tv_cronograma_y_detalle.bien_codigo where to_cronograma_diario_inst.fmes_plan = " & Ado_datos.Recordset!fmes_plan & " AND tv_cronograma_y_detalle.zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
'
'        Call ABRIR_TABLA_DET
'        MsgBox "Se Actualizó el <#Horas> por equipo y el <Orden> actual de la Organización de Zonas ...", vbInformation, "Información"
'    End If
End Sub

Private Sub dg_datos_ButtonClick(ByVal ColIndex As Integer)
    busca3 = 2
End Sub

Private Sub dg_datos_Click()
    buscados = 0
    busca3 = 2
End Sub

Private Sub dg_datos_GotFocus()
    busca3 = 2
End Sub

Private Sub dg_datos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    busca3 = 2
End Sub

Private Sub dg_det1_DblClick()
    cmd_campo2.Text = "2"
    Call BtnGrabar2_Click
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub


Private Sub dtc_desc5_Click(Area As Integer)
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub DTPfecha3_LostFocus()
    'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)
    'If Val(Txt_descripcion) < 0 Then
    If Val(DateDiff("y", CDate(DTPfecha2), CDate(DTPfecha3))) < 0 Then
        MsgBox "La Fecha Inicial NO puede ser MAYOR a la Fecha Final, Vuelva a Intentar ...", vbExclamation, "Validación de Registro"
        DTPfecha3.SetFocus
    End If
End Sub

Private Sub DTPfecha5_LostFocus()
    'Txt_descripcion = DateDiff("y", DTPfechaIni, DTPfechaFin)
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    VAR_SW2 = ""
    busca3 = 0
    VAR_SW0 = 0
    cmd_campo2.Text = "2"
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
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTOC = "2"
    End If
    VAR_UORIGEN = "DNINS"
    parametro = Aux
    VAR_ANL = ""
    'ACTUALIZA DATOS DEL CONTRATO
    db.Execute " update tc_zona_piloto_edif_inst SET tc_zona_piloto_edif_inst.venta_codigo  = AV_VENTAS_NUEVAS_APR.venta_codigo , tc_zona_piloto_edif_inst.unimed_codigo = 'MES', tc_zona_piloto_edif_inst.codigo_empresa =codigo_empresa, tc_zona_piloto_edif_inst.solicitud_tipo ='3', tc_zona_piloto_edif_inst.Gratuito ='SI' FROM tc_zona_piloto_edif_inst INNER JOIN AV_VENTAS_NUEVAS_APR ON tc_zona_piloto_edif_inst.edif_codigo = AV_VENTAS_NUEVAS_APR.edif_codigo where tc_zona_piloto_edif_inst.venta_codigo Is Null "
    db.Execute " update tc_zona_piloto_edif_inst SET tc_zona_piloto_edif_inst.unidad_codigo_ant  = AV_VENTAS_NUEVAS_APR.unidad_codigo_ant FROM tc_zona_piloto_edif_inst INNER JOIN AV_VENTAS_NUEVAS_APR ON tc_zona_piloto_edif_inst.venta_codigo = AV_VENTAS_NUEVAS_APR.venta_codigo where tc_zona_piloto_edif_inst.unidad_codigo_ant Is Null "
    'Actualiza Responsables de Zona
    db.Execute " UPDATE tc_zona_piloto_edif_inst SET tc_zona_piloto_edif_inst.fecha_ini_max  = av_ventas_alcance_INST.fecha_inicio_alcance, tc_zona_piloto_edif_inst.fecha_fin_max = av_ventas_alcance_INST.fecha_fin_alcance FROM tc_zona_piloto_edif_inst INNER JOIN av_ventas_alcance_INST ON tc_zona_piloto_edif_inst.edif_codigo = av_ventas_alcance_INST .edif_codigo where tc_zona_piloto_edif_inst.fecha_ini_max Is Null "
    
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    
'    Fra_datos.Enabled = False
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
'    dtc_desc1.BoundText = dtc_codigo1.BoundText
        
    'tc_zonas_piloto
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from tc_zonas_piloto order by zpiloto_descripcion ", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
'    dtc_desc3.BoundText = dtc_codigo3.BoundText
    
    'Beneficiario Funcionario CGI (Vendedor, Cobrador, Adm, etc.)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & VAR_UORIGEN & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
'    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'INSUMOS
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "select distinct * from av_bienes_vs_venta_detalle where par_codigo = '33100' or par_codigo = '34110' ORDER BY bien_descripcion ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
'    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub



Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros 2021)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    queryinicial = "select * From tc_zonas_piloto_inst  "      'WHERE (ges_gestion = '2022' )
    'queryinicial = "select * From to_cronograma_mensual_inst  "      'WHERE (ges_gestion = '2022' )
'    Select Case VAR_DPTOC
'        Case "1"    ' Chuquisaca
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='34' or zpiloto_codigo='35' or zpiloto_codigo='36' or zpiloto_codigo='38') AND ges_gestion = '2022' ) "
'        Case "2"    'La Paz - Tecnico
'            If glusuario = "OCOLODRO" Then
'                queryinicial = "select * From to_cronograma_mensual WHERE (ges_gestion = '2022' ) "
'            Else
'                queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo<'16' OR zpiloto_codigo='28' OR zpiloto_codigo='29' OR zpiloto_codigo='30' OR zpiloto_codigo='37' OR zpiloto_codigo='39' OR zpiloto_codigo='40')  AND ges_gestion = '2022' ) "
'            End If
'        Case "3"    'Cochabamba
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='17' or zpiloto_codigo='18' or zpiloto_codigo='19' or zpiloto_codigo='20') AND ges_gestion = '2022' ) "
'        Case "7"    'Santa Cruz
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='21' or zpiloto_codigo='22' or zpiloto_codigo='23' or zpiloto_codigo='24' or zpiloto_codigo='25' or zpiloto_codigo='26' or zpiloto_codigo='27' or zpiloto_codigo='31' or zpiloto_codigo='32' or zpiloto_codigo='33' or zpiloto_codigo = '34') AND ges_gestion = '2022' ) "
'        Case "4"    'Oruro - Tecnico
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='16' ) AND ges_gestion = '2022' ) "
'        Case "5"    ' Potosi
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='35' ) AND ges_gestion = '2022' ) "
'        Case "6"    ' Tarija
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='36' ) AND ges_gestion = '2022' ) "
'        Case "8"    ' Beni
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='32' ) AND ges_gestion = '2022' ) "
'        Case "9"    ' Pando
'            queryinicial = "select * From to_cronograma_mensual WHERE ((zpiloto_codigo='33' ) AND ges_gestion = '2022' ) "
'        Case Else    ' TODO
'            queryinicial = "select * From to_cronograma_mensual where ( ges_gestion = '2022' ) "
'     End Select

    rs_datos.Sort = "zpiloto_codigo"           'ges_gestion, fmes_correl,
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

Private Sub ABRIR_TABLA_DET()
'On Error GoTo UpdateErr
'    Set rs_det1 = New ADODB.Recordset
'    If rs_det1.State = 1 Then rs_det1.Close
'    rs_det1.Open "select * from tv_zona_piloto_edif_inst where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' order by zona_edif_orden ", db, adOpenKeyset, adLockOptimistic, adCmdText
'    Set Ado_detalle1.Recordset = rs_det1
'    Set dg_det1.DataSource = Ado_detalle1.Recordset
'    If Ado_detalle1.Recordset.RecordCount > 0 Then
'        dg_det1.Visible = True
'        VAR_SW0 = 1
''        If swnuevo = 0 Then
''            'gc_edificaciones
''            Set rs_datos5 = New ADODB.Recordset
''            If rs_datos5.State = 1 Then rs_datos5.Close
''            rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
''            Set Ado_datos5.Recordset = rs_datos5
''            dtc_desc5.BoundText = dtc_codigo5.BoundText
''        End If
'    Else
'        dg_det1.Visible = False
'        VAR_SW0 = 2
'    End If
'
''    If Option3.Value = True Then
''        Set rs_det1 = New ADODB.Recordset
''        If rs_det1.State = 1 Then rs_det1.Close
''        rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & VAR_FMES & "'  and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        rs_det1.Sort = "bien_orden"
''        Set Ado_detalle1.Recordset = rs_det1
''        If Ado_detalle1.Recordset.RecordCount > 0 Then
''            Set dg_det1.DataSource = Ado_detalle1.Recordset
''        Else
''            Set dg_det1.DataSource = rsNada
''        End If
''    End If
''    If Option4.Value = True Then
''        Set rs_det1 = New ADODB.Recordset
''        If rs_det1.State = 1 Then rs_det1.Close
''        'rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "' and estado_activo <> 'ANL' AND estado_activo <> 'APR'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        rs_det1.Open "select * from to_cronograma_diario_inst where fmes_plan = '" & VAR_FMES & "' and estado_codigo =  'REG' AND bien_codigo <> ''  ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        rs_det1.Sort = "bien_orden"
''        Set Ado_detalle1.Recordset = rs_det1
''        If Ado_detalle1.Recordset.RecordCount > 0 Then
''            Set dg_det1.DataSource = Ado_detalle1.Recordset
''        Else
''            Set dg_det1.DataSource = rsNada
''        End If
''    End If
''    If Option1.Value = True Then
''        Set rs_det2 = New ADODB.Recordset
''        If rs_det2.State = 1 Then rs_det2.Close
''        rs_det2.Open "select * from to_cronograma_diario_final_INST where fmes_plan = '" & VAR_FMES & "' and estado_activo <> 'ANL' AND bien_codigo <> '' ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        'rs_det2.Sort = "bien_orden"
''        Set Ado_detalle2.Recordset = rs_det2
''        If Ado_detalle2.Recordset.RecordCount > 0 Then
''            Ado_detalle2.Recordset.MoveLast
''            Set dg_det2.DataSource = Ado_detalle2.Recordset
''            dg_det2.Visible = True
''        Else
''            Set dg_det2.DataSource = rsNada
''            dg_det2.Visible = False
''        End If
''    End If
''    If Option2.Value = True Then
''        Set rs_det2 = New ADODB.Recordset
''        If rs_det2.State = 1 Then rs_det2.Close
''        rs_det2.Open "select * from to_cronograma_diario_final_INST where fmes_plan = '" & VAR_FMES & "'  ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        'rs_det2.Sort = "bien_orden"
''        Set Ado_detalle2.Recordset = rs_det2
''        If Ado_detalle2.Recordset.RecordCount > 0 Then
''            Ado_detalle2.Recordset.MoveLast
''            Set dg_det2.DataSource = Ado_detalle2.Recordset
''            dg_det2.Visible = True
''        Else
''            Set dg_det2.DataSource = rsNada
''            dg_det2.Visible = False
''        End If
''    End If
''    If Option5.Value = True Then
''        Set rs_det2 = New ADODB.Recordset
''        If rs_det2.State = 1 Then rs_det2.Close
''        rs_det2.Open "select * from to_cronograma_diario_final_INST where fmes_plan = '" & VAR_FMES & "' and estado_activo <> 'ANL' AND estado_activo <> 'APR' ", db, adOpenKeyset, adLockOptimistic, adCmdText
''        'rs_det2.Sort = "bien_orden"
''        Set Ado_detalle2.Recordset = rs_det2
''        If Ado_detalle2.Recordset.RecordCount > 0 Then
''            Ado_detalle2.Recordset.MoveLast
''            Set dg_det2.DataSource = Ado_detalle2.Recordset
''            dg_det2.Visible = True
''        Else
''            Set dg_det2.DataSource = rsNada
''            dg_det2.Visible = False
''        End If
''    End If
'   Exit Sub
'UpdateErr:
'  MsgBox Err.Description
End Sub

Private Sub Option1_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option2_Click()
    Call ABRIR_TABLA_DET
End Sub

Private Sub Option3_Click()
    'Call ABRIR_TABLA_DET
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from tv_zona_piloto_edif_inst where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' and estado_codigo = 'REG' order by fecha_ini_max ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        dg_det1.Visible = True
        VAR_SW0 = 1
        'If swnuevo = 0 Then
        '    'gc_edificaciones
        '    Set rs_datos5 = New ADODB.Recordset
        '    If rs_datos5.State = 1 Then rs_datos5.Close
        '    rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
        '    Set Ado_datos5.Recordset = rs_datos5
        '    dtc_desc5.BoundText = dtc_codigo5.BoundText
        'End If
    Else
        dg_det1.Visible = False
        VAR_SW0 = 2
    End If
End Sub

Private Sub Option4_Click()
    'Call ABRIR_TABLA_DET
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    rs_det1.Open "select * from tv_zona_piloto_edif_inst where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' and estado_codigo = 'REG' order by fecha_ini_max ", db, adOpenKeyset, adLockOptimistic, adCmdText
    Set Ado_detalle1.Recordset = rs_det1
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        dg_det1.Visible = True
        VAR_SW0 = 1
    Else
        dg_det1.Visible = False
        VAR_SW0 = 2
    End If
End Sub

Private Sub Option5_Click()
    Call ABRIR_TABLA_DET
End Sub


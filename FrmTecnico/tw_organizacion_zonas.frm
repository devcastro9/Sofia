VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tw_organizacion_zonas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tecnico - Organizacion de Zonas"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "tw_organizacion_zonas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11280
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraNewZona 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Elija nueva Zona Piloto a la que se enviará el registro elegido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2520
      Left            =   6360
      TabIndex        =   79
      Top             =   1080
      Visible         =   0   'False
      Width           =   12540
      Begin VB.ComboBox CmbGestion 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ges_gestion"
         DataSource      =   "Ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_organizacion_zonas.frx":0A02
         Left            =   3600
         List            =   "tw_organizacion_zonas.frx":0A21
         TabIndex        =   87
         Text            =   "2023"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   10560
         TabIndex        =   83
         Top             =   690
         Width           =   270
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   180
         ScaleHeight     =   660
         ScaleWidth      =   12180
         TabIndex        =   80
         Top             =   1680
         Width           =   12180
         Begin VB.PictureBox BtnGraba3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4560
            Picture         =   "tw_organizacion_zonas.frx":0A5B
            ScaleHeight     =   585
            ScaleWidth      =   1245
            TabIndex        =   82
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelar3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6360
            Picture         =   "tw_organizacion_zonas.frx":1249
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   81
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSDataListLib.DataCombo dtc_desc10 
         Bindings        =   "tw_organizacion_zonas.frx":1B35
         Height          =   315
         Left            =   3600
         TabIndex        =   84
         Top             =   675
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "zpiloto_descripcion"
         BoundColumn     =   "zpiloto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo10 
         Bindings        =   "tw_organizacion_zonas.frx":1B4E
         Height          =   315
         Left            =   9840
         TabIndex        =   85
         Top             =   675
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Gestion"
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
         Left            =   2640
         TabIndex        =   88
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Piloto Destino ..."
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
         Left            =   1560
         TabIndex        =   86
         Top             =   645
         Width           =   1935
      End
   End
   Begin VB.Frame FraDet2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Edificio (Detalle)"
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
      Height          =   4200
      Left            =   6120
      TabIndex        =   53
      Top             =   720
      Visible         =   0   'False
      Width           =   13020
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Para programar los insumos 3 y 4, en meses IMPARES (ENE, MAR, MAY, JUL, SEP, NOV)."
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
         Left            =   5280
         TabIndex        =   75
         Top             =   2760
         Width           =   7575
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Para programar los insumos 3 y 4, en meses PARES (FEB, ABR, JUN, AGO, OCT, DIC)."
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
         Left            =   5280
         TabIndex        =   74
         Top             =   3120
         Width           =   7335
      End
      Begin VB.PictureBox fra_opciones2 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   150
         ScaleHeight     =   660
         ScaleWidth      =   12705
         TabIndex        =   69
         Top             =   3480
         Width           =   12705
         Begin VB.PictureBox BtnGrabarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   4800
            Picture         =   "tw_organizacion_zonas.frx":1B67
            ScaleHeight     =   615
            ScaleWidth      =   1275
            TabIndex        =   71
            Top             =   0
            Width           =   1280
         End
         Begin VB.PictureBox BtnCancelarDet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   6600
            Picture         =   "tw_organizacion_zonas.frx":233D
            ScaleHeight     =   615
            ScaleWidth      =   1395
            TabIndex        =   70
            Top             =   0
            Width           =   1400
         End
      End
      Begin VB.TextBox txt_obs 
         BackColor       =   &H00FFFFFF&
         DataField       =   "observaciones"
         DataSource      =   "Ado_detalle1"
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
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2280
         Width           =   10125
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   290
         Left            =   2160
         TabIndex        =   64
         Top             =   380
         Width           =   270
      End
      Begin VB.ComboBox cmd_campo2 
         BackColor       =   &H00FFFFFF&
         DataField       =   "zorden_cambio"
         DataSource      =   "Ado_detalle1"
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "tw_organizacion_zonas.frx":2C29
         Left            =   4080
         List            =   "tw_organizacion_zonas.frx":2CFF
         TabIndex        =   11
         Text            =   "0"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Txt_campo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "zona_edif_orden"
         DataSource      =   "Ado_detalle1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "tw_organizacion_zonas.frx":2E12
         Top             =   2760
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "correlativo"
         DataSource      =   "Ado_detalle1"
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
         Height          =   360
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtc_desc6 
         Bindings        =   "tw_organizacion_zonas.frx":2E14
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo6 
         Bindings        =   "tw_organizacion_zonas.frx":2E2D
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9960
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc5 
         Bindings        =   "tw_organizacion_zonas.frx":2E46
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "edif_descripcion"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo5 
         Bindings        =   "tw_organizacion_zonas.frx":2E5F
         DataField       =   "edif_codigo"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         BackColor       =   12632256
         ListField       =   "edif_codigo"
         BoundColumn     =   "edif_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc7 
         Bindings        =   "tw_organizacion_zonas.frx":2E78
         DataField       =   "beneficiario_codigo_rep"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   1320
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc8 
         Bindings        =   "tw_organizacion_zonas.frx":2E91
         DataField       =   "beneficiario_codigo_cobr"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   1800
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo7 
         Bindings        =   "tw_organizacion_zonas.frx":2EAA
         DataField       =   "beneficiario_codigo_rep"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9960
         TabIndex        =   61
         Top             =   1320
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo8 
         Bindings        =   "tw_organizacion_zonas.frx":2EC3
         DataField       =   "beneficiario_codigo_cobr"
         DataSource      =   "Ado_detalle1"
         Height          =   315
         Left            =   9960
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label dtc_aux5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         DataField       =   "edif_descripcion"
         DataSource      =   "Ado_detalle1"
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
         Height          =   315
         Left            =   2520
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   7365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Left            =   240
         TabIndex        =   67
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label lbl_campo5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Edificio"
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
         Left            =   240
         TabIndex        =   63
         Top             =   405
         Width           =   660
      End
      Begin VB.Label lbl_orden_camb 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar a -->"
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
         Left            =   2760
         TabIndex        =   60
         Top             =   2775
         Width           =   1200
      End
      Begin VB.Label lbl_orden 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Prioridad"
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
         Left            =   240
         TabIndex        =   59
         Top             =   2775
         Width           =   1830
      End
      Begin VB.Label lbl_campo7 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Reparaciones / Emergencias"
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
         Height          =   480
         Left            =   240
         TabIndex        =   58
         Top             =   1215
         Width           =   2100
      End
      Begin VB.Label lbl_campo8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrador"
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
         Left            =   240
         TabIndex        =   57
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lbl_campo6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Tecnico Mantenimiento"
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
         Left            =   240
         TabIndex        =   56
         Top             =   840
         Width           =   2085
      End
   End
   Begin VB.Frame Fra_datos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Registro Cabecera"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3240
      Left            =   6240
      TabIndex        =   27
      Top             =   720
      Width           =   12900
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Programa insumos 3 y 4, en mes PAR (FEB, ABR, JUN, AGO, OCT, DIC)."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   6960
         TabIndex        =   77
         Top             =   2640
         Width           =   5775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Programa insumos 3 y 4, en mes IMPAR (ENE, MAR, MAY, JUL, SEP, NOV)."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   330
         Left            =   240
         TabIndex        =   76
         Top             =   2640
         Width           =   5895
      End
      Begin VB.TextBox DtpFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "fecha_registro"
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
         Height          =   360
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Txt_campo2 
         DataField       =   "zpiloto_descripcion"
         DataSource      =   "Ado_datos"
         Height          =   285
         Left            =   6840
         MultiLine       =   -1  'True
         TabIndex        =   0
         Text            =   "tw_organizacion_zonas.frx":2EDC
         Top             =   480
         Width           =   5685
      End
      Begin VB.TextBox txt_codigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "zpiloto_codigo"
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
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   855
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
         Height          =   360
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dtc_codigo3 
         Bindings        =   "tw_organizacion_zonas.frx":2EDE
         DataField       =   "munic_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5520
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "munic_codigo"
         BoundColumn     =   "munic_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo2 
         Bindings        =   "tw_organizacion_zonas.frx":2EF7
         DataField       =   "prov_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   11880
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prov_codigo"
         BoundColumn     =   "prov_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc2 
         Bindings        =   "tw_organizacion_zonas.frx":2F10
         DataField       =   "prov_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   6840
         TabIndex        =   2
         Top             =   960
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "prov_descripcion"
         BoundColumn     =   "prov_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc3 
         Bindings        =   "tw_organizacion_zonas.frx":2F29
         DataField       =   "munic_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "munic_descripcion"
         BoundColumn     =   "munic_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc4 
         Bindings        =   "tw_organizacion_zonas.frx":2F42
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   2040
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_denominacion"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo4 
         Bindings        =   "tw_organizacion_zonas.frx":2F5B
         DataField       =   "beneficiario_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   8880
         TabIndex        =   40
         Top             =   2040
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "beneficiario_codigo"
         BoundColumn     =   "beneficiario_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc1 
         Bindings        =   "tw_organizacion_zonas.frx":2F74
         DataField       =   "depto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   900
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "depto_descripcion"
         BoundColumn     =   "depto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo1 
         Bindings        =   "tw_organizacion_zonas.frx":2F8D
         DataField       =   "depto_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   5160
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "depto_codigo"
         BoundColumn     =   "depto_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_desc9 
         Bindings        =   "tw_organizacion_zonas.frx":2FA6
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   9120
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_denominacion"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin MSDataListLib.DataCombo dtc_codigo9 
         Bindings        =   "tw_organizacion_zonas.frx":2FBF
         DataField       =   "zona_codigo"
         DataSource      =   "Ado_datos"
         Height          =   315
         Left            =   11760
         TabIndex        =   65
         Top             =   2640
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "zona_codigo"
         BoundColumn     =   "zona_codigo"
         Text            =   "Todos"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Zona Geográfica"
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
         Left            =   9960
         TabIndex        =   66
         Top             =   2520
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl_campo1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
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
         Left            =   240
         TabIndex        =   42
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label lbl_campo4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Técnico Responsable Zona"
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
         Left            =   240
         TabIndex        =   41
         Top             =   2040
         Width           =   2520
      End
      Begin VB.Label lbl_campo3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Municipio"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl_campo2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Left            =   5880
         TabIndex        =   36
         Top             =   960
         Width           =   840
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
         Left            =   5145
         TabIndex        =   35
         Top             =   135
         Visible         =   0   'False
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
         Left            =   2440
         TabIndex        =   34
         Top             =   380
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Zona"
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
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label lbl_zona 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Zona Piloto"
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
         Left            =   4680
         TabIndex        =   32
         Top             =   480
         Width           =   2085
      End
   End
   Begin VB.PictureBox fra_opciones 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   20280
      TabIndex        =   48
      Top             =   0
      Width           =   20280
      Begin VB.PictureBox BtnImprimir1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   8280
         Picture         =   "tw_organizacion_zonas.frx":2FD8
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   73
         ToolTipText     =   "Edificios en Cronograma vs. Contratos de Mantenimiento"
         Top             =   0
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.PictureBox BtnAñadir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         Picture         =   "tw_organizacion_zonas.frx":38A5
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   12
         ToolTipText     =   "Crea una Nueva Zona Piloto"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1305
         Picture         =   "tw_organizacion_zonas.frx":4064
         ScaleHeight     =   615
         ScaleWidth      =   1425
         TabIndex        =   13
         ToolTipText     =   "Modifica datos de la Zona elegida"
         Top             =   0
         Width           =   1430
      End
      Begin VB.PictureBox BtnEliminar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2640
         Picture         =   "tw_organizacion_zonas.frx":4979
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   14
         ToolTipText     =   "Anula Zona elegida"
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnAprobar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3960
         Picture         =   "tw_organizacion_zonas.frx":50C5
         ScaleHeight     =   615
         ScaleWidth      =   1320
         TabIndex        =   15
         ToolTipText     =   "Aprueba el Registro Elegido"
         Top             =   0
         Width           =   1320
      End
      Begin VB.PictureBox BtnBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   5400
         Picture         =   "tw_organizacion_zonas.frx":58F8
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   16
         ToolTipText     =   "Busca Registros "
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox BtnImprimir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6840
         Picture         =   "tw_organizacion_zonas.frx":60AD
         ScaleHeight     =   615
         ScaleWidth      =   1395
         TabIndex        =   17
         ToolTipText     =   "Imprimir Todas las Zonas Piloto"
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
         Picture         =   "tw_organizacion_zonas.frx":697A
         ScaleHeight     =   615
         ScaleWidth      =   1245
         TabIndex        =   49
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   12735
         TabIndex        =   50
         Top             =   180
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
      TabIndex        =   44
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
         Picture         =   "tw_organizacion_zonas.frx":713C
         ScaleHeight     =   615
         ScaleWidth      =   1275
         TabIndex        =   46
         Top             =   0
         Width           =   1280
      End
      Begin VB.PictureBox BtnCancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   6435
         Picture         =   "tw_organizacion_zonas.frx":7912
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   45
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
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   13095
         TabIndex        =   47
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame FraDet1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LISTADO DE EDIFICIOS"
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
      Height          =   5415
      Left            =   0
      TabIndex        =   30
      Top             =   4020
      Width           =   19125
      Begin VB.PictureBox fra_opciones_det 
         BackColor       =   &H80000015&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   150
         ScaleHeight     =   660
         ScaleWidth      =   18825
         TabIndex        =   51
         Top             =   240
         Width           =   18825
         Begin VB.PictureBox BtnAddDetalle3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000015&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   10080
            Picture         =   "tw_organizacion_zonas.frx":81FE
            ScaleHeight     =   615
            ScaleWidth      =   1095
            TabIndex        =   78
            Top             =   0
            Width           =   1095
         End
         Begin VB.PictureBox BtnModificar2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   5760
            Picture         =   "tw_organizacion_zonas.frx":8B53
            ScaleHeight     =   615
            ScaleWidth      =   1545
            TabIndex        =   72
            Top             =   0
            Width           =   1545
         End
         Begin VB.PictureBox BtnAddDetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            Picture         =   "tw_organizacion_zonas.frx":9AFC
            ScaleHeight     =   615
            ScaleWidth      =   1200
            TabIndex        =   20
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
            Picture         =   "tw_organizacion_zonas.frx":A2BB
            ScaleHeight     =   615
            ScaleWidth      =   1425
            TabIndex        =   21
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
            Picture         =   "tw_organizacion_zonas.frx":ABD0
            ScaleHeight     =   615
            ScaleWidth      =   1215
            TabIndex        =   22
            Top             =   0
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid dg_det1 
         Bindings        =   "tw_organizacion_zonas.frx":B31C
         Height          =   4335
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   18855
         _ExtentX        =   33258
         _ExtentY        =   7646
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
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "zona_edif_orden"
            Caption         =   "Orden"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16394
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "zpiloto_codigo"
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
         BeginProperty Column02 
            DataField       =   "edif_codigo_corto"
            Caption         =   "Cod_Edificio"
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
            DataField       =   "edif_descripcion"
            Caption         =   "Nombre_Edificio"
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
            DataField       =   "edif_tipo"
            Caption         =   "Tipo_Edificio"
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
         BeginProperty Column05 
            DataField       =   "sigla_emprea"
            Caption         =   "Empresa"
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
            DataField       =   "unimed_codigo"
            Caption         =   "Periodicidad"
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
            DataField       =   "mes_par_impar"
            Caption         =   "Insumo3y4"
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
            DataField       =   "Gratuito"
            Caption         =   "Gratuito"
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
            DataField       =   "fecha_fin_max"
            Caption         =   "F.Fin.Ultima"
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
            DataField       =   "venta_codigo"
            Caption         =   "#Venta"
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
            DataField       =   "estado_activo"
            Caption         =   "Estado.H"
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
            DataField       =   "estado_codigo"
            Caption         =   "Estado.A"
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
         BeginProperty Column13 
            DataField       =   "observaciones"
            Caption         =   "Observaciones"
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
            DataField       =   "zona_denominacion"
            Caption         =   "Zona.Geografica"
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
            DataField       =   "calle_tipo"
            Caption         =   "Via"
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
            DataField       =   "calle_denominacion"
            Caption         =   "Nombre.Calle, Av, Plaza...."
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
            DataField       =   "solicitud_tipo"
            Caption         =   "Tipo"
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
            DataField       =   "beneficiario_tecnico1"
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
         BeginProperty Column19 
            DataField       =   "beneficiario_tecnico2"
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
         BeginProperty Column20 
            DataField       =   "beneficiario_cobrador"
            Caption         =   "Cobrador"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   3945.26
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column10 
               Alignment       =   2
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2310.236
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   2190.047
            EndProperty
            BeginProperty Column15 
               Alignment       =   2
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   2489.953
            EndProperty
            BeginProperty Column17 
               Alignment       =   2
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column18 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column19 
               Locked          =   -1  'True
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1319.811
            EndProperty
         EndProperty
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
      Height          =   3240
      Left            =   0
      TabIndex        =   23
      Top             =   720
      Width           =   6180
      Begin VB.OptionButton OptFilGral2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TODOS"
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
         TabIndex        =   26
         Top             =   2955
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptFilGral1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "VIGENTES"
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
         TabIndex        =   25
         Top             =   2955
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Ado_datos 
         Height          =   330
         Left            =   120
         Top             =   2880
         Width           =   5955
         _ExtentX        =   10504
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
         Height          =   2610
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5960
         _ExtentX        =   10504
         _ExtentY        =   4604
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
            DataField       =   "zpiloto_codigo"
            Caption         =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "zpiloto_descripcion"
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
            DataField       =   "munic_codigo"
            Caption         =   "Municipio"
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
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               Object.Visible         =   0   'False
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   2489.953
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_datos1 
      Height          =   330
      Left            =   8760
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Ado_datos2 
      Height          =   330
      Left            =   10920
      Top             =   9240
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
      Left            =   13080
      Top             =   9240
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
   Begin Crystal.CrystalReport CR01 
      Left            =   4560
      Top             =   9600
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
      Left            =   8760
      Top             =   8520
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
   Begin MSAdodcLib.Adodc Ado_datos4 
      Height          =   330
      Left            =   2280
      Top             =   9600
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
      Left            =   120
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Ado_datos6 
      Height          =   330
      Left            =   2280
      Top             =   9240
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
      Left            =   4440
      Top             =   9240
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
      Left            =   6600
      Top             =   9240
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
   Begin MSAdodcLib.Adodc Ado_datos9 
      Height          =   330
      Left            =   120
      Top             =   9600
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
   Begin Crystal.CrystalReport CR02 
      Left            =   5160
      Top             =   9600
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
End
Attribute VB_Name = "tw_organizacion_zonas"
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
Dim rs_datos5 As New ADODB.Recordset
Dim rs_datos6 As New ADODB.Recordset
Dim rs_datos7 As New ADODB.Recordset
Dim rs_datos8 As New ADODB.Recordset
Dim rs_datos9 As New ADODB.Recordset

Dim rsNada As New ADODB.Recordset

Dim rs_det1 As New ADODB.Recordset

Dim rs_aux0 As New ADODB.Recordset
Dim rs_aux1 As New ADODB.Recordset
Dim rs_aux2 As New ADODB.Recordset
Dim rs_aux3 As New ADODB.Recordset
Dim rs_aux4 As New ADODB.Recordset
Dim rs_aux5 As New ADODB.Recordset
Dim rs_aux6 As New ADODB.Recordset
Dim rs_aux7 As New ADODB.Recordset
Dim rs_aux18 As New ADODB.Recordset
Dim rs_aux21 As New ADODB.Recordset

'Dim CAMPOS As ADODB.Field
'BUSCADOR
Dim ClBuscaGrid As ClBuscaEnGridExterno
'Dim queryinicial As String

'OTROS
Dim VAR_MOD, VAR_MOD1, VAR_MOD2 As String
Dim SQL_FOR As String
Dim sql As String
Dim sino As String
Dim NombreCarpeta, e As String
Dim parametro As String
Dim var_titulo As String
Dim VAR_SubTitulo As String
Dim var_cod, VAR_GES As String
Dim VAR_VAL, VAR_ARCH, VAR_ARCH2 As String
Dim VAR_SW, DIA_ORDEN As String
Dim VAR_LUN, VAR_PRIM As String

Dim imag2 As Long

Dim VAR_AUX, VAR_CONT2 As Double

Dim var_campoc31, var_campoc32, var_campoc33, var_campoc34 As Double
Dim var_campod11, var_campod12, var_campod13, var_campod14 As Double
Dim var_campoe11, var_campoe12, var_campoe13, var_campoe14 As Double
Dim var_campoe21, var_campoe22, var_campoe23, var_campoe24 As Double
Dim var_campoe31, var_campoe32, var_campoe33, var_campoe34 As Double
Dim var_campoe41, var_campoe42, var_campoe43, var_campoe44 As Double
Dim var_campog11, var_campog12, var_campog13, var_campog14 As Double
Dim var_campog21, var_campog22, var_campog23, var_campog24 As Double

Dim VAR_5, VAR_6, VAR_7, VAR_8 As String
Dim VAR_PROY2, VAR_EDIF As String
Dim VAR_DA, VAR_UORIGEN, VAR_DPTO As String
Dim VAR_IMPAR, VAR_MED, VAR_COD4, MControl As String
Dim VAR_EQP, VAR_UNITEC, VAR_OBS, VAR_SOLCODIGO As String

Dim VAR_CONT, VAR_MES, var_cod5 As Integer
Dim VAR_ZPILOTO, VAR_ZONA, VAR_FMPLAN, VAR_FMES As Integer
Dim VAR_ORDEN, VAR_EMPRESA, VAR_TIPO, VAR_SOL, VAR_AUX2 As Integer

Dim FInicio, FFin, FControl As Date

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
'        Call ABRIR_TABLA_AUX2
        If Ado_datos.Recordset.RecordCount > 0 Then
            Select Case Ado_datos.Recordset!mes_par_impar
                Case 1
                    'Programar Meses IMPARES
                    VAR_IMPAR = "1"
                    Option2.Value = False
                    Option1.Value = True
                    'LblParImpar = "MESES IMPARES"
                Case 2
                    'PROGRAMAR en Meses PARES
                    VAR_IMPAR = "2"
                    Option2.Value = True
                    Option1.Value = False
                    'LblParImpar = "MESES PARES"
                Case Else
                    'NO ASIGNADO
                    VAR_IMPAR = "0"
                    Option2.Value = False
                    Option1.Value = False
'                    ¿LblParImpar = "NO ASIGNADO"
            End Select
            If Ado_datos.Recordset!zpiloto_codigo <> 0 Then
                'Actualiza tc_zona_piloto_edif  - CONTRATOS VENTAS NUEVAS Y ALCANCE
                'db.Execute "UPDATE tc_zona_piloto_edif SET Gratuito = 'XX' where Zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
                Set rs_aux5 = New ADODB.Recordset
                If rs_aux5.State = 1 Then rs_aux5.Close
                'rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX_ALCANCE WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' ", db, adOpenStatic
                rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX_ALCANCE WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' order by venta_fecha_fin ", db, adOpenStatic
                If rs_aux5.RecordCount > 0 Then
                    rs_aux5.MoveFirst
                    While Not rs_aux5.EOF
                        'db.Execute "UPDATE ao_ventas_cabecera SET ao_ventas_cabecera.unimed_codigo_tec = tc_zona_piloto_edif.unimed_codigo FROM ao_ventas_cabecera INNER JOIN tc_zona_piloto_edif ON ao_ventas_cabecera.edif_codigo = tc_zona_piloto_edif.edif_codigo where ao_ventas_cabecera.venta_codigo = " & rs_aux5!venta_codigo & " "
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        'rs_aux6.Open "Select * from ao_ventas_cabecera where venta_fecha_fin = '" & rs_aux5!venta_fecha_fin & "' and edif_codigo = '" & rs_aux5!EDIF_CODIGO & "' and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND estado_codigo = 'APR' ", db, adOpenStatic
                        rs_aux6.Open "Select * from aV_ventas_alcance where venta_codigo = " & rs_aux5!venta_codigo & " and (unidad_codigo='DVTA' OR unidad_codigo LIKE '%COM%' )  ", db, adOpenStatic
                        If rs_aux6.RecordCount > 0 Then
                            db.Execute "UPDATE tc_zona_piloto_edif SET codigo_empresa= " & rs_aux6!codigo_empresa & ", unimed_codigo = 'MES', solicitud_tipo = '6', fecha_fin_max = '" & rs_aux6!fecha_fin_real & "', Gratuito = 'SI', mes_par_impar = '" & VAR_IMPAR & "', venta_codigo = " & rs_aux5!venta_codigo & "  WHERE edif_codigo = '" & rs_aux6!edif_codigo & "'  "
                        End If
                        rs_aux5.MoveNext
                    Wend
                End If
                'Actualiza tc_zona_piloto_edif  - CONTRATOS MANTENIMIENTO
                'db.Execute "UPDATE tc_zona_piloto_edif SET tc_zona_piloto_edif.unimed_codigo_tec = ao_ventas_cabecera.unimed_codigo FROM ao_ventas_cabecera INNER JOIN tc_zona_piloto_edif ON ao_ventas_cabecera.edif_codigo = tc_zona_piloto_edif.edif_codigo where ao_ventas_cabecera.venta_codigo = " & rs_aux5!venta_codigo & " and ao_ventas_cabecera.unimed_codigo_tec is null "
                Set rs_aux5 = New ADODB.Recordset
                If rs_aux5.State = 1 Then rs_aux5.Close
                rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' order by venta_fecha_fin ", db, adOpenStatic
                If rs_aux5.RecordCount > 0 Then
                    db.Execute "UPDATE tc_zona_piloto_edif SET tc_zona_piloto_edif.unimed_codigo = ao_ventas_cabecera.unimed_codigo FROM ao_ventas_cabecera INNER JOIN tc_zona_piloto_edif ON ao_ventas_cabecera.edif_codigo = tc_zona_piloto_edif.edif_codigo where ao_ventas_cabecera.venta_codigo = " & rs_aux5!venta_codigo & " and tc_zona_piloto_edif.unimed_codigo is null "
                    rs_aux5.MoveFirst
                    While Not rs_aux5.EOF
                        'db.Execute "UPDATE "
                        db.Execute "UPDATE ao_ventas_cabecera SET ao_ventas_cabecera.unimed_codigo_tec = tc_zona_piloto_edif.unimed_codigo FROM ao_ventas_cabecera INNER JOIN tc_zona_piloto_edif ON ao_ventas_cabecera.edif_codigo = tc_zona_piloto_edif.edif_codigo where ao_ventas_cabecera.venta_codigo = " & rs_aux5!venta_codigo & " and ao_ventas_cabecera.unimed_codigo_tec is null  "
                        Set rs_aux6 = New ADODB.Recordset
                        If rs_aux6.State = 1 Then rs_aux6.Close
                        'rs_aux6.Open "Select * from ao_ventas_cabecera where venta_fecha_fin = '" & rs_aux5!venta_fecha_fin & "' and edif_codigo = '" & rs_aux5!EDIF_CODIGO & "' and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND estado_codigo = 'APR' ", db, adOpenStatic
                        rs_aux6.Open "Select * from av_ventas_cabecera_mant where venta_codigo = " & rs_aux5!venta_codigo & " ", db, adOpenStatic
                        If rs_aux6.RecordCount > 0 Then

                            db.Execute "UPDATE tc_zona_piloto_edif SET codigo_empresa= " & rs_aux6!codigo_empresa & ", unimed_codigo = '" & IIf(IsNull(rs_aux6!unimed_codigo_tec), "MES", rs_aux6!unimed_codigo_tec) & "', solicitud_tipo = " & rs_aux5!solicitud_tipo & ", fecha_fin_max = '" & rs_aux5!venta_fecha_fin & "', Gratuito = 'NO', mes_par_impar = '" & VAR_IMPAR & "', venta_codigo = " & rs_aux5!venta_codigo & "  WHERE edif_codigo = '" & rs_aux6!edif_codigo & "'  "
                        End If
                        rs_aux5.MoveNext
                    Wend
                End If
'                'Actualiza tc_zona_piloto_edif  - CONTRATOS VENTAS NUEVAS Y ALCANCE
'                db.Execute "UPDATE tc_zona_piloto_edif SET Gratuito = 'XX' where Zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
'                Set rs_aux5 = New ADODB.Recordset
'                If rs_aux5.State = 1 Then rs_aux5.Close
'                'rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX_ALCANCE WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' ", db, adOpenStatic
'                rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX_GRAL WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' ", db, adOpenStatic
'                If rs_aux5.RecordCount > 0 Then
'                    rs_aux5.MoveFirst
'                    While Not rs_aux5.EOF
'                        'db.Execute "UPDATE ao_ventas_cabecera SET ao_ventas_cabecera.unimed_codigo_tec = tc_zona_piloto_edif.unimed_codigo FROM ao_ventas_cabecera INNER JOIN tc_zona_piloto_edif ON ao_ventas_cabecera.edif_codigo = tc_zona_piloto_edif.edif_codigo where ao_ventas_cabecera.venta_codigo = " & rs_aux5!venta_codigo & " "
'                        Set rs_aux6 = New ADODB.Recordset
'                        If rs_aux6.State = 1 Then rs_aux6.Close
'                        'rs_aux6.Open "Select * from ao_ventas_cabecera where venta_fecha_fin = '" & rs_aux5!venta_fecha_fin & "' and edif_codigo = '" & rs_aux5!EDIF_CODIGO & "' and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND estado_codigo = 'APR' ", db, adOpenStatic
'                        rs_aux6.Open "Select * from aV_ventas_alcance where venta_codigo = " & rs_aux5!venta_codigo & " and (unidad_codigo='DVTA' OR unidad_codigo LIKE '%COM%' ) ", db, adOpenStatic
'                        If rs_aux6.RecordCount > 0 Then
'                            db.Execute "UPDATE tc_zona_piloto_edif SET codigo_empresa= " & rs_aux6!codigo_empresa & ", unimed_codigo = 'MES', solicitud_tipo = '6', fecha_fin_max = '" & rs_aux6!fecha_fin_real & "', Gratuito = 'SI', mes_par_impar = '" & VAR_IMPAR & "', venta_codigo = " & rs_aux5!venta_codigo & "  WHERE edif_codigo = '" & rs_aux6!EDIF_CODIGO & "'  "
'                        End If
'                        rs_aux5.MoveNext
'                    Wend
'                End If
            End If
            Call ABRIR_TABLA_DET
'            If lbl_texto1.Caption <> "" And lbl_texto1.Caption <> "0" Then
'                lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
'            End If
            'mes2 = MonthName(Month(DTPFec_Inicio.Value))
        End If
    Else
        'Set rs_det1 = New ADODB.Recordset
        Set dg_det1.DataSource = rsNada
        'Set DtgLaborales.DataSource = rsNada
    End If
End Sub

Private Sub BtnAddDetalle_Click()
  marca1 = Ado_datos.Recordset.Bookmark
  If Ado_datos.Recordset!estado_codigo <> "ANL" Then
    swnuevo = 1
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    dg_det1.Enabled = False
    Fra_datos.Visible = False
    FraDet2.Visible = True
    
    fra_opciones_det.Visible = False
    If VAR_UORIGEN = "DNINS" Then
        lbl_campo6.Caption = "Tecnico Instalaciones"
    Else
        lbl_campo6.Caption = "Tecnico Mantenimiento"
    End If
    
    Call ABRIR_DET
    'Ado_detalle1.Recordset.AddNew
    dtc_codigo6.Text = dtc_codigo4.Text
    dtc_codigo7.Text = dtc_codigo4.Text
    dtc_desc6.Text = dtc_desc4.Text
    dtc_desc7.Text = dtc_desc4.Text
    lbl_orden_camb.Visible = False
    cmd_campo2.Visible = False
'    dtc_codigo5.Locked = False
    dtc_desc5.Locked = False
  Else
    MsgBox "No se puede Adicionar un nuevo registro, porque este ya está Aprobado!! ", vbExclamation
  End If
  If Ado_datos.Recordset!estado_codigo = "REG" Then
'    Call OptFilGral1_Click
  Else
    Call OptFilGral2_Click
    Ado_datos.Recordset.Move marca1 - 1
  End If
  'Call ABRIR_TABLA_DET
End Sub

Private Sub ABRIR_DET()
    'gc_edificaciones
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    If VAR_UORIGEN = "DNINS" Then
        rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' and tomado = 'N' order by edif_descripcion", db, adOpenStatic
    Else
        rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' AND depto_codigo = '" & Ado_datos.Recordset!depto_codigo & "' and tomado = 'N' order by edif_descripcion", db, adOpenStatic
    End If
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub BtnAddDetalle3_Click()
If Ado_detalle1.Recordset!estado_codigo = "REG" Then
    'INI ENVIA A OTRA ZONA
    fra_opciones.Visible = False
    fra_opciones_det.Visible = False
    FraGrabarCancelar.Visible = False
    FraNewZona.Visible = True
    VAR_ZPILOTO = Ado_datos.Recordset!zpiloto_codigo
    Set rs_aux7 = New ADODB.Recordset
    If rs_aux7.State = 1 Then rs_aux7.Close
    'If Ado_datos.Recordset!zpiloto_codigo = "37" Then
        rs_aux7.Open "Select * from tc_zonas_piloto WHERE zpiloto_codigo <> " & VAR_ZPILOTO & " and depto_codigo = '" & Left(Ado_detalle1.Recordset!edif_codigo, 1) & "'  ", db, adOpenStatic
    'Else
    '    rs_aux7.Open "Select * from tc_zonas_piloto WHERE zpiloto_codigo = '37'  ", db, adOpenStatic
    'End If
    Set Ado_datos2.Recordset = rs_aux7
    If rs_aux7.RecordCount > 0 Then
        dtc_desc10.BoundText = dtc_codigo10.BoundText
        VAR_ZONA = rs_aux7!zpiloto_codigo
    Else
        VAR_ZONA = 0
    End If

    'FIN ENVIA A OTRA ZONA
  Else
      MsgBox "No se puede ENVIAR A OTRA ZONA DEFINITIVAMENTE, porque el Edificio está en Ejecución, puede hacerlo Sólo para un mes en Cronograma Mensual ...", vbExclamation, "Validación de Registro"
  End If
End Sub

Private Sub BtnAnlDetalle_Click()
   If Ado_detalle1.Recordset("estado_activo") = "REG" Then
      sino = MsgBox("Está Seguro de Anular este registro ? (Este ya no será considerado en la presente Zona) ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
        cmd_campo2.Text = Ado_detalle1.Recordset!zona_edif_orden
        
        db.Execute "update gc_edificaciones set tomado= 'N' where edif_codigo = '" & dtc_codigo5.Text & "' "
        'If cmd_campo2.Text <> "0" Then
        db.Execute "DELETE tc_zona_piloto_edif WHERE zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND edif_codigo = '" & dtc_codigo5.Text & "' "
'            db.Execute "update tc_zona_piloto_edif set zorden_cambio = zona_edif_orden - 1 where zona_edif_orden >= " & cmd_campo2.Text & "  and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & ""
'            db.Execute "update tc_zona_piloto_edif set zona_edif_orden = zorden_cambio  where zorden_cambio > '0'  and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " "
'            db.Execute "delete tc_zona_piloto_edif where correlativo = " & Text1.Text & " "
'            db.Execute "update tc_zona_piloto_edif set zorden_cambio = '0'  where zorden_cambio > '0'"
'            db.Execute "update tc_zona_piloto_edif set estado_activo = 'ANL'"
        'End If
        Call ABRIR_TABLA_DET
      End If
   Else
      MsgBox "No se puede ANULAR, el registro ya fue APROBADO o ya fue ANULADO anteriormente ...", vbExclamation, "Validación de Registro"
   End If
End Sub

Private Sub BtnAñadir_Click()
    If glusuario = "CCRUZ" Then         'Or glusuario = "LNAVA"
        MsgBox "el Usuario NO tiene acceso, consulte con el Administrador del Sistema!! ", vbExclamation
        Exit Sub
    End If
  On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    'If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fra_opciones.Visible = False
        FraGrabarCancelar.Visible = True
        dg_datos.Enabled = False
        VAR_SW = "ADD"
        'fra_opciones_det.Visible = False
        FraDet1.Visible = False
        Ado_datos.Recordset.AddNew
    '    BtnVer.Visible = True
    'Else
    '  MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    'End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnAprobar_Click()
  On Error GoTo UpdateErr
  Set rs_aux2 = New ADODB.Recordset
  If rs_aux2.State = 1 Then rs_aux2.Close
  rs_aux2.Open "select * from tv_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' order by zona_edif_orden ", db, adOpenKeyset, adLockOptimistic, adCmdText
  If rs_aux2.RecordCount > 0 Then
   If rs_datos!estado_codigo = "REG" Then
      sino = MsgBox("Está Seguro de APROBAR el Registro ? ", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "APR"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
       MsgBox "No se puede APROBAR un registro Anulado (ANL) o Aprobado (APR) anteriormente ...", vbExclamation, "Validación de Registro"
   End If
  Else
    MsgBox "No se puede APROBAR debe asignar por lo menos un Edificio a esta Zona ...", vbExclamation, "Validación de Registro"
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
        rs_datos.CancelUpdate
        Call ABRIR_TABLA
        'rs_datos.MoveFirst
        'mbDataChanged = False
        Fra_datos.Enabled = False
        fra_opciones.Visible = True
        FraGrabarCancelar.Visible = False
        dg_datos.Enabled = True
        VAR_SW = ""
        'fra_opciones_det.Visible = True
        FraDet1.Visible = True
    End If
End Sub

Private Sub BtnCancelar3_Click()
    fra_opciones.Visible = True
    fra_opciones_det.Visible = True
    FraGrabarCancelar.Visible = True
    FraNewZona.Visible = False
End Sub

Private Sub BtnCancelarDet_Click()
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    dg_det1.Enabled = True
    Fra_datos.Visible = True
    FraDet2.Visible = False
    FraDet1.Visible = True
    If Ado_detalle1.Recordset.RecordCount > 0 Then
        Ado_detalle1.Recordset.CancelUpdate
    End If
    fra_opciones_det.Visible = True
    
    dtc_aux5.Visible = False
    dtc_desc5.Visible = True
End Sub

Private Sub btnEliminar_Click()
  On Error GoTo UpdateErr
   If ExisteReg(Ado_datos.Recordset!zpiloto_codigo) Then MsgBox "No se puede ANULAR el Registro que ya fue utilizado ..", vbInformation + vbOKOnly, "Atención": Exit Sub
   If rs_datos!estado_codigo = "APR" Then
      sino = MsgBox("Está Seguro de ANULAR el Registro? Este ya no podrá ser utilizado...", vbYesNo + vbQuestion, "Atención")
      If sino = vbYes Then
         rs_datos!estado_codigo = "ANL"
         rs_datos!fecha_registro = Date
         rs_datos!usr_codigo = glusuario
         rs_datos.UpdateBatch adAffectAll
      End If
   Else
      MsgBox "No se puede ANULAR un registro Elaborado (REG) o Anulado (ANL) ...", vbExclamation, "Validación de Registro"
   End If
   Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub CRONO2_ALCANCE()
''    Set rs_aux5 = New ADODB.Recordset
''    If rs_aux5.State = 1 Then rs_aux5.Close
''    rs_aux5.Open "select * from av_ventas_alcance where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
''    'Set AdoAux.Recordset = rsAuxDetalle
''    If rs_aux5.RecordCount > 0 Then
'    If IsNull(Ado_datos.Recordset!fecha_inicio_real) Then                   ' Is Null
'        MsgBox "Debe registrar la Fecha Inicio Real, verifique y vuelva a intentar ...", , "Atencion"
'        Exit Sub
'    End If
'      CONT2 = 1
'      'FInicio = DTPfechasol.Value                        'Fecha Inicio Alcance
'      FInicio = Ado_datos.Recordset!fecha_inicio_real                       '
''      CANTOT = rs_aux5!venta_cantidad_total
'      gestion0 = Year(FInicio)                              'Ado_datos.Recordset!ges_gestion
'      VAR_BENEF = Ado_datos.Recordset!beneficiario_codigo
'      VAR_MED2 = "MES"                                          'Ado_datos.Recordset!unimed_codigo_cobr
'      'MES = DateDiff("M", fecha, Date)
'      VAR_COBR2 = DateDiff("M", FInicio, Ado_datos.Recordset!fecha_fin_real) + 1
'      MControl = UCase(MonthName(Month(FInicio)))                        'Ado_datos.Recordset!mes_inicio_crono
'      VAR_MES2 = Month(FInicio)
'      FControl = RTrim("01/" + RTrim(Str(Month(FInicio))) + "/" + Str(Year(FInicio)))
'      VAR_FCONTROL = Format(FControl, "dd/mm/yyyy")
'      FControl = VAR_FCONTROL
'      CONT3 = 0
'      CONT4 = 0
'      CONT_MED = 1              'MES = 1 (Mensual)
'      Set rs_aux2 = New ADODB.Recordset
'      If rs_aux2.State = 1 Then rs_aux2.Close
'      rs_aux2.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "  ", db, adOpenKeyset, adLockOptimistic
'      If rs_aux2.RecordCount = 0 Then
'        db.Execute "UPDATE ao_ventas_cabecera SET correl_cobro_prog = '0' WHERE  venta_codigo = " & NumComp & " "
'        correldet2 = "0"
'      End If
'      While (CONT2 <= VAR_COBR2)
'        Set rs_aux2 = New ADODB.Recordset
'        If rs_aux2.State = 1 Then rs_aux2.Close
'        rs_aux2.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "  ", db, adOpenKeyset, adLockOptimistic
'        If rs_aux2.RecordCount > 0 And corrprog >= VAR_COBR2 Then
'            MsgBox "El Cronograma ya fue generado... ", , "Atención"
'            CONT2 = CONT2 + 1
'        Else
'           'wwwwwwwwwwwwwwwwwwwwww
''          If correldet2 > 0 Then
''            correldet2 = Ado_datos.Recordset!correl_cobro_prog + 1                          'rs_aux5!correl_cobro_prog + 1
''          End If
'          correldet2 = correldet2 + 1
'          corrprog = correldet2
'          db.Execute "UPDATE ao_ventas_cabecera SET correl_cobro_prog = " & corrprog & " WHERE  venta_codigo = " & NumComp & " "
'
'          Set rs_aux8 = New ADODB.Recordset
'          If rs_aux8.State = 1 Then rs_aux8.Close
'          rs_aux8.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(VAR_FCONTROL) & "'  AND MONTH(cobranza_fecha_prog) = '" & Month(VAR_FCONTROL) & "'  ", db, adOpenKeyset, adLockOptimistic
'          If rs_aux8.RecordCount = 0 Then
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            'rs_aux2.AddNew
'            'rs_aux2!cobranza_prog_codigo = correldet2
'            'rs_aux2!beneficiario_codigo = VAR_BENEF                   'Codigo Beneficiario/Cliente
'            ''OJO MODIFICAR COBRADOR - JQA 03-ENE-2015
'            'rs_aux2!beneficiario_codigo_resp = IIf(dtc_codigo5.Text = "", "4333735", dtc_codigo5.Text)      ''Codigo Cobrador
'            'Ado_datos.Recordset!correl_cobro_prog = corrprog
'
'            fdia = Day(VAR_FCONTROL)
'            fanio = Year(VAR_FCONTROL)
'            'CONT3 = CONT2 * CONT_MED
'            CONT3 = 1
'            While (CONT3 <= CONT_MED)
'                fmes = Month(VAR_FCONTROL)
'                Select Case fmes
'                    Case 2
'                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
'                            Dias_Mes = 29
'                        Else
'                            Dias_Mes = 28
'                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
'                        End If
'                    Case 1, 3, 5, 7, 8, 10, 12
'                        Dias_Mes = 31
'                    Case 4, 6, 9, 11
'                        Dias_Mes = 30
'                End Select
'                If Val(VAR_MES2) = Month(VAR_FCONTROL) Then
'                    'rs_aux2!cobranza_fecha_prog = FControl
'                    'rs_aux2!cobranza_fecha_cobro = FControl + 20
'                    FControl = Format(VAR_FCONTROL, "dd/mm/yyyy")
'                    VAR_MES2 = VAR_MES2 + CONT_MED
'                    If Val(VAR_MES2) > 12 Then
'                        VAR_MES2 = Val(VAR_MES2) - 12
'                    End If
'                End If
'                'FControl = FControl + Dias_Mes
'                VAR_FCONTROL = CDate(VAR_FCONTROL) + Dias_Mes
'                CONT3 = CONT3 + 1
'                CONT4 = CONT4 + Dias_Mes
'            Wend
'            'FControl = Str(fdia) + "/" + Str(fmes) + "/" + Str(fanio)
'            'rs_aux2!cobranza_fecha_prog = FInicio + (30 * CONT2)
'            'rs_aux2!cobranza_fecha_prog = FControl
'            'If Ado_datos.Recordset!cobranza_fecha_prog = Null Then
'                'rs_aux2!cobranza_fecha_prog = Date
'                'FControl = Date
'            '    VAR_FCONTROL = Date
'            'End If
'            'rs_aux2!gestion = Year(rs_aux2!cobranza_fecha_prog)
'            'rs_aux2!cobranza_mes = Month(rs_aux2!cobranza_fecha_prog)
'
'            ''VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), Date, rs_aux2!cobranza_fecha_prog)))
'            'VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
'
'            'CONT2 = CONT2 + 1
'            'rs_aux2!cobranza_requisito_plazo = "S"
'            ''rs_aux2!cobranza_concepto_plazo = "CONFORMIDAD DEL SERVICIO"
'            'If rs_aux2!cobranza_programada_bs <> 0 Then
'            '    rs_aux2!Literal = Literal(CStr(rs_aux2!cobranza_programada_bs)) + " BOLIVIANOS"
'            'End If
'            'rs_aux2!proceso_codigo = "TEC"
'            'rs_aux2!subproceso_codigo = "TEC-02"
'            'rs_aux2!etapa_codigo = "TEC-02-02"
'            'rs_aux2!clasif_codigo = "TEC"
'            'rs_aux2!doc_codigo = "R-105"    ' R-307 Certificado de Mantenimiento ' Colocar en la conformidad
'            'rs_aux2!doc_numero = "0"        'NumComp
'            'rs_aux2!poa_codigo = "3.2.3"
'            'rs_aux2!estado_codigo = "REG"
'            'rs_aux2!usr_codigo = glusuario
'            'rs_aux2!fecha_registro = Format(Date, "dd/mm/yyyy")
'            'rs_aux2!hora_registro = Format(Time, "hh:mm:ss")
'            'rs_aux2!correl_ac = 0
'            'rs_aux2!estado_ac = "REG"
'            'rs_aux2.Update
'            ' JQA 2022-10-22
'            VAR_FEC2 = UCase(MonthName(Month(FControl)))
'            CONT2 = CONT2 + 1
'            VAR_GLOSA = "MTTO. GRATUITO POR LA GESTION: " + Str(Year(CDate(FControl))) + " - MES: " + VAR_FEC2
'            VAR_SOLTIPO = Ado_datos.Recordset!solicitud_tipo
'            'VAR_ZPILOTO = IIf(IsNull(Ado_datos.Recordset!zpiloto_codigo), VAR_ZPILOTO, Ado_datos.Recordset!zpiloto_codigo)
'            'cobranza_fecha_conformidad, correl_prog , usr_aprueba, fecha_aprueba
'            db.Execute "INSERT INTO ao_ventas_cobranza_inst (venta_codigo, cobranza_prog_codigo, venta_codigo_new, ges_gestion, cobranza_fecha_prog, Gestion, cobranza_mes, fmes_plan, doc_numero_crono, edif_codigo, zpiloto_codigo, " & _
'                       "  trans_codigo, estado_codigo, usr_codigo, fecha_registro, Observaciones)  " & _
'                       " VALUES (" & NumComp & ", " & correldet2 & ", '0', '" & gestion0 & "', '" & FControl & "', '" & Year(FControl) & "', '" & Month(FControl) & "', '0', '0', '" & GlEdificio & "', " & VAR_ZPILOTO & ",  " & _
'                       " '42', 'REG', '" & glusuario & "', '" & Date & "', '" & VAR_GLOSA & "' )"
'
'            'Asigna IdCrono (fmes_plan)
'            'VAR_ZPILOTO = Ado_datos.Recordset!zpiloto_codigo
'            Set rs_aux18 = New ADODB.Recordset
'            If rs_aux18.State = 1 Then rs_aux18.Close
'            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZPILOTO & " AND ges_gestion = '" & Year(FControl) & "' AND fmes_correl = " & Month(FControl) & "  ", db, adOpenKeyset, adLockOptimistic
'            If rs_aux18.RecordCount > 0 Then
'                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
'            Else
'                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
'            End If
'            '
'          Else
'            db.Execute "UPDATE ao_ventas_cobranza_inst SET gestion = '" & Year(rs_aux2!cobranza_fecha_prog) & "', cobranza_mes = '" & Month(rs_aux2!cobranza_fecha_prog) & "' where  venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & ""
'            db.Execute "UPDATE ao_ventas_cobranza_inst SET cobranza_prog_codigo = " & correldet2 & " where venta_codigo = " & NumComp & " and YEAR(cobranza_fecha_prog) ='" & Year(FControl) & "'  AND MONTH(cobranza_fecha_prog) ='" & Month(FControl) & "'  "
'
'            'Asigna IdCrono (fmes_plan) WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            Set rs_aux18 = New ADODB.Recordset
'            If rs_aux18.State = 1 Then rs_aux18.Close
'            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZPILOTO & " AND ges_gestion = '" & rs_aux2!gestion & "' AND fmes_correl = " & rs_aux2!cobranza_mes & "  ", db, adOpenKeyset, adLockOptimistic
'            If rs_aux18.RecordCount > 0 Then
'                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
'            Else
'                db.Execute "update ao_ventas_cobranza_inst set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
'            End If
'            'WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW
'            fdia = Day(FControl)
'            fanio = Year(FControl)
'            'CONT3 = CONT2 * CONT_MED
'            CONT3 = 1
'            While (CONT3 <= CONT_MED)
'                fmes = Month(FControl)
'                Select Case fmes
'                    Case 2
'                        If fanio = "2012" Or fanio = "2016" Or fanio = "2020" Or fanio = "2024" Then
'                            Dias_Mes = 29
'                        Else
'                            Dias_Mes = 28
'                            'Dias_Del_Mes = IIf(saltarYear(Fecha), 29, 28)
'                        End If
'                    Case 1, 3, 5, 7, 8, 10, 12
'                        Dias_Mes = 31
'                    Case 4, 6, 9, 11
'                        Dias_Mes = 30
'                End Select
'                If Val(VAR_MES2) = Month(FControl) Then
'                    rs_aux2!cobranza_fecha_prog = FControl
'                    'rs_aux2!cobranza_fecha_conformidad = FControl + 10
'                    rs_aux2!cobranza_fecha_cobro = FControl + 20
'                    VAR_MES2 = VAR_MES2 + CONT_MED
'                    If Val(VAR_MES2) > 12 Then
'                        VAR_MES2 = Val(VAR_MES2) - 12
'                    End If
'                End If
'                FControl = FControl + Dias_Mes
'                CONT3 = CONT3 + 1
'                CONT4 = CONT4 + Dias_Mes
'            Wend
'            VAR_FEC2 = MonthName(Month(IIf(IsNull(rs_aux2!cobranza_fecha_prog), FControl, rs_aux2!cobranza_fecha_prog)))
'            CONT2 = CONT2 + 1
'          End If
'        End If
'      Wend
'      MsgBox "El Cronograma fue Generado Exitosamente... ", , "Atención"
'      'db.Execute "update ao_ventas_cabecera set ao_ventas_cabecera.estado_codigo_verif = 'APR' Where ao_ventas_cabecera.venta_codigo = " & NumComp & " "
'      If corrprog > 0 Then
'        db.Execute "update ao_ventas_cabecera set correl_cobro_prog = '" & corrprog & "' "
'        db.Execute "update ao_ventas_cabecera set venta_plazo_dias_calendario = " & CONT4 & " "
'      End If
'
''    Else
''       MsgBox "Error Verifique la Venta de Productos..."
''    End If
End Sub

Private Sub BtnGraba3_Click()
  If Ado_detalle1.Recordset!estado_codigo = "APR" Then
    MsgBox "No se puede enviar a Otra Zona Piloto DEFINITIVAMENTE porque ya está en Ejecución, puede enviar en el Cronograma (Solo un mes específico) ..."
    Exit Sub
  End If
  sino = MsgBox("Está seguro de enviar DEFINITIVAMENTE a Otra Zona Piloto ? ", vbYesNo + vbQuestion, "Atención")
  If sino = vbYes Then

    gestion = IIf(CmbGestion.Text < Year(Date), Year(Date), CmbGestion.Text)          '  Ado_detalle1.Recordset!ges_gestion
    VAR_PROY2 = Ado_detalle1.Recordset!edif_codigo
    VAR_EMPRESA = Ado_detalle1.Recordset!codigo_empresa
    VAR_TIPO = Ado_detalle1.Recordset!solicitud_tipo
    NumComp = Ado_detalle1.Recordset!venta_codigo
    VAR_ZONA = dtc_codigo10.Text
    VAR_ZPILOTO = Ado_detalle1.Recordset!zpiloto_codigo
    
    Set rs_aux0 = New ADODB.Recordset
    If rs_aux0.State = 1 Then rs_aux0.Close
    rs_aux0.Open "Select * from gc_edificaciones WHERE edif_codigo = '" & VAR_PROY2 & "'   ", db, adOpenStatic
    If rs_aux0.RecordCount > 0 Then
        VAR_EDIF = rs_aux0!edif_descripcion                      'RTrim(dtc_desc3.Text)          'edif_descripcion
    End If
    VAR_LUN = "SI"                                                  'Ado_datos.Recordset!lunes_cambia
    VAR_PRIM = "SI"                                                 'Ado_datos.Recordset!primero_mes
    VAR_MED = Ado_detalle1.Recordset!unimed_codigo
    DIA_ORDEN = Ado_detalle1.Recordset!zona_edif_orden
    
'    'VAR_EMES = "Error: No se encontró el Mes de Inicio del Cronograma, verifique y vuelva a intentar..."
'    ' jalar ORDEN de tc_zona_piloto_edif
'    Set rs_datos6 = New ADODB.Recordset
'    If rs_datos6.State = 1 Then rs_datos6.Close
'    rs_datos6.Open "Select * from tc_zona_piloto_edif WHERE edif_codigo = '" & VAR_PROY2 & "'    ", db, adOpenStatic
'    If rs_datos6.RecordCount > 0 Then
'        DIA_ORDEN = rs_datos6!zona_edif_orden
'    Else
        Set rs_aux18 = New ADODB.Recordset
        If rs_aux18.State = 1 Then rs_aux18.Close
        rs_aux18.Open "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = " & VAR_ZONA & " ", db, adOpenKeyset, adLockOptimistic
        If rs_aux18.RecordCount > 0 Then
            VAR_ORDEN = IIf(IsNull(rs_aux18!Orden), 1, rs_aux18!Orden + 1)
        Else
            VAR_ORDEN = 1
        End If
                  
'        DIA_ORDEN = "1"
'    End If
    'DIA_ORDEN = Ado_datos.Recordset!zona_edif_orden
    'NumComp = Ado_detalle1.Recordset!venta_codigo
    Set rs_aux5 = New ADODB.Recordset
    If rs_aux5.State = 1 Then rs_aux5.Close
    rs_aux5.Open "select * from av_ventas_alcance where venta_codigo= " & NumComp & "  ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux5.RecordCount > 0 Then
        FInicio = rs_aux5!fecha_inicio_real
        FFin = rs_aux5!fecha_fin_real
        VAR_COD4 = rs_aux5!unidad_codigo
        VAR_SOL = rs_aux5!solicitud_codigo
    End If
    'zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, correlativo, mes_par_impar, observaciones,
    'unimed_codigo, codigo_empresa, solicitud_tipo, Gratuito,    fecha_fin_max , venta_codigo, estado_codigo, estado_activo, fecha_registro, hora_registro, usr_codigo

       db.Execute "INSERT INTO tc_zona_piloto_edif (zpiloto_codigo, edif_codigo, ges_gestion, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, mes_par_impar, observaciones, " & _
                  " estado_codigo , estado_activo, fecha_registro, usr_codigo, unimed_codigo, codigo_empresa, solicitud_tipo, Gratuito, fecha_fin_max, venta_codigo ) " & _
                  " VALUES (" & VAR_ZONA & ", '" & VAR_PROY2 & "', '" & gestion & "',      " & VAR_ORDEN & ",       '0',            '0',                    '0',                    '0',                    '0',            '1',        '',  " & _
                  " 'REG',              'APR', '" & Date & "', '" & glusuario & "', '" & VAR_MED & "', " & VAR_EMPRESA & ", " & VAR_TIPO & ", 'SI', '" & FFin & "', " & NumComp & ")"
    
    'lbl_texto2.Caption = UCase(MonthName(Val(lbl_texto1.Caption)))
    'VAR_MOD = Format(FInicio, "dd/mm/yyyy")
    VAR_MOD = Month(FInicio)
    MControl = UCase(MonthName(VAR_MOD))                                    'mes_inicio_crono
    gestion = Year(FInicio)
    FControl = FInicio
    Set rs_aux1 = New ADODB.Recordset
    rs_aux1.Open "select * from ao_ventas_cobranza_inst where venta_codigo = " & NumComp & "   ", db, adOpenKeyset, adLockBatchOptimistic
    If rs_aux1.RecordCount > 0 Then
        var_cod5 = rs_aux1.RecordCount
        rs_aux1.MoveFirst
        While Not rs_aux1.EOF
            'Asigna IdCrono (fmes_plan)
            Set rs_aux18 = New ADODB.Recordset
            If rs_aux18.State = 1 Then rs_aux18.Close
            rs_aux18.Open "Select fmes_plan from to_cronograma_mensual where zpiloto_codigo = " & VAR_ZONA & " AND ges_gestion = '" & Year(FControl) & "' AND fmes_correl = " & Month(FControl) & "  ", db, adOpenKeyset, adLockOptimistic
            If rs_aux18.RecordCount > 0 Then
                'db.Execute "update ao_ventas_cobranza_inst set fmes_plan = " & rs_aux18!fmes_plan & " where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
                VAR_AUX2 = rs_aux18!fmes_plan
            'Else
            '    'db.Execute "update ao_ventas_cobranza_inst set fmes_plan = '0' where venta_codigo = " & NumComp & " and cobranza_prog_codigo = " & correldet2 & " "
            '    VAR_AUX2 = rs_aux18!fmes_plan
            End If

            Set rs_aux2 = New ADODB.Recordset
            If rs_aux2.State = 1 Then rs_aux2.Close
            rs_aux2.Open "select * from ao_ventas_detalle where venta_codigo = " & NumComp & " and par_codigo = '43340'   ", db, adOpenKeyset, adLockBatchOptimistic
            If rs_aux2.RecordCount > 0 Then
                rs_aux2.MoveFirst
                While Not rs_aux2.EOF
                    'VERIFICA SI EXITE EQUIPO EN ESTE MES
                    Set rs_aux21 = New ADODB.Recordset
                    If rs_aux21.State = 1 Then rs_aux21.Close
                    rs_aux21.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = '" & rs_aux2!bien_codigo & "'  ", db, adOpenKeyset, adLockBatchOptimistic
                    If rs_aux21.RecordCount > 0 Then
                        db.Execute "update to_cronograma_diario set unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                        db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "   "
                        db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux21!dia_correl & " AND horario_codigo = " & rs_aux21!horario_codigo & "  "
                    Else
                        Set rs_aux3 = New ADODB.Recordset
                        If rs_aux3.State = 1 Then rs_aux3.Close
                        rs_aux3.Open "select * from to_cronograma_diario where fmes_plan = " & VAR_AUX2 & " AND bien_codigo = ''  ", db, adOpenKeyset, adLockBatchOptimistic
                        If rs_aux3.RecordCount > 0 Then
                            rs_aux3.MoveFirst
                            'If VAR_COD0 < var_cod5 Then     'And rs_aux3!estado_activo = "REG"
                                'db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_UNITEC & "',  tec_plan_codigo = " & VAR_TECCOD & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "'   WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_codigo = '" & rs_aux2!bien_codigo & "', unidad_codigo_tec = '" & VAR_COD4 & "',  tec_plan_codigo = " & VAR_SOL & ", observaciones = 'HORARIO LABORABLE', edif_descripcion = '" & VAR_EDIF & "', edif_codigo = '" & VAR_PROY2 & "' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                db.Execute "update to_cronograma_diario set bien_orden = " & DIA_ORDEN & " WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  and bien_orden='0' "
                                db.Execute "update to_cronograma_diario set estado_activo = 'APR' WHERE fmes_plan = " & VAR_AUX2 & " AND dia_correl = " & rs_aux3!dia_correl & " AND horario_codigo = " & rs_aux3!horario_codigo & "  "
                                'VAR_COD0 = VAR_COD0 + 1
                                'CONT3 = 1
                                db.Execute "Update ao_ventas_cabecera Set estado_crono = 'APR' Where venta_codigo = " & NumComp & "  "
                                'VAR_EMES = "NADA"
                            'End If
                        Else
                            'POR SI NO TIENE fmes_plan
                        End If
                    End If
                    rs_aux2.MoveNext
                Wend
            rs_aux1.MoveNext
            End If
        Wend
        db.Execute "DELETE tc_zona_piloto_edif WHERE zpiloto_codigo = " & VAR_ZPILOTO & " AND edif_codigo = '" & VAR_PROY2 & "' "
        db.Execute "update to_cronograma_diario set bien_orden = 0, bien_codigo = '', unidad_codigo_tec = '', tec_plan_codigo = 0, edif_descripcion = '', edif_codigo = '', estado_codigo = 'REG' where unidad_codigo_tec = '" & VAR_COD4 & "' AND tec_plan_codigo = " & VAR_SOL & " AND edif_codigo= '" & VAR_PROY2 & "' "
        'db.Execute "UPDATE tc_zona_piloto_edif SET Gratuito = 'SI' WHERE     (solicitud_tipo = 6) AND (Gratuito <> 'SI') AND (venta_codigo IS NOT NULL) "
        MsgBox "El Edificio fue ENVIADO satisfactoriamente a la Zona Piloto DESTINO ...", vbInformation, "Información de Registro"
        
        fra_opciones.Visible = True
        fra_opciones_det.Visible = True
        FraGrabarCancelar.Visible = True
        FraNewZona.Visible = False
        
        Call OptFilGral2_Click
    End If
  End If
End Sub

Private Sub BtnGrabar_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_campos
  If VAR_VAL = "OK" Then
     rs_datos!zpiloto_descripcion = Txt_campo2.Text
     rs_datos!pais_codigo = "BOL"
     rs_datos!depto_codigo = dtc_codigo1.Text
     rs_datos!prov_codigo = dtc_codigo2.Text
     rs_datos!munic_codigo = dtc_codigo3.Text
     rs_datos!zona_codigo = "0"     'dtc_codigo9.Text = txt_codigo.Text
     rs_datos!beneficiario_codigo = dtc_codigo4.Text
     rs_datos!fecha_registro = Date     'no cambia
     rs_datos!usr_codigo = IIf(glusuario = "", "ADMIN", glusuario) 'no cambia
     rs_datos.Update    'Batch 'adAffectAll
     
'     db.Execute "Update to_cronograma_diario Set beneficiario_codigo_resp = " & dtc_codigo4.Text & " Where fmes_plan = '" & Ado_datos.Recordset!fmes_plan & "'   "
     'Call OptFilGral2_Click
     'rs_datos.MoveFirst
'     mbDataChanged = False

     Fra_datos.Enabled = False
     fra_opciones.Visible = True
     FraGrabarCancelar.Visible = False
     dg_datos.Enabled = True
     'dtc_desc1.BackColor = &HFFFFC0
     VAR_SW = ""
     'fra_opciones_det.Visible = True
     FraDet1.Visible = True
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub valida_campos()
  'Valida compos para editables
  If (dtc_codigo1.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo1.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo2 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo2.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo3.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo3.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo4.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo4.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (Txt_campo2.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_zona.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
End Sub

Private Sub valida_det()
  'Valida compos para editables
  If (dtc_codigo5.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo5.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo6 = "") Then
    MsgBox "Debe registrar ... " + lbl_campo6.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo7.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo7.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
  If (dtc_codigo8.Text = "") Then
    MsgBox "Debe registrar ... " + lbl_campo8.Caption, vbCritical + vbExclamation, "Validación de datos"
    VAR_VAL = "ERR"
    Exit Sub
  End If
'  If (Txt_campo1.Text = "") Then
'    MsgBox "Debe registrar ... " + lbl_orden.Caption, vbCritical + vbExclamation, "Validación de datos"
'    VAR_VAL = "ERR"
'    Exit Sub
'  End If
End Sub

Private Sub BtnGrabarDet_Click()
  On Error GoTo UpdateErr
  VAR_VAL = "OK"
  Call valida_det
  If VAR_VAL = "OK" Then
    If Option11.Value = True Then
        'PROGRAMAR en Meses PARES y quitar Mes IMPARES
        VAR_IMPAR = "2"
    Else
        'Programar Meses IMPARES y quitar PARES
        VAR_IMPAR = "1"
    End If
    If swnuevo = 1 Then
        Set rs_aux1 = New ADODB.Recordset
        If rs_aux1.State = 1 Then rs_aux1.Close
        SQL_FOR = "Select ISNULL(max(zona_edif_orden),0) as Orden from tc_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
        rs_aux1.Open SQL_FOR, db, adOpenKeyset, adLockOptimistic
        If rs_aux1.RecordCount > 0 Then
            Txt_campo1.Text = IIf(IsNull(rs_aux1!Orden), 1, rs_aux1!Orden + 1)
        Else
            Txt_campo1.Text = 1
        End If
        'db.Execute "SELECT Txt_campo1.Text  = ISNULL(MAX(zona_edif_orden),0)+1 FROM tc_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
        db.Execute "insert into  tc_zona_piloto_edif(zpiloto_codigo, edif_codigo, zona_edif_orden, zona_codigo, beneficiario_codigo, beneficiario_codigo_rep, beneficiario_codigo_cobr, zorden_cambio, observaciones, estado_codigo, fecha_registro, usr_codigo, mes_par_impar) " & _
        "values (" & Ado_datos.Recordset!zpiloto_codigo & ", '" & dtc_codigo5.Text & "', '" & Txt_campo1.Text & "', '0', '" & dtc_codigo6.Text & "', '" & dtc_codigo7.Text & "', '" & dtc_codigo8.Text & "', 0, '" & txt_obs.Text & "', 'REG', GETDATE(), 'ADMIN', '" & VAR_IMPAR & "')"
        
        db.Execute "update gc_edificaciones set tomado= 'S' where edif_codigo = '" & dtc_codigo5.Text & "' "
    End If
    If swnuevo = 2 Then
        db.Execute "update tc_zona_piloto_edif set edif_codigo= '" & dtc_codigo5.Text & "', zona_edif_orden='" & Txt_campo1.Text & "', beneficiario_codigo= '" & dtc_codigo6.Text & "', beneficiario_codigo_rep= '" & dtc_codigo7.Text & "',beneficiario_codigo_cobr= '" & dtc_codigo8.Text & "', zorden_cambio= " & cmd_campo2.Text & ", observaciones = '" & txt_obs.Text & "', fecha_registro='" & Date & "' where correlativo=" & Text1.Text & " "
        db.Execute "update tc_zona_piloto_edif set mes_par_impar = '" & VAR_IMPAR & "'  where zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " and edif_codigo = '" & dtc_codigo5.Text & "' "
        If cmd_campo2.Text <> "0" Then
            db.Execute "update tc_zona_piloto_edif set zorden_cambio = zona_edif_orden + 1 where zona_edif_orden >= " & cmd_campo2.Text & " and zona_edif_orden < " & Txt_campo1.Text & " and " & Txt_campo1.Text & " > " & cmd_campo2.Text & " and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & ""
            db.Execute "update tc_zona_piloto_edif set zorden_cambio = zona_edif_orden - 1 where zona_edif_orden <= " & cmd_campo2.Text & " and zona_edif_orden > " & Txt_campo1.Text & " and " & Txt_campo1.Text & " < " & cmd_campo2.Text & " and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & ""
            db.Execute "update tc_zona_piloto_edif set zona_edif_orden = zorden_cambio  where zorden_cambio > '0'  and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & ""
            db.Execute "update tc_zona_piloto_edif set zorden_cambio = '0'  where zorden_cambio > '0'"
        End If
     End If
    '   rs_datos.MoveFirst
    '   mbDataChanged = False
    Call ABRIR_TABLA_DET
    swnuevo = 0
    fra_opciones.Enabled = True
    FraNavega.Enabled = True
    dg_det1.Enabled = True
    Fra_datos.Visible = True
    FraDet2.Visible = False
    
    fra_opciones_det.Visible = True
    
    lbl_orden_camb.Visible = True
    cmd_campo2.Visible = True
    dtc_desc5.Locked = False
    dtc_aux5.Visible = False
    dtc_desc5.Visible = True
     VAR_SW = ""
  End If
'  dtc_desc1.Visible = True
'  lbl_aux1.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub BtnImprimir_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    'Dim co As New ADODB.Command
    CR01.ReportFileName = App.Path & "\Reportes\tecnico\tr_zonas_vs_edificios.rpt"
    CR01.WindowShowPrintSetupBtn = True
    CR01.WindowShowRefreshBtn = True
    'MsgBox rs.RecordCount
    var_titulo = "ZONAS PILOTO"
    VAR_SubTitulo = "TODAS LAS ZONAS"
      CR01.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR01.Formulas(1) = "subtitulo = '" & VAR_SubTitulo & "' "
    ' CR01.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
    CR01.StoredProcParam(0) = Me.Ado_datos.Recordset!zpiloto_codigo
    
    iResult = CR01.PrintReport
    If iResult <> 0 Then MsgBox CR01.LastErrorNumber & " : " & CR01.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR01.WindowState = crptMaximized
End Sub

Private Sub BtnImprimir1_Click()
If Ado_datos.Recordset.RecordCount > 0 Then
    Dim iResult As Integer
    CR02.ReportFileName = App.Path & "\Reportes\tecnico\tr_zonas_vs_edificios_id.rpt"
    CR02.WindowShowPrintSetupBtn = True
    CR02.WindowShowRefreshBtn = True
    var_titulo = "ZONAS PILOTO"
    VAR_SubTitulo = Ado_datos.Recordset!zpiloto_descripcion
      CR02.Formulas(0) = "titulo = '" & var_titulo & "' "
      CR02.Formulas(1) = "subtitulo = '" & VAR_SubTitulo & "' "
    ' CR02.Formulas(1) = "subtitulo = '" & lbl_titulo.Caption & "' "
    CR02.StoredProcParam(0) = Me.Ado_datos.Recordset!zpiloto_codigo
    iResult = CR02.PrintReport
    If iResult <> 0 Then MsgBox CR02.LastErrorNumber & " : " & CR02.LastErrorString, vbCritical, "Error de impresión"
Else
    MsgBox "No se puede Imprimir. Debe registrar los datos correspondientes ...", , "Atención"
End If
    CR02.WindowState = crptMaximized

End Sub

Private Sub BtnModDetalle_Click()
 marca1 = Ado_datos.Recordset.Bookmark
  If Ado_detalle1.Recordset.RecordCount > 0 And Ado_datos.Recordset!estado_codigo = "REG" Then
    swnuevo = 2
    fra_opciones.Enabled = False
    FraNavega.Enabled = False
    dg_det1.Enabled = False
    Fra_datos.Visible = False
    FraDet2.Visible = True
    
    fra_opciones_det.Visible = False
    If Ado_detalle1.Recordset!mes_par_impar = "2" Then
        'PROGRAMAR en Meses PARES y quitar Mes IMPARES
        VAR_IMPAR = "2"
        Option11.Value = True
        Option10.Value = False
    Else
        'Programar Meses IMPARES y quitar PARES
        VAR_IMPAR = "1"
        Option11.Value = False
        Option10.Value = True
    End If
    'Call ABRIR_DET
    VAR_EDIF = Ado_detalle1.Recordset!edif_codigo
    dtc_desc5.BoundText = dtc_codigo5.BoundText
    lbl_orden_camb.Visible = True
    cmd_campo2.Visible = True
    cmd_campo2.Text = "0"
    dtc_codigo5.Locked = True
    dtc_desc5.Locked = True
    dtc_aux5.Visible = True
    dtc_desc5.Visible = False
  Else
    MsgBox "No se puede Modificar el registro, porque este ya está Aprobado!! ", vbExclamation
  End If
End Sub

Private Sub BtnModificar_Click()
On Error GoTo EditErr
'  lblStatus.Caption = "Modificar registro"
    If Ado_datos.Recordset!estado_codigo = "REG" Then
        Fra_datos.Enabled = True
        fra_opciones.Visible = False
        'fra_opciones_det.Visible = False
        FraGrabarCancelar.Visible = True
        FraDet1.Visible = False
        dg_datos.Enabled = False
        VAR_SW = "MOD"
    '    BtnVer.Visible = True
    Else
      MsgBox "No se puede MODIFICAR un registro ya APROBADO ...", vbExclamation, "Validación de Registro"
    End If
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub

Private Sub BtnModificar2_Click()
    Set rs_aux4 = New ADODB.Recordset
    If rs_aux4.State = 1 Then rs_aux4.Close
    rs_aux4.Open "select * from tc_zona_piloto_edif where zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' order by zona_edif_orden ", db, adOpenKeyset, adLockOptimistic, adCmdText
    If rs_aux4.RecordCount > 0 Then
        VAR_CONT = 0
        rs_aux4.MoveFirst
        While Not rs_aux4.EOF
            VAR_CONT = VAR_CONT + 1
            rs_aux4!zorden_cambio = VAR_CONT
            rs_aux4.Update
            rs_aux4.MoveNext
        Wend
        db.Execute "UPDATE tc_zona_piloto_edif SET zona_edif_orden = zorden_cambio WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
        db.Execute "UPDATE tc_zona_piloto_edif SET zorden_cambio ='0' WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' "
        Call ABRIR_TABLA_DET
        MsgBox "Se recodificó la columna ORDEN, satisfactoriamente ...", vbInformation, "Información"
    Else
        MsgBox "No Existen Registros para Ordenar ...", vbExclamation, "Información"
    End If
    
End Sub

Private Sub BtnSalir_Click()
    Unload Me
End Sub

Private Sub dtc_codigo10_Click(Area As Integer)
    dtc_desc10.BoundText = dtc_codigo10.BoundText
End Sub

Private Sub dtc_codigo2_Click(Area As Integer)
    dtc_desc2.BoundText = dtc_codigo2.BoundText
End Sub

Private Sub dtc_codigo5_Click(Area As Integer)
    dtc_desc5.BoundText = dtc_codigo5.BoundText
End Sub

Private Sub dtc_codigo6_Click(Area As Integer)
    dtc_desc6.BoundText = dtc_codigo6.BoundText
End Sub

Private Sub dtc_codigo7_Click(Area As Integer)
    dtc_desc7.BoundText = dtc_codigo7.BoundText
End Sub

Private Sub dtc_codigo8_Click(Area As Integer)
    dtc_desc8.BoundText = dtc_codigo8.BoundText
End Sub

Private Sub dtc_codigo9_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub dtc_desc10_Click(Area As Integer)
    dtc_codigo10.BoundText = dtc_desc10.BoundText
End Sub

Private Sub dtc_desc2_Click(Area As Integer)
    dtc_codigo2.BoundText = dtc_desc2.BoundText
    Call pnivel2(dtc_codigo2.BoundText)
    dtc_desc3.Enabled = True
End Sub

Private Sub dtc_codigo4_Click(Area As Integer)
    dtc_desc4.BoundText = dtc_codigo4.BoundText
End Sub

Private Sub pnivel2(codigo2 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from gc_municipio where prov_codigo = '" & codigo2 & "'"
   Set dtc_codigo3.RowSource = Nothing
   Set dtc_codigo3.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo3.ReFill
   dtc_codigo3.BoundText = Empty
   
   Set dtc_desc3.RowSource = Nothing
   Set dtc_desc3.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc3.ReFill
   dtc_desc3.BoundText = Empty
End Sub

Private Sub dtc_desc4_Click(Area As Integer)
    dtc_codigo4.BoundText = dtc_desc4.BoundText
End Sub

Private Sub dtc_desc5_Click(Area As Integer)
    VAR_5 = dtc_desc5.Text
    dtc_codigo5.BoundText = dtc_desc5.BoundText
End Sub

Private Sub dtc_desc5_LostFocus()
    dtc_desc5.Text = VAR_5
End Sub

Private Sub dtc_desc6_Click(Area As Integer)
    VAR_6 = dtc_desc6.Text
    dtc_codigo6.BoundText = dtc_desc6.BoundText
End Sub

Private Sub dtc_desc6_LostFocus()
    dtc_desc6.Text = VAR_6
End Sub

Private Sub dtc_desc7_Click(Area As Integer)
    VAR_7 = dtc_desc7.Text
    dtc_codigo7.BoundText = dtc_desc7.BoundText
End Sub

Private Sub dtc_desc7_LostFocus()
    dtc_desc7.Text = VAR_7
End Sub

Private Sub dtc_desc8_Click(Area As Integer)
    VAR_8 = dtc_desc8.Text
    dtc_codigo8.BoundText = dtc_desc8.BoundText
End Sub

Private Sub dtc_desc8_LostFocus()
    dtc_desc8.Text = VAR_8
End Sub

Private Sub dtc_desc9_Click(Area As Integer)
    dtc_codigo9.BoundText = dtc_desc9.BoundText
End Sub

Private Sub Form_Load()
    swnuevo = 0
    VAR_SW = ""
    Set rs_aux3 = New ADODB.Recordset
    If rs_aux3.State = 1 Then rs_aux3.Close
    rs_aux3.Open "Select * from gc_usuarios where usr_codigo = '" & glusuario & "' ", db, adOpenStatic
    If rs_aux3.RecordCount > 0 Then
        usuario2 = rs_aux3!beneficiario_codigo
        VAR_DA = rs_aux3!da_codigo
        VAR_DPTO = rs_aux3!depto_codigo
    Else
        usuario2 = "3361040"
        VAR_DA = "1.3"
        VAR_DPTO = "2"
    End If
    VAR_IMPAR = "0"
    VAR_UORIGEN = Aux
    If Aux = "DNMAN" Then
        Select Case VAR_DPTO
            Case "1"    ' Chuquisaca
                VAR_UORIGEN = "DMANC"
            Case "2"    'La Paz - Tecnico
                VAR_UORIGEN = "DNMAN"
            Case "3"    'Cochabamba
                VAR_UORIGEN = "DMANB"
                'VAR_DPTOC = "3"
            Case "4"    'Oruro - Tecnico
                VAR_UORIGEN = "DNMAN"
                'VAR_DPTOC = "2"
            Case "5"    ' Potosi
                VAR_UORIGEN = "DMANC"
            Case "6"    ' Tarija
                VAR_UORIGEN = "DMANC"
            Case "7"    'Santa Cruz
                VAR_UORIGEN = "DMANS"
                'VAR_DPTOC = "7"
            Case "8"    ' Beni
                VAR_UORIGEN = "DMANS"
            Case "9"    ' Pando
                VAR_UORIGEN = "DMANS"
            Case Else    ' TODO
                VAR_UORIGEN = "DNMAN"
                VAR_DPTO = "0"
         End Select
    End If
    
    If Aux = "DNINS" Then
        Select Case VAR_DA
            Case "1.8"    'Cochabamba
                Aux = "DINSB"
                VAR_DPTO = "3"
            Case "1.7"    'Santa Cruz
                Aux = "DINSS"
                VAR_DPTO = "7"
            Case "1.3"    'La Paz - Tecnico
                Aux = "DNINS"
                VAR_DPTO = "2"
            Case "1.9"    ' Chuquisaca
                Aux = "DINSC"
                VAR_DPTO = "1"
            Case Else    ' TODO
                Aux = "DNINS"
                VAR_DPTO = "2"
         End Select
    End If

    'Fra_Gestion.Visible = True
'    VAR_GES = Cmb_gestion.Text
    parametro = Aux
    'ACTUALIZA ESTADO DE USO DE EDIFICIOS EN UNA ZONA PILOTO
    db.Execute "update tc_zona_piloto_edif SET estado_activo = 'REG', estado_codigo = 'REG' "
    db.Execute "update tc_zona_piloto_edif SET estado_activo = 'APR' FROM tc_zona_piloto_edif INNER JOIN tv_crono_diario_agrupado_por_edificio ON tc_zona_piloto_edif.edif_codigo = tv_crono_diario_agrupado_por_edificio.edif_codigo "
    db.Execute "update tc_zona_piloto_edif SET estado_codigo = 'APR' FROM tc_zona_piloto_edif INNER JOIN tv_crono_FINAL_agrupado_por_edificio ON tc_zona_piloto_edif.edif_codigo = tv_crono_FINAL_agrupado_por_edificio.edif_codigo "

    'Actualiza Edificios Tomados en Organizacion de Zonas
    db.Execute "update gc_edificaciones set tomado = 'N' "
    db.Execute "update gc_edificaciones set gc_edificaciones.tomado= 'S' from gc_edificaciones inner join tc_zona_piloto_edif on gc_edificaciones.edif_codigo = tc_zona_piloto_edif.edif_codigo "
    
    db.Execute "update tc_zona_piloto_edif SET ges_gestion = '2022' WHERE ges_gestion IS NULL "
    'db.Execute "UPDATE tc_zona_piloto_edif SET tc_zona_piloto_edif.unimed_codigo  = to_cronograma.unimed_codigo FROM tc_zona_piloto_edif INNER JOIN to_cronograma ON to_cronograma.edif_codigo  =tc_zona_piloto_edif.edif_codigo WHERE (tc_zona_piloto_edif.unimed_codigo IS NULL OR tc_zona_piloto_edif.unimed_codigo = 'EQP' ) "
    'db.Execute "UPDATE tc_zona_piloto_edif SET tc_zona_piloto_edif.codigo_empresa = to_cronograma.codigo_empresa FROM tc_zona_piloto_edif INNER JOIN to_cronograma ON to_cronograma.edif_codigo  =tc_zona_piloto_edif.edif_codigo where (tc_zona_piloto_edif.codigo_empresa Is Null OR tc_zona_piloto_edif.codigo_empresa ='0' ) "
    'db.Execute "UPDATE tc_zona_piloto_edif SET tc_zona_piloto_edif.solicitud_tipo  = to_cronograma.solicitud_tipo FROM tc_zona_piloto_edif INNER JOIN to_cronograma ON to_cronograma.edif_codigo  =tc_zona_piloto_edif.edif_codigo where (tc_zona_piloto_edif.solicitud_tipo Is Null OR (tc_zona_piloto_edif.solicitud_tipo <> '10' AND tc_zona_piloto_edif.solicitud_tipo <> '6' )) "

    'IDENTIDICA VENTAS NUEVAS QUE YA TIENEN CONTRATO DE MANTENIMIENTO
    db.Execute "UPDATE ao_ventas_cabecera SET ao_ventas_cabecera.correl_detalle ='1' FROM ao_ventas_cabecera INNER JOIN av_ventas_cabecera_mant ON ao_ventas_cabecera.edif_codigo = av_ventas_cabecera_mant.edif_codigo WHERE ao_ventas_cabecera.unidad_codigo ='DVTA' OR ao_ventas_cabecera.unidad_codigo ='DCOMC' OR ao_ventas_cabecera.unidad_codigo ='DCOMS' OR ao_ventas_cabecera.unidad_codigo ='DCOMB' "
    
    Call ABRIR_TABLAS_AUX
    Call OptFilGral2_Click
    
    Fra_datos.Enabled = False
    dg_datos.Enabled = True
    

        Call SeguridadSet(Me)
End Sub

Private Sub ABRIR_TABLAS_AUX()
    'gc_departamento
    Set rs_datos1 = New ADODB.Recordset
    If rs_datos1.State = 1 Then rs_datos1.Close
    rs_datos1.Open "Select * from gc_departamento order by depto_codigo ", db, adOpenStatic
    Set Ado_datos1.Recordset = rs_datos1
    dtc_desc1.BoundText = dtc_codigo1.BoundText

    'gc_provincia
    Set rs_datos2 = New ADODB.Recordset
    If rs_datos2.State = 1 Then rs_datos2.Close
    rs_datos2.Open "Select * from gc_provincia order by prov_descripcion", db, adOpenStatic
    Set Ado_datos2.Recordset = rs_datos2
    dtc_desc2.BoundText = dtc_codigo2.BoundText

    'gc_municipio
    Set rs_datos3 = New ADODB.Recordset
    If rs_datos3.State = 1 Then rs_datos3.Close
    rs_datos3.Open "Select * from gc_municipio where region_codigo = 'SI' order by munic_descripcion", db, adOpenStatic
    Set Ado_datos3.Recordset = rs_datos3
    dtc_desc3.BoundText = dtc_codigo3.BoundText

    'gc_zonas
    Set rs_datos9 = New ADODB.Recordset
    If rs_datos9.State = 1 Then rs_datos3.Close
    rs_datos9.Open "Select * from gc_zonas order by zona_denominacion", db, adOpenStatic
    Set Ado_datos9.Recordset = rs_datos9
    dtc_desc9.BoundText = dtc_codigo9.BoundText

    'Beneficiario Funcionario CGI (Tecnico Responsable)
    Set rs_datos4 = New ADODB.Recordset
    If rs_datos4.State = 1 Then rs_datos4.Close
    rs_datos4.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos4.Recordset = rs_datos4
    dtc_desc4.BoundText = dtc_codigo4.BoundText
    
    'gc_edificaciones
    Set rs_datos5 = New ADODB.Recordset
    If rs_datos5.State = 1 Then rs_datos5.Close
    rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
    Set Ado_datos5.Recordset = rs_datos5
    dtc_desc5.BoundText = dtc_codigo5.BoundText

    'Beneficiario Funcionario CGI (Tecnico Mantenimiento)
    Set rs_datos6 = New ADODB.Recordset
    If rs_datos6.State = 1 Then rs_datos6.Close
    rs_datos6.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos6.Recordset = rs_datos6
    dtc_desc6.BoundText = dtc_codigo6.BoundText
    
    'Beneficiario Funcionario CGI (Tecnico Instaciones)
    Set rs_datos7 = New ADODB.Recordset
    If rs_datos7.State = 1 Then rs_datos7.Close
    rs_datos7.Open "rv_unidad_vs_responsable where unidad_codigo = '" & parametro & "' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos7.Recordset = rs_datos7
    dtc_desc7.BoundText = dtc_codigo7.BoundText
    
    'Beneficiario Funcionario CGI (Cobrador)
    Set rs_datos8 = New ADODB.Recordset
    If rs_datos8.State = 1 Then rs_datos8.Close
    rs_datos8.Open "rv_unidad_vs_responsable where unidad_codigo = 'DCOBR' ORDER BY beneficiario_denominacion ", db, adOpenStatic
    Set Ado_datos8.Recordset = rs_datos8
    dtc_desc8.BoundText = dtc_codigo8.BoundText
    
End Sub

Private Sub dtc_codigo1_Click(Area As Integer)
    dtc_desc1.BoundText = dtc_codigo1.BoundText
End Sub

Private Sub dtc_codigo3_Click(Area As Integer)
    dtc_desc3.BoundText = dtc_codigo3.BoundText
End Sub

Private Sub dtc_desc1_Click(Area As Integer)
    dtc_codigo1.BoundText = dtc_desc1.BoundText
    Call pnivel1(dtc_codigo1.BoundText)
    dtc_desc2.Enabled = True
End Sub

Private Sub pnivel1(codigo1 As String)
   Dim strConsultaF As String
   strConsultaF = "select * from gc_provincia where depto_codigo = '" & codigo1 & "'"

   Set dtc_codigo2.RowSource = Nothing
   Set dtc_codigo2.RowSource = db.Execute(strConsultaF, , adCmdText)
'   Set dtc_codigo10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_codigo2.ReFill
   dtc_codigo2.BoundText = Empty

   Set dtc_desc2.RowSource = Nothing
   Set dtc_desc2.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc10.RowSource = db.Execute(" EXEC pp_listar_mediante_padre_pc_poa_actividad '" & codigo1 & "' ")
   dtc_desc2.ReFill
   dtc_desc2.BoundText = Empty
End Sub

Private Sub dtc_desc3_Click(Area As Integer)
    dtc_codigo3.BoundText = dtc_desc3.BoundText
    'Call pnivel5(dtc_codigo3.BoundText)
    'dtc_desc9.Enabled = True
End Sub
   
Private Sub pnivel5(codigo7 As String)
   Dim strConsultaF As String
   
   strConsultaF = "select * from gc_zonas where munic_codigo = '" & codigo7 & "' order by zona_denominacion"
   Set dtc_codigo9.RowSource = Nothing
   Set dtc_codigo9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_codigo4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_codigo9.ReFill
   dtc_codigo9.BoundText = Empty
   
   Set dtc_desc9.RowSource = Nothing
   Set dtc_desc9.RowSource = db.Execute(strConsultaF, , adCmdText)
   'Set dtc_desc4.RowSource = db.Execute(" EXEC gp_listar_mediante_padre_gc_calles '" & codigo3 & "' ")
   dtc_desc9.ReFill
   dtc_desc9.BoundText = Empty
End Sub


Private Sub OptFilGral1_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If VAR_UORIGEN = "DNINS" Then
        queryinicial = "Select * from tc_zonas_piloto WHERE zpiloto_codigo = '0' "
    Else
        Select Case VAR_DPTO
           Case "1"    ' Chuquisaca
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '5') "
           Case "2"    'La Paz - Tecnico
               If glusuario = "ADMIN" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "JSAAVEDRA" Or glusuario = "CSALINAS" Then
                    queryinicial = "Select * from tc_zonas_piloto  "
               Else
                    queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
               End If
           Case "3"    'Cochabamba
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "7"    'Santa Cruz
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '1' OR depto_codigo = '8') "
           Case "4"    'Oruro - Tecnico
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "5"    ' Potosi
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "6"    ' Tarija
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "8"    ' Beni
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "9"    ' Pando
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case Else    ' TODO
               queryinicial = "select * From tc_zonas_piloto  "     'tv_cronograma_edificaciones
        End Select

'        If VAR_DPTO = "7" Then
'            queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '1') "
'        Else
'            queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'        End If
    End If
    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
    Set Ado_datos.Recordset = rs_datos.DataSource
    Set dg_datos.DataSource = Ado_datos.Recordset
End Sub

'Private Sub OptFilGral1_Click()
'    '===== Proceso para filtrado general de datos (registros no aprobados)
'    Set rs_datos = New Recordset
'    If rs_datos.State = 1 Then rs_datos.Close
'    'queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '" & VAR_GES & "' "
'    queryinicial = "select * From to_cronograma_mensual WHERE estado_codigo = 'REG' AND unidad_codigo_tec = '" & parametro & "' AND ges_gestion = '2015' "
'    'queryinicial = "Select * from ao_solicitud where " + parametro
'    rs_datos.Open queryinicial, db, adOpenKeyset, adLockOptimistic
'    Set Ado_datos.Recordset = rs_datos.DataSource
'    Set dg_datos.DataSource = Ado_datos.Recordset
'End Sub

Private Sub OptFilGral2_Click()
    '===== Proceso para filtrado general de datos (todos los registros)
    Set rs_datos = New Recordset
    If rs_datos.State = 1 Then rs_datos.Close
    If VAR_UORIGEN = "DNINS" Then
        queryinicial = "Select * from tc_zonas_piloto WHERE zpiloto_codigo = '0' "
    Else
        Select Case VAR_DPTO
           Case "1"    ' Chuquisaca
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '5') "
           Case "2"    'La Paz - Tecnico
               If glusuario = "ADMIN" Or glusuario = "OCOLODRO" Or glusuario = "JORAQUENI" Or glusuario = "LNAVA" Or glusuario = "JSAAVEDRA" Or glusuario = "CSALINAS" Then
                    queryinicial = "Select * from tc_zonas_piloto  "
               Else
                    queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
               End If
           Case "3"    'Cochabamba
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "7"    'Santa Cruz
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '1' OR depto_codigo = '8') "
           Case "4"    'Oruro - Tecnico
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "5"    ' Potosi
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "6"    ' Tarija
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "8"    ' Beni
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case "9"    ' Pando
               queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
           Case Else    ' TODO
               queryinicial = "select * From tc_zonas_piloto  "     'tv_cronograma_edificaciones
        End Select

'        If VAR_DPTO = "7" Then
'            queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "' OR depto_codigo = '1') "
'        Else
'            queryinicial = "Select * from tc_zonas_piloto WHERE (depto_codigo = '" & VAR_DPTO & "') "
'        End If
    End If
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
    Set rs_det1 = New ADODB.Recordset
    If rs_det1.State = 1 Then rs_det1.Close
    'If VAR_UORIGEN = "DNINS" Then
    '    rs_det1.Open "select * from tv_zona_piloto_edif where zpiloto_codigo = '0' order by zona_edif_orden ", db, adOpenKeyset, adLockOptimistic, adCmdText
    'Else
        rs_det1.Open "select * from tv_zona_piloto_edif where zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " order by zona_edif_orden ", db, adOpenKeyset, adLockOptimistic, adCmdText
    '    ' and ges_gestion = '" & g & "'
    'End If
    Set Ado_detalle1.Recordset = rs_det1
    If Ado_detalle1.Recordset.RecordCount > 0 Then
'        'gc_edificaciones
'        Set rs_aux5 = New ADODB.Recordset
'        If rs_aux5.State = 1 Then rs_aux5.Close
'        rs_aux5.Open "Select * from AV_VENTAS_FECHA_MAX WHERE zpiloto_codigo = '" & Ado_datos.Recordset!zpiloto_codigo & "' ", db, adOpenStatic
'        If rs_aux5.RecordCount > 0 Then
'            rs_aux5.MoveFirst
'            While Not rs_aux5.EOF
'                Set rs_aux6 = New ADODB.Recordset
'                If rs_aux6.State = 1 Then rs_aux6.Close
'                rs_aux6.Open "Select * from ao_ventas_cabecera where venta_fecha_fin = '" & rs_aux5!venta_fecha_fin & "' and edif_codigo = '" & rs_aux5!EDIF_CODIGO & "' and zpiloto_codigo = " & Ado_datos.Recordset!zpiloto_codigo & " AND estado_codigo = 'APR' ", db, adOpenStatic
'                If rs_aux6.RecordCount > 0 Then
'                    db.Execute "UPDATE tc_zona_piloto_edif SET codigo_empresa= " & rs_aux6!codigo_empresa & ", unimed_codigo = '" & IIf(IsNull(rs_aux6!unimed_codigo_tec), "MES", rs_aux6!unimed_codigo_tec) & "', solicitud_tipo = " & rs_aux5!solicitud_tipo & ", fecha_fin_max = '" & rs_aux5!venta_fecha_fin & "', Gratuito = 'NO' WHERE edif_codigo = '" & rs_aux6!EDIF_CODIGO & "'  "
'                End If
'                rs_aux5.MoveNext
'            Wend
'        End If
    End If
    Set dg_det1.DataSource = Ado_detalle1.Recordset
    If swnuevo = 0 Then
        'gc_edificaciones
        Set rs_datos5 = New ADODB.Recordset
        If rs_datos5.State = 1 Then rs_datos5.Close
        rs_datos5.Open "Select * from gc_edificaciones where estado_codigo = 'APR' order by edif_descripcion", db, adOpenStatic
        Set Ado_datos5.Recordset = rs_datos5
        dtc_desc5.BoundText = dtc_codigo5.BoundText
    End If
    
End Sub

Private Function ExisteReg(Codigo As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'estado_codigo = 'APR' and
    GlSqlAux = "SELECT Count(*) AS Cuantos FROM to_cronograma WHERE zpiloto_codigo = '" & Codigo & "'"
    rs.Open GlSqlAux, db, adOpenStatic
    ExisteReg = rs!Cuantos > 0
End Function

